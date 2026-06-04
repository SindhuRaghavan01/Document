import json
import os
import pickle
from multiprocessing import Pool

import numpy as np
from numpy.random import rand

from src.GeneticAlgorithm import GA
from utils.db_utils import create_db_table, get_data, push_to_db
from utils.helpers import (fitness_array_to_dict,
                           get_layer_ylts, save_optimization_output)
from utils.logger import Logger


class Crown_GA(GA):
    """Crown specific optimization
    """

    def __init__(self, optimization_id, connection_config, queries, output_dir=None, defaults={}, logger=None):
        """Crown specific optimization
        """
        self.optimization_id = optimization_id
        self.logger = logger
        # # self.logger.basicConfig(filename=self.logfilepath, level=LOG_LEVEL, format='%(asctime)s :%(levelname)s: %(message)s')  # , datefmt='%m/%d/%Y %I:%M:%S %p')
        self.logger.info("Initializing Optimization")
        self.connection_config = connection_config
        self.queries = queries
        self.base_table_name = f'ARAPL_Optimization.dbo.Optimization_{optimization_id}'
        if not output_dir:
            output_dir = f"./output/Optimization_{optimization_id}"
            # for subdir in ("population", "fitnesses"):
        os.makedirs(output_dir, exist_ok=True)
        self.output_dir = output_dir
        self.output_progress_file = os.path.join(output_dir, "progress.csv")
        self.output_generations_file  = os.path.join(output_dir, "Iter_{iteration}_shares.csv")
        self.output_objectives_file = os.path.join(output_dir, "Iter_{iteration}_objectives.json")
        self.output_constraints_file  = os.path.join(output_dir, "Iter_{iteration}_constraints.json")
        # self.output_population_dir = os.path.join(output_dir, "population")
        # self.output_fitness_dir = os.path.join(output_dir, "fitnesses")
        self.setup_optimization(defaults)
        self.logger.info("Retrieving layer YLTs")
        self.collect_base_layer_data()
        self.initial_generation = self.get_used_share(self.initial_population)
        self.initial_fitness = self.get_fitness(self.initial_generation)
        self.logger.info("Initialized Optimization")
        self.logger.info("Optimization ready to run")

    def setup_optimization(self, defaults):
        """setup optimization parameters and initialize the optimization.
        If the parameter values are not provided in the database, they will take be taken from config/defaults.json
        """
        # self.logger.info("setting up optimization parameters")

        # input = {layer_id: {ILSName: "abcd", "ILSInstrumentID": 1234, "Min_Share": 0, "Max_Share": 1}}
        optimization_input, input_columns = get_data(self.queries["optimization_input_query"].format(
            OptimizationID=self.optimization_id), self.connection_config, as_type="dict", do_enumerate=True)
        # self.seeded = False
        if "Participation" in input_columns:
            self.seeded = True
        if not optimization_input:
            raise
        # objectives = {objective_id: {"type": "min", "field": "AUM", layer_group: "CATBond"}}
        objectives, _ = get_data(self.queries["optimization_objectives_query"].format(
            OptimizationID=self.optimization_id), self.connection_config, as_type="dict", do_enumerate=True, key_column="ObjConsID")
        # constraints = {constraint_id: {"type": "min", "field": "AUM", layer_group: "CATBond", "value": 50e6}}
        constraints, _ = get_data(self.queries["optimization_constraints_query"].format(
            OptimizationID=self.optimization_id), self.connection_config, as_type="dict", do_enumerate=True, key_column="ObjConsID")

        for constraint_id in constraints:
            constraint = constraints[constraint_id]
            constraint_type = constraint["Types"]
            constraint_min = constraint["Min"]
            constraint_max = constraint["Max"]
            target = constraint_min if constraint_max == 0 else constraint_max
            if not constraint_type or constraint_type == '0':
                if constraint_max == 0:
                    constraint_type = "min"
                elif constraint_min == 0:
                    constraint_type = "max"
                else:
                    constraint_type = "minmax"
            if constraint_type == "min":
                constraints[constraint_id]["target"] = target
            elif constraint_type == "max":
                constraints[constraint_id]["target"] = -target
            constraints[constraint_id]["Types"] = constraint_type
        # parameters = from `default.json`
        optimization_parameters, _ = get_data(self.queries["optimization_parameters_query"].format(
            OptimizationID=self.optimization_id), self.connection_config, as_type="dict", do_enumerate=False)
        optimization_parameters = optimization_parameters[0]
        param_default = defaults["parameters"]
        optimization_parameters["Solver"] = optimization_parameters.get(
            "Solver", param_default["solver"])
        optimization_parameters["n_bits"] = optimization_parameters.get(
            "n_bits", param_default["n_bits"])
        optimization_parameters["mutation_probability"] = optimization_parameters.get(
            "MutationRate", param_default["mutation_probability"])
        optimization_parameters["population_size"] = optimization_parameters.get(
            "PopulationSize", param_default["population_size"])
        optimization_parameters["crossover_method"] = optimization_parameters.get(
            "crossover_method", param_default["crossover_method"])
        optimization_parameters["crossover_eta_c"] = optimization_parameters.get(
            "crossover_eta_c", param_default["crossover_eta_c"])
        optimization_parameters["mutation_method"] = optimization_parameters.get(
            "mutation_method", param_default["mutation_method"])
        optimization_parameters["mutation_eta_m"] = optimization_parameters.get(
            "mutation_eta_m", param_default["mutation_eta_m"])
        optimization_parameters["iterations"] = optimization_parameters.get(
            "Iterations", param_default["iterations"])
        # self.logger.info("Setting up Optimization")

        self._ids = list(optimization_input.keys())
        self.n_layers = len(self._ids)
        self.input_data = optimization_input
        self.population_size = optimization_parameters["population_size"]
        self.initial_shares = np.zeros(self.n_layers)
        self.min_shares = np.zeros(self.n_layers)
        self.max_shares = np.zeros(self.n_layers)
        self.enable_exclude = np.zeros(self.n_layers)
        self.premiums = np.zeros(self.n_layers)
        self.profit_commissions = np.zeros(self.n_layers)
        self.initial_fees = np.zeros(self.n_layers)
        self.proportional_expenses = np.zeros(self.n_layers)
        self.hurdles = np.zeros(self.n_layers)
        self.initial_population = {"shares": np.zeros((self.population_size, self.n_layers)),
                                   "includes": np.ones((self.population_size, self.n_layers)),
                                   }

        # self.seeded = False
        for i in self.input_data:

            self.min_shares[i] = self.input_data[i]["Min_Share"]
            self.max_shares[i] = self.input_data[i]["Max_Share"]
            if not self.seeded:
                # initial_share = min_share + (max_share - min_share) * rand()
                u = rand()
                r = np.random.choice([0, 1], p=[0.3, 0.7])
                self.input_data[i]["initial_share"] = self.min_shares[i] + \
                    (self.max_shares[i] - self.min_shares[i]) * u * r
            else:
                initial_share = self.input_data[i]["Participation"]
                self.input_data[i]["initial_share"] = initial_share
            # used_share = self.get_used_share(initial_share, min_share, max_share, enable_exclude)
            # self.input_data[i]["used_share"] = used_share

            self.initial_shares[i] = self.input_data[i]["initial_share"]

            self.enable_exclude[i] = 1 if self.input_data[i]["Enable_Exclude"] in ('1', 'TRUE') else 0
            self.premiums[i] = self.input_data[i]["Premium"]
            # Initial_Fees, ProfitCommission, Hurdle, ["Reinstatement", "ManagementFees", "FET", "Setup", "Origination"]
            self.initial_fees[i] = self.input_data[i]["Initial_Fees"]
            self.profit_commissions[i] = self.input_data[i]["ProfitCommission"]
            self.hurdles[i] = self.input_data[i]["Hurdle"]
            self.proportional_expenses[i] = sum([self.input_data[i][j] for j in [
                                                "Reinstatement", "ManagementFees", "FET", "Setup", "Origination", "UWOverride"]])

            shares_i = (self.initial_shares[i] - self.min_shares[i]) / (
                self.max_shares[i] - self.min_shares[i]) if self.max_shares[i] != self.min_shares[i] else 0
            shares_i = max(min(shares_i, 1), 0)
            u = np.random.rand()
            self.initial_population["shares"][:, i] = np.random.choice(
                [u, shares_i], size=self.population_size, replace=True, p=[0.1, 0.9])
            # shares_i if r > 0.5 else np.random.rand()
            # self.initial_population["includes"][:, i] = np.sign(self.initial_shares[i])
            
            # is_included = 1 if self.initial_shares[i] != 0 else 0
            # if is_included:
            #     self.initial_population["includes"][:, i] = np.random.choice([0, 1], size=self.population_size, replace=True, p=[0.2, 0.8])
            # else:
            #     self.initial_population["includes"][:, i] = np.random.choice([0, 1], size=self.population_size, replace=True, p=[0.8, 0.2])
            self.initial_population["includes"][:, i] = np.random.choice([0, 1], size=self.population_size, replace=True, p=[0.9, 0.1])
            
            # else:
            #     self.initial_population["shares"][:, i] = np.random.rand(self.population_size)
            #     self.initial_population["includes"][:, i] = np.random.choice([0, 1], size=self.population_size, replace=True, p=[0.5, 0.5])

        # self.initial_population = self.get_used_share(self.mutate(self.initial_population, mutation_probability=0.5))

        super(Crown_GA, self).__init__(optimization_input, objectives,
                                       constraints, optimization_parameters, seeded=self.seeded)
        

    def get_used_share(self, population):
        """ base GA gives values between 0,1. Convert the values to shares for use in fitness calculations
        """
        population["includes"][:, self.enable_exclude == 0] = 1
        return np.round(population["includes"]) * (self.min_shares + population["shares"] * (self.max_shares - self.min_shares))

    def get_layer_ylt(self, instrument_id, loss_filter, calculate_expenses=True):
        """get layer ylt with expenses
        loss_filter = None or {'Peril': 'TC', 'Region': '', 'SubRegion': '', 'Country': 'USA'}
        """
        crown_activity_id = self.input_data[instrument_id]["CrownActivityID"]
        peril, region, subregion, country = '', '', '', ''
        # get_ylt_query = self.queries["get_ylt_query"]
        self.logger.debug(f"Getting layer ylt for Crown Activity {crown_activity_id}")
        if loss_filter:
            peril = loss_filter["Peril"]
            region = loss_filter["Region"]
            subregion = loss_filter["SubRegion"]
            country = loss_filter["Country"]
        ylt_query = self.queries["get_ylt_query"].format(
            peril=peril, region=region, subregion=subregion, country=country, crown_activity_id=crown_activity_id)
        ylt, _ = get_data(ylt_query, self.connection_config,
                          as_type="dict", key_column="Year")
        # self.logger.debug("pad ylt where loss = 0 and calculate expenses")
        if not ylt:
            max_rank = 0
            min_tvar = 0
            aal = 0
            return None
        else:
            max_rank = max([ylt[i]['OrderRank'] for i in ylt])
            min_tvar = [i for i in ylt if ylt[i]["OrderRank"] == max_rank][0]
            aal = min_tvar * max_rank / 10000.0

        if calculate_expenses:
            initial_fees = self.input_data[instrument_id]["Initial_Fees"]
            proportional_expenses = sum([self.input_data[instrument_id][i] for i in [
                                        "Reinstatement", "ManagementFees", "FET", "Setup", "Origination"]])
            profit_commission = self.input_data[instrument_id]["ProfitCommission"]
            hurdle = self.input_data[instrument_id]["Hurdle"]
            premium = self.input_data[instrument_id]["Premium"]

        for i in range(1, 10001):
            max_rank += 1
            if i not in ylt:
                ylt[i] = {"Loss": 0, "TVAR": aal, "OrderRank": max_rank,
                          "ReturnPeriod": 10000.0 / max_rank}
            if calculate_expenses:
                profit_net_expenses = premium * \
                    (1 - proportional_expenses - initial_fees) - ylt[i]["Loss"]
                if profit_net_expenses > premium * hurdle:
                    profit_net_expenses -= premium * profit_commission
                ylt[i]["Return"] = profit_net_expenses
                ylt[i]["LossNetExpenses"] = premium - profit_net_expenses
        return ylt

    def collect_base_layer_data(self):
        """collect all layer ylts with and without expenses
        set AUM by layer, returns YLTs by layer, loss YLTs by loss filter and layer
        self._AUM : numpy array (n_layers, )
        self._layer_groups : numpy array (n_layers, )
        self._filtered_layer_ylt_returns : numpy array (n_layers, 10000)
        self._filtered_layer_ylt_losses : {filter_name: numpy array (n_layers, 10000)}
        """
        filtered_ylts = {}
        loss_filter_list = [self.constraints[i]['LossFilter']
                            for i in self.constraints if self.constraints[i]["LossFilter"] != '0']
        loss_filters = {}
        if loss_filter_list:
            loss_filter_quoted_list = ",".join(
                [f"'{i}'" for i in loss_filter_list])
            # loss_filter = {'Peril': 'TC', 'Region': '', 'SubRegion': '', 'Country': 'USA'}
            loss_filters, _ = get_data(self.queries["loss_filter_peril_query"].format(
                loss_filter_quoted_list=loss_filter_quoted_list), self.connection_config, do_enumerate=True, key_column="LossFilter")
        loss_filters["all"] = {'Peril': '',
                               'Region': '', 'SubRegion': '', 'Country': ''}
        for filter_name in loss_filters:
            layer_ylts = {}
            loss_filter = loss_filters[filter_name]
            # self.logger.debug(f"get ylts for {filter_name}")
            ylt_queries = [self.queries["get_ylt_query"].format(peril=loss_filter["Peril"], region=loss_filter["Region"], subregion=loss_filter["SubRegion"],
                                                                country=loss_filter["Country"], crown_activity_id=self.input_data[i]["CrownActivityID"]) for i in self.input_data]
            layer_ylts = get_layer_ylts(ylt_queries, self.connection_config)

            filtered_ylts[filter_name] = layer_ylts
        # get expenses
        premium = np.tile(self.premiums, (10000, 1)).T
        initial_fees = np.tile(self.initial_fees, (10000, 1)).T
        profit_commission = np.tile(self.profit_commissions, (10000, 1)).T
        hurdle = np.tile(self.hurdles, (10000, 1)).T
        proportional_expenses = np.tile(
            self.proportional_expenses, (10000, 1)).T
        # calculate returns
        # return_ylt = premium * (1 - proportional_expenses - initial_fees) - filtered_ylts["all"]
        # Analyze Re ignores initial fees
        return_ylt = premium * (1 - proportional_expenses) - filtered_ylts["all"]
        # Profit Commission calculation
        # profit_1 = premium - loss - proportional expenses
        # if profit_1 > hurdle then profit commission = PC * (profit_1 - hurdle)
        # we are using hard hurdle here
        hurdle_filter = return_ylt > premium * (hurdle)
        return_ylt[hurdle_filter] = return_ylt[hurdle_filter] - \
            profit_commission[hurdle_filter] * (return_ylt[hurdle_filter] - hurdle[hurdle_filter] * premium[hurdle_filter])
        self._field_data = {}
        for field in ["AUM", "Premium", "Qscore"]:
            self._field_data[field] = np.array(
                [self.input_data[i][field] for i in self.input_data])
        self._layer_groups = np.array(
            [self.input_data[i]["layer_group"] for i in self.input_data])
        self._layer_returns_ylt = return_ylt
        self._filtered_layer_losses_ylt = filtered_ylts

    def _get_generation_array(self, generation):
        """convert the population object to numpy array
        returns: numpy array (population_size, n_layers)
        """
        if isinstance(generation, np.ndarray):
            return generation.copy()
        return np.array([[generation[solution_id][layer_id] for layer_id in generation[solution_id]] for solution_id in generation])

    def get_generation_data(self, generation):
        """get population ylt
        calls get_solution_ylt for solution_ids in the population
        """
        loss_filters = [self.objectives[i]["LossFilter"] for i in self.objectives if self.objectives[i]["LossFilter"] !=
                        '0'] + [self.constraints[i]["LossFilter"] for i in self.constraints if self.constraints[i]["LossFilter"] != '0']
        loss_filters = list(set(loss_filters))
        loss_filters.append("all")
        pop_array = self._get_generation_array(generation)
        pop_ylts = {}
        # (population_size, 10000) = (population_size, n_layers) x (n_layers, 10000)
        pop_returns = np.dot(pop_array, self._layer_returns_ylt)

        for filter_name in loss_filters:
            ylt = np.dot(
                pop_array, self._filtered_layer_losses_ylt[filter_name])
            pop_ylts[filter_name] = ylt
        return pop_array, pop_returns, pop_ylts

    def criterion_value(self, generation_array, generation_returns, generation_ylts, criterion_object):
        """evaluates a criterion and return the value for the criterion for the entire population
        criterion_value : numpy array (population_size, )
        """
        gen_size = generation_array.shape[0]
        functions = criterion_object["Functions"]
        return_period = criterion_object["Return_Period"]
        loss_filter = criterion_object["LossFilter"]
        loss_filter = 'all' if loss_filter == '0' else loss_filter
        layer_groups = criterion_object["Layer_Groups"]
        loss_perspective = criterion_object.get("Loss_Perspective", "")
        if layer_groups == '0':
            layer_groups = None
            layer_group_filter = np.ones(self.n_layers, dtype=np.bool)
        else:
            layer_groups = [i.strip() for i in layer_groups.split(",")]
            layer_group_filter = np.isin(self._layer_groups, layer_groups)
        fields = [criterion_object["Fields"]]
        if criterion_object["ProportionateTo"] != '0':
            fields.append(criterion_object["ProportionateTo"])

        if functions == 'aggregate_terms':
            # calculate agg terms using fields. use data from self.input

            value = np.ones(self.n_layers)
            if layer_group_filter.size:
                value[~layer_group_filter] = 0
            for field in fields:
                value *= self._field_data[field]
            criterion_value = np.sum(
                generation_array * np.tile(value, (gen_size, 1)), axis=1)

        elif functions == 'tail_value_at_risk':
            # calculate TVAR using loss filter at given return period
            rank = np.int(10000 * return_period)
            if loss_perspective == "losses_net_premium":
                criterion_value = -np.mean(np.sort(generation_returns, axis=1)[:, :rank], axis=1)
            else:
                criterion_value = np.mean(np.sort(generation_ylts[loss_filter], axis=1)[:, -rank:], axis=1)
            # criterion_value = [solution_ylts[loss_filter]["TVAR"] for i in solution_ylts if solution_ylts[loss_filter]["ReturnPeriod"] == return_period][0]

        elif functions == 'expected_return':
            # calculate expected return using portfolio ylt using loss filter
            # just add the returns and divide by 10000 since the return is calculated in the ylt already
            criterion_value = generation_returns.sum(axis=1) / 1e4

        elif functions == 'value_at_risk':
            rank = np.int(10000 * return_period)
            if loss_perspective == "losses_net_premium":
                criterion_value = -np.sort(generation_returns, axis=1)[:, :rank]
            else:
                criterion_value = np.sort(
                    generation_ylts[loss_filter], axis=1)[:, -rank]
        return criterion_value

    def evaluate_generation(self, generation):
        """evaluate population and return objective and constraint values as required dict

        fitness_objective_values: {solution_id: {violated: 1,"objectives": {obj1: obj_val1, obj2: val2, obj3: val3}, "constraints": {cons1: cons1_value}
        """
        gen_size = generation.shape[0]
        generation_array, generation_returns, generation_ylts = self.get_generation_data(
            generation)
        fitness = {i: {"objectives": {}, "constraints": {}, "violations": {}}
                   for i in range(gen_size)}
        objectives_fitness = {}
        constraints_fitness = {}
        constraints_violation = {}
        is_infeasible = np.zeros(gen_size)
        constraint_count = len(self.constraints)
        constraint_conditions = np.array(
            np.ones((gen_size, constraint_count)), dtype=np.bool)
        for objective_id in self.objectives:
            objective = self.objectives[objective_id]
            objective_value = self.criterion_value(
                generation_array, generation_returns, generation_ylts, objective)
            if self.objectives[objective_id]["Types"].lower() == "max":
                objectives_fitness[objective_id] = -objective_value
            else:
                objectives_fitness[objective_id] = objective_value
        for i, constraint_id in enumerate(self.constraints):
            constraint = self.constraints[constraint_id]
            constraint_value = self.criterion_value(
                generation_array, generation_returns, generation_ylts, constraint)
            constraint_type = constraint["Types"].lower()
            constraint_min = constraint["Min"]
            constraint_max = constraint["Max"]
            constraints_violation[constraint_id] = np.zeros(gen_size)
            target = constraint_min if constraint_max == 0 else constraint_max

            if not constraint_type or constraint_type == '0':
                if constraint_max == 0:
                    constraint_type = "min"
                elif constraint_min == 0:
                    constraint_type = "max"
                else:
                    constraint_type = "minmax"
            if constraint_type == "min":
                target = constraint_min
                target_filter = constraint_value < target
                constraints_fitness[constraint_id] = constraint_value
                constraints_violation[constraint_id][target_filter] = constraint_value[target_filter] - target
            elif constraint_type == "max":
                target = constraint_max
                target_filter = constraint_value > target
                constraints_fitness[constraint_id] = constraint_value
                constraints_violation[constraint_id][target_filter] = target - \
                    constraint_value[target_filter]
            else:
                target_filter_min = (constraint_value < constraint_min)
                target_filter_max = (constraint_value > constraint_max)
                target_filter = target_filter_min | target_filter_max
                constraints_fitness[constraint_id] = constraint_value
                constraints_violation[constraint_id][target_filter_min] = constraint_value[target_filter_min] - constraint_min
                constraints_violation[constraint_id][target_filter_max] = constraint_max - \
                    constraint_value[target_filter_max]

            if target_filter.size:
                constraint_conditions[target_filter] = False
                # constraints_violation[constraint_id][target_filter] = constraints_fitness[constraint_id][target_filter]
        constraints = constraint_conditions.prod(axis=1)

        for i in range(gen_size):
            fitness[i]["is_infeasible"] = 1 - constraints[i]
            for o, objective_id in enumerate(self.objectives):
                fitness[i]["objectives"][objective_id] = objectives_fitness[objective_id][i]
            for c, constraint_id in enumerate(self.constraints):
                fitness[i]["constraints"][constraint_id] = constraints_fitness[constraint_id][i]
                fitness[i]["violations"][constraint_id] = constraints_violation[constraint_id][i]
        return fitness

    def get_fitness_python(self, generation, save_to_db=False):
        """get fitness for the population
        """
        gen_size = generation.shape[0]
        generation_fitness = self.evaluate_generation(generation)
        is_infeasible = np.array(
            [generation_fitness[i]["is_infeasible"] for i in range(gen_size)])
        objectives_array = np.array([[generation_fitness[i]["objectives"][self._objective_mapper[j]]
                                      for j in self._objective_mapper] for i in generation_fitness])
        constraints_array = np.array([[generation_fitness[i]["constraints"][self._constraint_mapper[j]]
                                       for j in self._constraint_mapper] for i in generation_fitness])
        violations_array = np.array([[generation_fitness[i]["violations"][self._constraint_mapper[j]]
                                      for j in self._constraint_mapper] for i in generation_fitness])
        return {"is_infeasible": is_infeasible,
                "objectives": objectives_array,
                "constraints": constraints_array,
                "violations": violations_array
                }

    def get_fitness(self, generation, save_to_db=False, calc_source="python"):
        """get generation fitness using the source.
        calc_source: "python", in-memory calculations are done in python locally
                     "crown", calculations are done in Crown database
        """
        # # self.logger.debug("getting generation fitness")
        if calc_source == "python":
            return self.get_fitness_python(generation, save_to_db)
        return self.get_fitness_crown_db(generation, save_to_db)

    def get_normalized_fitness(self, fitness):
        """Normalize fitness values
        Objectives are normalized using max, min objective values
        Constraints are normalized by subtracting modified_target and then dividing by modified_target
        """
        objectives_arr = fitness["objectives"]
        objectives_max = np.max(objectives_arr, axis=0)
        objectives_min = np.min(objectives_arr, axis=0)
        objectives_arr = (objectives_arr - objectives_min) / \
            (objectives_max - objectives_min + 1e-6)

        constraints_arr = fitness["constraints"]
        violations_arr = fitness["violations"]
        constraints_target = np.array(
            [self.constraints[self._constraint_mapper[i]]["target"] for i in range(self.n_constraints)])
        constraints_arr = (constraints_arr - constraints_target)
        return {"is_infeasible": fitness["is_infeasible"],
                "objectives": objectives_arr,
                "constraints": constraints_arr,
                "violations": violations_arr
                }

    def get_fitness_crown_db(self, generation, save_to_db=True):
        """returns fitness for every solution in the current population
        fitness = {solution_id: {"violated": 1, "objectives": {objective_id: value}, "constraints": {constraint_id: value}}}
        """
        raise NotImplementedError
        iteration = self._current_iteration
        self.save_generation_db(generation)
        if save_to_db:
            fitness_query = self.queries["fitness_query"].format(
                optimization_id=self.optimization_id, iteration=iteration, save_to_db=1)
        else:
            fitness_query = self.queries["fitness_query"].format(
                optimization_id=self.optimization_id, iteration=iteration, save_to_db=0)
        fitness_data, fitness_columns = get_data(fitness_query, self.connection_config, as_type="list", do_enumerate=False, columns=[
                                                 "ID", "OptimizationID", "ObjConsID", "SolutionID", "Criterion", "Functions", "Names", "CriterionValue"])
        fitnesses = fitness_array_to_dict(
            fitness_data, fitness_columns, self.constraints)
        return fitnesses

    def run_step_zero(self):
        """Run the first iteration in the optimization.
        This does not modify anything in the optimization object
        """
        P = self.initial_population.copy()
        current_generation = self.get_used_share(P)
        P_fitness = self.initial_fitness
        if not self.initial_fitness:
            P_fitness = self.get_fitness(current_generation, save_to_db=True)

        Q = self.crossover_and_mutate_population(self.initial_population)
        Q_fitness = self.get_fitness(self.get_used_share(Q), save_to_db=False)
        return P, P_fitness, Q, Q_fitness

    def run_one_iteration(self, P, P_fitness, Q, Q_fitness):
        """Run a single iteration in the optimization.
        This does not modify anything in the optimization object
        """
        R, R_fitness = self.combine_populations(P, Q, P_fitness, Q_fitness)
        P, P_fitness = self.selection(
            R, R_fitness, self.objectives, self.constraints)
        Q = self.crossover_and_mutate_population(P)
        Q_fitness = self.get_fitness(self.get_used_share(Q), save_to_db=False)
        return P, P_fitness, Q, Q_fitness

    def run(self):
        self.logger.info("Running optimization")
        self.generations = []
        self.populations = []
        self.generation_fitnesses = []
        P = self.initial_population.copy()
        current_generation = self.get_used_share(P)
        P_fitness = self.initial_fitness
        # self.output_progress_file = os.path.join(output_dir, "progress.csv")
        if not P_fitness:
            P_fitness = self.get_fitness(current_generation, save_to_db=True)
        Q = self.crossover_and_mutate_population(self.initial_population)
        Q_fitness = self.get_fitness(self.get_used_share(Q), save_to_db=False)
        self._current_iteration = 0
        self.populations.append(P)
        self.generations.append(current_generation)
        self.generation_fitnesses.append(P_fitness)
        self.save_outputs(self._current_iteration, current_generation, P_fitness,new=True)
        for i in range(self.iterations):
            self.logger.info(f"running iteration ({i + 1}/{self.iterations})")
            self._current_iteration += 1
            R, R_fitness = self.combine_populations(P, Q, P_fitness, Q_fitness)
            P, P_fitness = self.selection(
                R, R_fitness, self.objectives, self.constraints)
            Q = self.crossover_and_mutate_population(P)
            Q_fitness = self.get_fitness(
                self.get_used_share(Q), save_to_db=False)
            current_generation = self.get_used_share(P)
            self.populations.append(P)
            self.generations.append(current_generation)
            self.generation_fitnesses.append(P_fitness)
            self.save_outputs(self._current_iteration, current_generation, P_fitness)
        # Run it one last time
        self._current_iteration += 1
        R, R_fitness = self.combine_populations(P, Q, P_fitness, Q_fitness)
        P, P_fitness = self.selection(
            R, R_fitness, self.objectives, self.constraints)
        current_generation = self.get_used_share(P)
        self.generations.append(current_generation)
        self.generation_fitnesses.append(P_fitness)
        self.save_outputs(self._current_iteration, current_generation, P_fitness)
        self.logger.info("optimization finished")
        final_output_fpath = os.path.join(self.output_dir, "Optimization_output.xlsx")
        save_optimization_output(final_output_fpath, self)
        self.logger.info(f"Output saved at {final_output_fpath}")

    def save_outputs(self, iteration, P, P_fitness, new=False):
        self.save_progress(iteration, P_fitness, new=new)
        # self.save_shares(iteration, P)
        # self.save_objectives(iteration, P_fitness)
        # self.save_constraints(iteration, P_fitness)

    def save_progress(self, iteration, generation_fitness, new=False):
            if new:
                f =  open(self.output_progress_file, "w")
            else:
                f =  open(self.output_progress_file, "a")
            feasible_count = self.population_size - np.sum(generation_fitness["is_infeasible"])
            header_string = "Iteration,FeasibleCount"
            data = f"{iteration},{feasible_count}"
            for o, id_ in enumerate(self.objectives):
                header_string += ',"' + self.objectives[id_]["Names"] + '"'
                data += f",{np.mean(generation_fitness['objectives'][:, o])}"
            for o, id_ in enumerate(self.constraints):
                header_string += ',"' + self.constraints[id_]["Names"] + '"'
                data += f",{np.mean(generation_fitness['constraints'][:, o])}"
            if new:
                f.write(header_string + "\n")
            f.write(data + "\n")
            f.close()
        
    def save_shares(self, iteration, generation):
        shares_file = self.output_generations_file.format(iteration=iteration)
        with open(shares_file, "w") as f:
            for j in range(self.population_size):                
                line = ','.join([str(generation[j, i]) for i in range(self.n_layers)])
                f.write(line + "\n")
                # f.flush()

    def save_objectives(self, iteration, generation_fitness):
        objectives_file = self.output_objectives_file.format(iteration=iteration)
        obj = {}
        with open(objectives_file, "w") as f:
            for j in self.objectives:
                id_ = self.objectives[j]["id"]
                obj[j] = generation_fitness["objectives"][:, id_].flatten().tolist()
            json.dump(obj, f)

    def save_constraints(self, iteration, generation_fitness):
        constraints_file = self.output_constraints_file.format(iteration=iteration)
        obj = {}
        with open(constraints_file, "w") as f:
            for j in self.constraints:
                id_ = self.constraints[j]["id"]
                obj[j] = generation_fitness["constraints"][:, id_].flatten().tolist()
                obj[-1] = generation_fitness["is_infeasible"].flatten().tolist()
            json.dump(obj, f)



    def save_generation_db(self, generation, fitness):
        raise NotImplementedError
        iteration = self._current_iteration
        table_name = self.base_table_name + f'_iter_{iteration}'
        shares_table = table_name + "_shares"
        # insert shares data
        port_ids_shares = [i for j in range(
            self.n_layers) for i in range(self.population_size)]
        # ordered by portfolio id then layer
        flattened_generation = generation.flatten()
        flattened_activities = [self.input_data[i]["CrownActivityID"]
                                for i in self.input_data] * self.population_size
        data = []
        for i in range(self.population_size * self.n_layers):
            data.append(
                (port_ids_shares[i], flattened_activities[i], flattened_generation[i]))

        columns = ["solution_id", "ActivityID", "share"]
        dtypes = ["int", "int", "float"]
        create_db_table(shares_table, columns, dtypes, self.connection_config)
        push_to_db(data, columns, self.connection_config,
                   shares_table, append=False, chunk_size=1000)

        # insert objectives and constraints data

    def save_generation_local(self, iteration, generation, fitness, output_dir=None):
        """save data locally
        """
        raise NotImplementedError
        if not output_dir:
            output_dir = self.output_dir
        else:
            output_dir = os.path.join(
                output_dir, f"Optimization_{self.optimization_id}")
            os.makedirs(output_dir, exist_ok=True)
        # iteration = self._current_iteration
        gen_filepath = os.path.join(output_dir, f"opt_{iteration}_shares.csv")
        objective_filepath = os.path.join(
            output_dir, f"opt_{iteration}_objectives.csv")
        constraint_filepath = os.path.join(
            output_dir, f"opt_{iteration}_constraints.csv")
        infeasible_filepath = os.path.join(
            output_dir, f"opt_{iteration}_infeasible.csv")
        generation = generation.T
        headers = "InstrumentID," + \
            ",".join([str(i) for i in range(self.population_size)]) + "\n"
        gen_file = open(gen_filepath, "w")
        gen_file.write(headers)
        gen_file.flush()
        for i in range(self.n_layers):
            instrument_id = self.input_data[i]["ILSInstrumentID"]
            data = f"{instrument_id}," + \
                ",".join([generation[i, j]
                          for j in range(self.population_size)]) + "\n"
            gen_file.write(data)
        gen_file.close()
