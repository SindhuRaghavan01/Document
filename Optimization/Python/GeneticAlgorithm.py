import pandas
from numpy.random import randint, rand, choice
import numpy as np


class GA():
    def __init__(self, input_data, objectives, constraints, parameters, seeded=False):
        """Genetic Algorithm Class
        """
        self.parameters = parameters
        self.seeded = seeded
        self.objectives = objectives
        self._objective_ids = list(objectives.keys())
        self.constraints = constraints
        self._constraint_ids = list(constraints.keys())
        self._objective_mapper = {i: k for i, k in enumerate(self.objectives)}
        self._constraint_mapper = {i: k for i,
                                   k in enumerate(self.constraints)}
        self.n_objectives = len(self._objective_mapper)
        self.n_constraints = len(self._constraint_mapper)
        self._current_iteration = 0
        self._set_params(parameters)
        for id_ in self.objectives:
            for j in self._objective_mapper:
                if id_ == self._objective_mapper[j]:
                    self.objectives[id_]["id"] = j
        for id_ in self.constraints:
            for j in self._constraint_mapper:
                if id_ == self._constraint_mapper[j]:
                    self.constraints[id_]["id"] = j
        # self._initialize_data(input_data)

    def _set_params(self, parameters):
        """Set up optimization parameters
        """
        self.solver = parameters.get("Solver", "NSGA2")
        self.n_bits = parameters.get("n_bits", 32)
        self.mutation_probability = parameters.get(
            "mutation_probability", 0.01)
        self.population_size = parameters.get("population_size", 500)
        self.crossover_method = parameters.get("crossover_method", "SBX")
        self.crossover_eta_c = parameters.get("crossover_eta_c", 2)
        self.crossover_probability = parameters.get(
            "crossover_probability", 0.1)
        self.mutation_method = parameters.get("mutation_method", "polynomial")
        self.mutation_eta_m = parameters.get("mutation_eta_m", 3)
        self.iterations = parameters.get("Iterations", 2000)
        # self.seeding_override = parameters.get("SeedingOverride", self.population_size // 2)

    def _initialize_data(self, input_data):
        pass

    def mutate(self, population, method="polynomial", xmin=0, xmax=1, mutation_probability=None):
        """Apply mutation to the entire population in one go with mutation probability.
        parameters:
            method:
                binary: mutate a single bit. Mutation turns ON/OFF a single bit in the input value
                polynomial: change the value based on distribution
                                delta = (2 * r)^(1/(eta_m + 1)) if r <= 0.5,
                                        1 - (2 * (1 - r))^(1 / (eta_m + 1)) otherwise
            xmin, xmax: used only if `polynomial` method is applied
        """
        if not mutation_probability:
            mutation_probability = self.mutation_probability
        # mutation_probability = mutation_probability / 100
        y = population.copy()
        if method == "binary":
            is_mutate = choice([0, 1], size=population.shape, replace=True, p=[
                               1 - mutation_probability, mutation_probability]).astype(int)
            mutation_sites = choice(np.linspace(
                0, self.n_bits, self.n_bits + 1, dtype=int), size=population.shape)
            mutation_values = (2 ** mutation_sites) * is_mutate
            is_subtract = population.astype(int) & mutation_values.astype(int)
            y[is_subtract > 0] = y[is_subtract > 0] - \
                mutation_values[is_subtract > 0]
            y[is_subtract < 0] = y[is_subtract < 0] + \
                mutation_values[is_subtract < 0]
        if method == "simple":
            is_mutate = choice([0, 1], size=population.shape, replace=True, p=[
                            1 - mutation_probability, mutation_probability]).astype(int)
            # print(np.sum(is_mutate), is_mutate.size)
            u = np.random.rand(*population.shape)
            y[is_mutate == 1] = u[is_mutate == 1]

        elif method == "polynomial":
            # u = choice([0, 1], p=[1 - self.mutation_probability, self.mutation_probability])
            y = population
            # if u == 1:
            r = rand(*population.shape)
            # c is used to choose which genes get mutated in a solution
            c = choice([0, 1], size=population.shape, replace=True, p=[
                       1 - self.mutation_probability, self.mutation_probability])
            delta = np.power(2 * r, (1 / (self.mutation_eta_m + 1))) - 1
            delta[r > 0.5] = 1 - \
                np.power(2 * (1 - r[r > 0.5]), (1 / (self.mutation_eta_m + 1)))
            # y = population + delta * (population - xmin)
            # y[r > 0.5] = population[r > 0.5] + \
            #     delta[r > 0.5] * (xmax - population[r > 0.5])
            y = population + delta * (xmax - xmin) * c
            y[c == 0] = population[c == 0]
            y[y < 0] = 0
            y[y > 1] = 1
        return y

    def crossover_solutions(self, solution_1, solution_2, method="SBX"):
        """Apply crossover for two solutions
        calls self.crossover_chromosomes for each component in the solution

        Select a crossover site, and apply binary splicing to generate two output values
        x1, x2 must be represented in (n_bits)-bit format
        method: single_binary -> simple single point crossover with uniform probability
                SBX -> Simulated Binary Crossover. Uses default distribution index=5
        """
        if method == "single_binary":
            positions = choice(list(range(0, self.n_bits - 1)),
                               size=self.n_layers, replace=True).astype(int)
            p = (solution_2 - solution_1).astype(int) & (2 ** positions - 1)
            output_1 = solution_1 + p
            output_2 = solution_2 - p
        elif method == "SBX":
            u = rand(self.n_layers)
            beta = np.power(2 * u, 1 / (self.crossover_eta_c + 1))
            beta[u >= 0.5] = np.power(
                1 / (2 * (1 - u[u >= 0.5])), 1 / (self.crossover_eta_c + 1))
            output_1 = 0.5 * ((solution_1 + solution_2) - beta * (solution_2 - solution_1))
            output_2 = 0.5 * ((solution_1 + solution_2) + beta * (solution_2 - solution_1))
            # u = choice([0, 1], size=self.n_layers, replace=True, p=[
            #            1 - self.crossover_probability, self.crossover_probability])
            # output_1[u == 0] = solution_1[u == 0]
            # output_2[u == 0] = solution_2[u == 0]

        elif method == "weighted":
            u = rand(self.n_layers) # * choice([0, 1], size=self.n_layers, replace=True, p=[
                # 1 - self.crossover_probability, self.crossover_probability])
            output_1 = u * solution_1 + (1 - u) * solution_2
            output_2 = u * solution_2 + (1 - u) * solution_1
        output_1 = np.maximum(np.minimum(output_1, np.ones(
            self.n_layers)), np.zeros(self.n_layers))
        output_2 = np.maximum(np.minimum(output_2, np.ones(
            self.n_layers)), np.zeros(self.n_layers))
        return output_1, output_2

    def crossover_population(self, population, method="SBX"):
        """Apply crossover step for the current population
        """
        n_rounds = self.population_size // 2
        new_ids = choice(list(range(self.population_size)),
                         self.population_size, replace=False)
        new_population = {key: np.empty_like(
            population[key]) for key in population}
        for k in range(n_rounds):
            combined_1 = {}
            combined_2 = {}
            for key in population:
                combined_1[key], combined_2[key] = self.crossover_solutions(
                    population[key][new_ids[2 * k], :], population[key][new_ids[2 * k + 1], :], method=method)
                new_population[key][2 * k, :] = combined_1[key][:]
                new_population[key][2 * k + 1, :] = combined_2[key][:]
        return new_population

    def crossover_and_mutate_population(self, population, crossover_method="SBX", mutation_method="polynomial"):
        crossed_population = self.crossover_population(
            population, method=crossover_method)
        mutated_population = {}
        for key in crossed_population:
            mutated_population[key] = self.mutate(
                crossed_population[key], method=mutation_method)
        return mutated_population

    def get_fitness(self, population):
        """Not Implemented. Overwrite with custom function
        get fitness values for the input population

        """
        raise NotImplementedError

    def selection(self, population, fitness_values, objectives, constraints, method="pareto"):
        """population is (2 * pop_size, n_layers) array
        fitness_values = {"is_infeasible": (2 * pop_size, ),
                          "objectives": (2 * pop_size, n_objectives),
                          "constraints": (2 * pop_size, n_constraints)
                          }
        method: pareto or scalar.
                pareto : find pareto fronts and use distance sorting
                scalar : use distance from circle centered at max values across each dimension
        """
        is_infeasible = fitness_values["is_infeasible"]
        is_feasible_filter = (is_infeasible == 0)
        feasible_count = np.int(
            2 * self.population_size - np.sum(is_infeasible))

        feasible_infeasibility = is_infeasible[is_feasible_filter]
        feasible_objectives = fitness_values["objectives"][is_feasible_filter, :]
        feasible_constraints = fitness_values["constraints"][is_feasible_filter, :]
        feasible_violations = fitness_values["violations"][is_feasible_filter, :]
        feasible_population = {
            key: population[key][is_feasible_filter, :] for key in population}

        infeasible_infeasibility = is_infeasible[~is_feasible_filter]
        infeasible_objectives = fitness_values["objectives"][~is_feasible_filter, :]
        infeasible_constraints = fitness_values["constraints"][~is_feasible_filter, :]
        infeasible_violations = fitness_values["violations"][~is_feasible_filter, :]
        # infeasible_population = population[~is_feasible_filter, :]
        infeasible_population = {
            key: population[key][~is_feasible_filter, :] for key in population}

        if feasible_count == self.population_size:
            new_population = feasible_population.copy()
            new_fitness_values = {"is_infeasible": feasible_infeasibility,
                                  "objectives": feasible_objectives,
                                  "constraints": feasible_constraints,
                                  "violations": feasible_violations
                                  }
        elif feasible_count < self.population_size:
            # select all feasible solutions, and add least constraint violating solutions
            remaining_count = np.int(self.population_size - feasible_count)
            new_population = feasible_population.copy()
            normalized_violations = (infeasible_violations - np.min(infeasible_violations, axis=0)) / (
                np.max(infeasible_violations, axis=0) - np.min(infeasible_violations, axis=0) + 1e-6)
            if method == "scalar":
                fronts = self.get_pareto_fronts(
                    infeasible_violations, target="min")
                counts = 0
                i = 0
                top_n = []
                while remaining_count > 0:
                    if len(fronts[i]) > remaining_count:
                        # select remaining solutions
                        args = self.get_distances(
                            normalized_violations, fronts[i], remaining_count)
                        top_n.extend(args)
                        remaining_count = 0
                    else:
                        # select all solutions in the current front
                        top_n.extend(fronts[i].tolist())
                        counts += len(fronts[i])
                        remaining_count -= len(fronts[i])
                    i += 1
            if method == "pareto":
                distances = np.ones(self.n_constraints) - normalized_violations
                distances = np.sum(distances ** 2, axis=1)
                argdist = np.argsort(distances, axis=0)
                top_n = argdist[: remaining_count]
            new_population = {key: np.concatenate(
                [new_population[key], infeasible_population[key][top_n, :]], axis=0) for key in population}
            new_fitness_values = {"is_infeasible": np.concatenate([feasible_infeasibility, infeasible_infeasibility[top_n]], axis=0),
                                  "objectives": np.concatenate([feasible_objectives, infeasible_objectives[top_n, :]], axis=0),
                                  "constraints": np.concatenate([feasible_constraints, infeasible_constraints[top_n, :]], axis=0),
                                  "violations": np.concatenate([feasible_violations, infeasible_violations[top_n, :]], axis=0),
                                  }
        else:
            normalized_objectives = (feasible_objectives - np.min(feasible_objectives)) / (
                np.max(feasible_objectives) - np.min(feasible_objectives) + 1e-6)
            # select top (population_size) solutions from the feasible solutions
            remaining_count = self.population_size
            if method == "pareto":
                fronts = self.get_pareto_fronts(
                    normalized_objectives, target="max")
                counts = 0
                i = 0
                top_n = []
                while len(top_n) < self.population_size:
                    if len(fronts[i]) > remaining_count:
                        # select remaining solutions
                        # args = self.get_distances(
                            # normalized_objectives, fronts[i], remaining_count)
                        args = choice(fronts[i], size=remaining_count, replace=False)
                        top_n.extend(args)
                        remaining_count = 0
                    else:
                        # select all solutions in the current front
                        top_n.extend(fronts[i].tolist())
                        counts += len(fronts[i])
                        remaining_count -= len(fronts[i])
                    i += 1
                if len(top_n) > self.population_size:
                    top_n = choice(
                        top_n, size=self.population_size, replace=True)
            if method == "scalar":
                distances = np.ones(self.n_objectives) - normalized_objectives
                distances = np.sum(distances ** 2, axis=1)
                argdist = np.argsort(-distances, axis=0)
                top_n = argdist[: self.population_size]
            new_population = {
                key: feasible_population[key][top_n, :] for key in population}
            new_fitness_values = {"is_infeasible": np.zeros(self.population_size),
                                  "objectives": feasible_objectives[top_n, :],
                                  "constraints": feasible_constraints[top_n, :],
                                  "violations": feasible_violations[top_n, :]
                                  }
        return new_population, new_fitness_values

    def get_distances(self, criteria_values, front, remaining_count):
        v = criteria_values[front, :]
        s = np.lexsort(v.T)  # sort solutions
        d = v[s, :]
        distances = np.zeros(len(front))
        distances[0] = 10
        distances[-1] = 10
        # total manhatten distance from nearest neighbors
        distances[1:-1] = np.sum(np.abs(d[2:, :] - d[1:-1, :]) +
                                 np.abs(d[:-2, :] - d[1:-1, :]), axis=1)
        t = np.argsort(distances)
        args = t[-remaining_count:]  # select solutions farther from each other
        args = front[s[args]].tolist()
        return args

    def combine_populations(self, population_1, population_2, fitness_values_1, fitness_values_2):
        """Combine populations from two steps before moving on to running the iteration steps
        """
        # create new ids
        # assign ids from population 1
        # assign ids from population 2
        combined_population_len = 2 * self.population_size
        # combined_population = np.concatenate((population_1, population_2), axis=0)
        combined_fitness = {}
        combined_population = {}
        for key in population_1.keys():
            combined_population[key] = np.concatenate(
                (population_1[key], population_2[key]), axis=0)
        for key in fitness_values_1.keys():
            combined_fitness[key] = np.concatenate(
                (fitness_values_1[key], fitness_values_2[key]))

        return combined_population, combined_fitness

    def run(self):
        self.populations = []
        self.population_fitnesses = []
        P = self.initial_population
        P_fitness = self.initial_fitness
        Q = self.crossover_and_mutate_population(self.initial_population)
        Q_fitness = self.get_fitness(Q)

        self.populations.append(P)
        self.population_fitnesses.append(P_fitness)
        for i in range(self.iterations):
            R, R_fitness = self.combine_populations(P, Q, P_fitness, Q_fitness)
            P, P_fitness = self.selection(
                R_fitness, self.objectives, self.constraints)
            Q = self.crossover_and_mutate_population(P)
            Q_fitness = self.get_fitness(Q)
            self.populations.append(P)
            self.population_fitnesses.append(P_fitness)
        # Run it one last time
        R, R_fitness = self.combine_populations(P, P_fitness, Q, Q_fitness)
        P, P_fitness = self.selection(
            R_fitness, self.objectives, self.constraints)
        self.populations.append(P)
        self.population_fitnesses.append(P_fitness)

    def get_pareto_front_zero(self, criteria_values, target="min"):
        """https://stackoverflow.com/questions/32791911/fast-calculation-of-pareto-front-in-python

        """
        i = 0
        n_pop = criteria_values.shape[0]
        is_eff = np.arange(n_pop)
        v = criteria_values.copy()
        while i < v.shape[0]:
            if target == "max":
                nd_mask = np.any(v >= v[i, :], axis=1)
            else:
                nd_mask = np.any(v <= v[i, :], axis=1)
            nd_mask[i] = True
            is_eff = is_eff[nd_mask]
            v = v[nd_mask]
            i = np.sum(nd_mask[:i]) + 1
        return is_eff

    def get_pareto_fronts(self, criteria_values, target="min"):
        n_pop = criteria_values.shape[0]
        _ids = np.arange(n_pop)
        counts = 0
        fronts = []
        v = criteria_values.copy()
        while counts < n_pop:
            is_eff = self.get_pareto_front_zero(v)
            fronts.append(_ids[is_eff])
            _ids = np.delete(_ids, is_eff)
            v = np.delete(v, is_eff, axis=0)
            counts += len(is_eff)
        return fronts
