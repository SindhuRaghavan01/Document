from .db_utils import push_to_db, get_connection, get_data
from datetime import datetime
import pandas
from sqlalchemy.sql import text
from multiprocessing import Pool
import numpy as np

def population_dict_to_sql_array(input_data, population):
    """convert the population data into list of lists to be used to save in the the database
    """
    data = []
    for solution_number in population:
        for layer_number in population[solution_number]:
            data.append([solution_number, layer_number, input_data[layer_number]["ILSInstrumentID"], population[solution_number][layer_number]])
    columns = ["solution_number", "layer_number", "ILSInstrumentID", "share"]
    return columns, data

def save_fitness_input_to_db(input_data, columns, connection_config, table_name, append=False):
    """send the data required for fitness calculations to the database
    """
    if not append:
        # drop existing table if found
        # create new table
        pass
    status = push_to_db(input_data, columns, connection_config, table_name, append=append)
    return status

def fitness_array_to_dict(array, columns, constraints):
    """array is N x 5 list of lists
    The columns are [SolutionID, Criterion, ObjConsID, CriterionValue]

    any additional columns must be after these 5 columns. They will NOT be used
    """
    sol_id_col = columns.index("SolutionID")
    # violated_col = columns.index("violated")
    criter_col = columns.index("Criterion")
    criter_id_col = columns.index("ObjConsID")
    criter_val_col = columns.index("CriterionValue")
    output = {}
    for row in array:
        for column in array:
            solution_id = column[sol_id_col]
            # violated = column[violated_col]
            criterion_id = column[criter_id_col]
            criterion_value = column[criter_val_col]
            criterion = column[criter_col]
            if not output:
                output = {solution_id: {criterion: {criterion_id: criterion_value}}}
            elif not output[solution_id]:
                output[solution_id] = {criterion: {criterion_id: criterion_value}}
            else:
                output[solution_id][criterion][criterion_id] = criterion_value
    # assign is_infeasible": is_infeasible, to every output[solution_id]
    for solution_id in output:
        is_infeasible = 0
        for constraint_id in constraints:
            constraint_type = constraints[constraint_id]["Types"]
            constraint_value = constraints[constraint_id]["value"]
            if constraint_type == "max":
                if output[solution_id][constraint_id] > constraint_value:
                    is_infeasible = 1
                    break
            if constraint_type == "min":
                if output[solution_id][constraint_id] < constraint_value:
                    is_infeasible = 1
                    break
        output[solution_id]["is_infeasible"] = is_infeasible
    return output

def create_fitness_table(connection_config, fitness_query, base_table_name, iteration):
    """create fitness table using the query defined in data_queries.py
    """
    fitness_table_name = f"{base_table_name}_{iteration}"
    query = fitness_query.format(fitness_table_name=fitness_table_name)
    # logging.info(query)
    connection = get_connection(**connection_config)
    with connection.cursor() as cursor:
        for q in query.split(";"):
            result = cursor.execute(q)
    connection.commit()
    return result

# @staticmethod
def get_layer_ylts(ylt_queries, connection_config):
    """get layer ylts as numpy array
    """
    with Pool(16) as p:
        params = [(query, connection_config, 'list') for query in ylt_queries]
        r = p.starmap(get_data, params)
        cat_ylt = np.array([i[0] for i in r])
        columns = r[0][1]
        loss_column_index = columns.index("Loss")
        year_column_index = columns.index("Year")
        layer_ylts = cat_ylt[:, :, loss_column_index]
        # layer_years = cat_ylt[:, :, year_column_index]
    return layer_ylts


def get_objective_indices_by_names(optimization, objective_names):
    index_list = {}
    for objective_name in objective_names:
        for i in optimization.objectives:
            if optimization.objectives[i]["Names"] == objective_name:
                for j in optimization._objective_mapper:
                    if optimization._objective_mapper[j] == i:
                        index_list[objective_name] = j
    return index_list

def get_constraint_indices_by_names(optimization, constraint_names):
    index_list = {}
    for constraints_name in constraint_names:
        for i in optimization.constraints:
            if optimization.constraints[i]["Names"] == constraints_name:
                for j in optimization._constraint_mapper:
                    if optimization._constraint_mapper[j] == i:
                        index_list[constraints_name] = j
    return index_list

def get_fitness_logger_string(iteration, P_fitness, optimization):
    # find indices for target variables
    string = f"{iteration}"
    objective_names = ["Expected Return", "100yr AEP TVaR"]
    constraint_names = ["Catbonds Constraints", "Non--Peak Cat B", "Portfolio Level Min"]
    objective_ids = get_objective_indices_by_names(optimization, objective_names)
    constraint_ids = get_constraint_indices_by_names(optimization, constraint_names)
    for objective_name in objective_names:
        string += f",{np.mean(P_fitness['objectives'][:, objective_ids[objective_name]]):.0f}"
    for i in constraint_names:
        string += f",{np.mean(P_fitness['constraints'][:, constraint_ids[i]]):.0f}"
    string += f",{500 - np.sum(P_fitness['is_infeasible']):.0f}"
    # "Iteration,ExpReturn,TVAR100,PeakCATBondAUM,NonPeakCATBondAUM,PortAUM,FeasibleCount"
    return string


def save_optimization_output(output_fpath, optimization, generation=None, objective_values=None, constraint_values=None, feasibility_values=None):
    generation = optimization.generations[-1] if generation is None else generation
    objective_values = optimization.generation_fitnesses[-1]["objectives"] if objective_values is None else objective_values
    constraint_values = optimization.generation_fitnesses[-1]["constraints"] if constraint_values is None else constraint_values
    feasibility_values = (1 - optimization.generation_fitnesses[-1]["is_infeasible"]) if feasibility_values is None else feasibility_values

    writer = pandas.ExcelWriter(output_fpath, engine='xlsxwriter')

    # shares data
    df_data = []
    headers = ["id", "ILSInstrumentID", "Updated_LayerID", "description", "CipherID", "OptimizationLayerID", "AUM", "Premium", "Min_Share", "Max_Share", "layer_group", "Participation", "CrownActivityID", "AnalyzeReActivityID", "Initial_Fees"]
    for i in optimization.input_data:
        data = optimization.input_data[i]
    #     print(data)
    #     print(",".join([str(i)] + [str(data[j]) for j in ["ILSInstrumentID", "Updated_LayerID", "description", "CipherID", "OptimizationLayerID", "AUM", "Premium", "Min_Share", "Max_Share", "layer_group", "Participation", "CrownActivityID", "AnalyzeReActivityID"]]))
        df_data.append([i] + [data[j] for j in ["ILSInstrumentID", "Updated_LayerID", "description", "CipherID", "OptimizationLayerID", "AUM", "Premium", "Min_Share", "Max_Share", "layer_group", "Participation", "CrownActivityID", "AnalyzeReActivityID", "Initial_Fees"]])
    #     break
    df = pandas.DataFrame(df_data, columns=headers)
    df_shares = pandas.DataFrame(generation.T, columns=[str(i) for i in range(optimization.population_size)])
    data_df = pandas.concat([df, df_shares], axis=1)
    
    # objective values
    data = []
    headers = ["Names", "Functions", "Types", "Return_Period", "Loss_Perspective", "Fields", "ProportionateTo", "Layer_Groups", "LossFilter"]
    for i, id_ in enumerate(optimization.objectives):
        d = optimization.objectives[id_]
        data.append([i, id_] + [d[j] for j in headers])
    headers = ["id", "ObjConsID"] + headers
    df_obj = pandas.DataFrame(data, columns=headers).T
    df_obj_values = pandas.DataFrame(objective_values, columns=[(i) for i in range(optimization.n_objectives)])
    df_obj = pandas.concat([df_obj, df_obj_values], axis=0)

    df_obj.columns = df_obj.loc["Names"]
    df_obj = df_obj.drop(["id", "Names"]).reset_index()

    #constraint values
    data = []
    headers = ["Names", "Functions", "Types", "Min", "Max","Return_Period", "Loss_Perspective", "Fields", "ProportionateTo", "Layer_Groups", "LossFilter"]
    for i, id_ in enumerate(optimization.constraints):
        d = optimization.constraints[id_]
        data.append([i, id_] + [d[j] for j in headers])
    headers = ["id", "ObjConsID"] + headers
    df_cons = pandas.DataFrame(data, columns=headers).T
    df_cons_values = pandas.DataFrame(constraint_values, columns=[(i) for i in range(optimization.n_constraints)])
    df_cons = pandas.concat([df_cons, df_cons_values], axis=0)
    df_cons.columns = df_cons.loc["Names"]
    df_cons["IsFeasible"] = ""
    for i in range(optimization.population_size):
        df_cons.loc[i, "IsFeasible"] = feasibility_values[i]
    df_cons = df_cons.drop(["id", "Names"]).reset_index()

    # save file
    data_df.to_excel(writer, "Optimized Shares", index=False)
    df_obj.to_excel(writer, "Objective Values", index=False)
    df_cons.to_excel(writer, "Constraint Values", index=False)
    writer.close()

def excelInt_to_dateStr(excelInt):
    if not excelInt:
        return None
    if isinstance(excelInt, str):
        return ""
    elif isinstance(excelInt, float):
        return excelInt
    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excelInt - 2)
    return datetime.strftime(dt, "%Y-%m-%d")
