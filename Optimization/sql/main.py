import json
import os
import pickle
import sys
from datetime import datetime
from time import process_time

import pandas
from tqdm import tqdm

from Crown_GA import Crown_GA
from sql.data_queries import optimization_queries
from utils.db_utils import get_connection, get_data, push_to_db
from utils.helpers import save_fitness_input_to_db, get_fitness_logger_string, save_optimization_output
from utils.logger import Logger

LOG_LEVEL = Logger.DEBUG

with open("config/sql_connection.json", "r") as f:
    j = json.load(f)
    server_config = {i: j["ServerConfig"][i] for i in j["ServerConfig"] if i in [
        "DRIVER_NAME", "SERVER_NAME", "DB_NAME", "USERNAME", "PASSWORD"]}

with open("config/defaults.json", "r") as f:
    defaults = json.load(f)

    default_parameters = defaults["parameters"]


# connection = get_connection(**server_config)
if __name__ == "__main__":
    OptimizationID = 2
    now = datetime.strftime(datetime.now(), "%Y%m%d-%H%M%S")
    logfilepath = f"logs/optimization_{OptimizationID}_{now}.log"
    if os.path.exists(logfilepath):
        os.remove(logfilepath)

    logger = Logger(logfilepath, log_level=LOG_LEVEL)
    output_dir = f"output/{OptimizationID}/"
    os.makedirs(output_dir, exist_ok=True)
    progress_file_path = os.path.join(output_dir, "progress.csv")
    optimization = Crown_GA(OptimizationID, connection_config=server_config,
                            queries=optimization_queries, defaults=defaults, logger=logger, output_dir=output_dir)
    optimization.iterations = 10
    optimization.run()
    
    # optim_dump = os.path.join(output_dir, "optimization_dump.pkl")
    # with open(optim_dump, "wb") as f:
    #     pickle.dump(optimization, f)
    