import urllib
import pandas
import logging
from sqlalchemy import create_engine
import pyodbc

def get_connection(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD):
    """create sql connection using pyodbc
    """
    param_string = f"DRIVER={DRIVER_NAME};SERVER={SERVER_NAME};DATABASE={DB_NAME};UID={USERNAME};PWD={PASSWORD}",
    connection = pyodbc.connect(param_string)
    return connection

def get_connection_sqlalchemy(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD, echo=False):
    """
    create sql connection
    """
    param_string = "DRIVER={};SERVER={};DATABASE={};UID={};PWD={}"
    params = urllib.parse.quote_plus(param_string.format(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD))
    conn_string = "mssql+pyodbc:///?odbc_connect={}".format(params)
    connection = create_engine(conn_string, echo=echo)
    return connection

def sql_stringify(x):
    """add quotes to the input if it's a string, otherwise pass it as-is
    TODO: add date formatting
    """
    if isinstance(x, str):
        return f"'{x}'"
    return f"{x}"

def push_to_db(data, columns, connection, table_name, append=True, chunk_size=1000):
    """push data to database
    paramaters
        data: DataFrame or list.
              if list, then columns must be given
    """
    try:
        if isinstance(data, pandas.DataFrame):
            if append:
                data.to_sql(table_name, con=connection, if_exists="append", index=False)
            else:
                data.to_sql(table_name, con=connection, if_exists="replace", index=False)
        else:
            if not columns:
                return False
            n_rows = len(data)
            n_inserts = n_rows // chunk_size + 1
            logging.debug(f"Total rows: {n_rows}")
            logging.debug(f"Total inserts: {n_inserts}")
            print(f"Total rows: {n_rows}")
            print(f"Total inserts: {n_inserts}")
            for i in range(n_inserts):
                query = f"insert into {table_name}({','.join(columns)}) values"
                if i * chunk_size == n_rows:
                    break
                for j in range(i * chunk_size, (i + 1) * chunk_size):
                    if j < n_rows:
                        query += "(" + ",".join([sql_stringify(k) for k in data[j]]) + "),"
                r = connection.execute(query[:-1])

    except Exception as e:
        print(e)
        return False
    return True

def get_data(query, connection, as_type="dict", key_column="", do_enumerate=True):
    """execute the query and get the data
    parameters:
        as_type: dict => enumerates the results. the result is a dict of dicts
                 dataframe => returns pandas dataframe
                 list => returns list of lists
        key_column: Used only if do_enumerate=False and as_type="dict". Takes precedence over do_enumerate
        do_enumerate: bool. Used only if as_type="dict"
    """
    if as_type == "dict" and key_column.strip() != "":
        do_enumerate = False
    if as_type == "dict":
        r = connection.execute(query)
        keys = r.keys()
        data = {} if do_enumerate or key_column else []
        for i, result in enumerate(r):
            if key_column:
                col_index = keys.index(key_column)
                obj = {keys[j]: result[j] for j in range(len(keys)) if keys[j] != key_column}
                data[result[col_index]] = obj
            elif do_enumerate:
                obj = {keys[j]: result[j] for j in range(len(keys))}
                data[i] = obj
            else:
                obj = {keys[j]: result[j] for j in range(len(keys))}
                data.append(obj)
        columns = list(keys)
    elif as_type == "list":
        r = connection.execute(query)
        keys = r.keys()
        data = []
        for row in r:
            data.append([i for i in r])
        columns = list(keys)
    else:
        data = pandas.query_as_df(query, con=connection)
        columns = data.columns.tolist()
    return data, columns
