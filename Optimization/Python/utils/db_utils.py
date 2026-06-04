import urllib
import pandas
import logging
# from sqlalchemy import create_engine
import pyodbc
import numpy as np

def get_connection(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD):
    """create sql connection using pyodbc
    """
    param_string = f"DRIVER={DRIVER_NAME};SERVER={SERVER_NAME};DATABASE={DB_NAME};UID={USERNAME};PWD={PASSWORD}"
    connection = pyodbc.connect(param_string, autocommit=True)
    return connection

def sql_stringify(x):
    """add quotes to the input if it's a string, otherwise pass it as-is
    TODO: add date formatting
    """
    if isinstance(x, str):
        return f"'{x}'"
    return f"{x}"

def push_to_db(data, columns, connection_config, table_name, append=True, chunk_size=1000):
    """push data to database
    paramaters
        data: DataFrame or list.
              if list, then columns must be given
    """
    connection = get_connection(**connection_config)
    id_ = 0
    try:
        if isinstance(data, pandas.DataFrame):
            if append:
                data.to_sql(table_name, con=connection, if_exists="append", index=False)
            else:
                data.to_sql(table_name, con=connection, if_exists="replace", index=False)
        else:
            if not columns:
                return False
            with connection.cursor() as cursor:
                query = f"insert into {table_name}({','.join(columns)}) values(" + ",".join(["?"] * len(columns)) + ")"
                cursor.fast_executemany = True
                r = cursor.executemany(query, data)
            id_ = cursor.fetchone()[0]
            connection.commit()
    except Exception as e:
        print(e)
        return id_
    return id_

def get_data(query, connection_config, as_type="dict", key_column="", do_enumerate=True, columns=None):
    """execute the query and get the data
    parameters:
        as_type: dict => enumerates the results. the result is a dict of dicts
                 dataframe => returns pandas dataframe
                 list => returns list of lists
                 array => numpy array
        key_column: Used only if do_enumerate=False and as_type="dict". Takes precedence over do_enumerate
        do_enumerate: bool. Used only if as_type="dict"
        columns: used only for lists. If cannot get columns implicitly, pass a list of column names for the output explicitly
    """
    connection = get_connection(**connection_config)

    if as_type == "dict" and key_column.strip() != "":
        do_enumerate = False
    if as_type == "dict":
        with connection.cursor() as cursor:
            r = cursor.execute(query)
            columns = [i[0] for i in cursor.description]
            data = {} if do_enumerate or key_column else []
            for i, result in enumerate(r):
                if key_column:
                    col_index = columns.index(key_column)
                    obj = {columns[j]: result[j] for j in range(len(columns)) if columns[j] != key_column}
                    data[result[col_index]] = obj
                elif do_enumerate:
                    obj = {columns[j]: result[j] for j in range(len(columns))}
                    data[i] = obj
                else:
                    obj = {columns[j]: result[j] for j in range(len(columns))}
                    data.append(obj)

    elif as_type in ("list", "array"):
        with connection.cursor() as cursor:
            r = cursor.execute(query)
            data = []
            for row in r:
                data.append(row)
            if not columns:
                columns = [i[0] for i in cursor.description]
            if as_type == "array":
                data = np.array(data)
    else:
        data = pandas.query_as_df(query, con=connection)
        columns = data.columns.tolist()
    connection.commit()
    connection.close()
    return data, columns

def create_db_table(table_name, columns, sql_dtypes, connection_config):
    """create table in database using columns and sql data types. Data type is not inferred in this function
    """
    if len(columns) != len(sql_dtypes):
        return -1

    query = f"if object_id('{table_name}', 'U') is not null drop table {table_name}"
    query += ";" + f"create table {table_name} (" + ",".join([f"[{columns[i]}] {sql_dtypes[i]}" for i in range(len(columns))]) + ")"
    # logging.info(query)
    connection = get_connection(**connection_config)
    with connection.cursor() as cursor:
        for q in query.split(";"):
            result = cursor.execute(q)
    connection.commit()
    connection.close()
    return result
