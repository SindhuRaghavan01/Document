
import urllib

from sqlalchemy import create_engine
from sqlalchemy import (
    Table,
    Column,
    Integer,
    String,
    MetaData,
    select,
    text,
    and_,
    or_,
    func,
)


DRIVER_NAME = "{SQL Server}"
SERVER_NAME = "sqlprdair"
DB_NAME = "ARAPL_Configuration"
USERNAME = "ARAPL_USR"
PASSWORD = "Arapl@123"


def get_engine(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD, echo=False):
    """
    create sql connection
    """
    param_string = "DRIVER={};SERVER={};DATABASE={};UID={};PWD={}"
    params = urllib.parse.quote_plus(
        param_string.format(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD)
    )
    conn_string = "mssql+pyodbc:///?odbc_connect={}".format(params)
    engine = create_engine(conn_string, echo=echo)
    return engine


engine = get_engine(DRIVER_NAME, SERVER_NAME, DB_NAME, USERNAME, PASSWORD)


from sqlalchemy.orm import sessionmaker, aliased


Session = sessionmaker(bind=engine)
session = Session()


metadata = MetaData(bind=engine)

def push_to_db(df, engine, table_name, append=True):
    try:
        if append:
            df.to_sql(table_name, con=engine, if_exists="append", index=False)
        else:
            df.to_sql(table_name, con=engine, if_exists="replace", index=False)
    except Exception as e:
        print(e)
        return False
    return True

