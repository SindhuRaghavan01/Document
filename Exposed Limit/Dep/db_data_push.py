from pandas import DataFrame

from DataLayer.db_utils import execute_sql_query, push_to_db, create_db_table


def save_location_results(loc_gr_df: DataFrame,
                          activity_id: int,
                          company_id: int,
                          connection_config: dict) -> int:
    """save net yelt to the database. If the table already exists, delete it.

    Parameters
    ----------
    loc_gr_df : DataFrame
        net loss yelt
    activity_id : int
        activity id
    company_id : int
        company id

    Returns
    -------
    int
        Flag to indicate if the insertion was successful
    """
    table_name = f"Analysis.dbo.Analysis_{activity_id}"
    query = f"drop table if exists {table_name}"
    execute_sql_query(query=query, connection_config=connection_config)

    # columns = '[EventID], [Year], [Day], [GrossLoss], [NetLoss], [RIP], [LossNetofRIP]'
    columns = loc_gr_df.columns.to_list()
    sql_data_types = ["INT", "INT", "INT", "FLOAT", "FLOAT", "FLOAT", "FLOAT"]
    create_db_table(table_name, columns, sql_data_types, connection_config)

    r = push_to_db(loc_gr_df.values.tolist(),
                   columns=columns,
                   connection_config=connection_config,
                   table_name=table_name,
                   append=True)
    return r


def save_policy_results(policy_gr_df: DataFrame,
                        activity_id: int,
                        company_id: int,
                        connection_config: dict) -> int:
    """save net yelt to the database. If the table already exists, delete it.

    Parameters
    ----------
    policy_gr_df : DataFrame
        net loss yelt
    activity_id : int
        activity id
    company_id : int
        company id

    Returns
    -------
    int
        Flag to indicate if the insertion was successful
    """
    table_name = f"Analysis.dbo.Analysis_{activity_id}"
    query = f"drop table if exists {table_name}"
    execute_sql_query(query=query, connection_config=connection_config)

    # columns = '[EventID], [Year], [Day], [GrossLoss], [NetLoss], [RIP], [LossNetofRIP]'
    columns = policy_gr_df.columns.to_list()
    sql_data_types = ["INT", "INT", "INT", "FLOAT", "FLOAT", "FLOAT", "FLOAT"]
    create_db_table(table_name, columns, sql_data_types, connection_config)

    r = push_to_db(policy_gr_df.values.tolist(),
                   columns=columns,
                   connection_config=connection_config,
                   table_name=table_name,
                   append=True)
    return r
