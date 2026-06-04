import pandas


def get_location_dr(location_df: pandas.DataFrame,
                    damage_ratio_df: pandas.DataFrame):
    loc_cols = set(location_df.columns.tolist())
    dr_cols = set(damage_ratio_df.columns.tolist())
    merge_cols = list(loc_cols.intersection(dr_cols))
    return pandas.merge(location_df, damage_ratio_df, on=merge_cols)
