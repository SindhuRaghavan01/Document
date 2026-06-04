import pandas


def get_exposed_limit_DR_df():
    fpath = "FinancialCalculator/ExposedLimitDRFile.csv"
    return pandas.read_csv(fpath)
