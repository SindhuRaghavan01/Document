from pandas import DataFrame, merge, Series


def apply_policy_ded(policy_df: DataFrame,
                     pcc_ded_df: DataFrame,
                     loc_dr_df: DataFrame):
    """_summary_

    Parameters
    ----------
    pc_df : DataFrame
        _description_
    pcc_df : DataFrame
        _description_
    loc_dr_df : DataFrame
        _description_
    """
    pcc_grouped = (pcc_ded_df[["POLICYID", "PCApplicableDed"]]
                   .groupby(by="POLICYID")["PCApplicableDed"]
                   .sum()
                   .reset_index()
                   )
    policy_ddf = merge(policy_df,
                       pcc_grouped,
                       how="left",
                       left_on="POLICYID",
                       right_on="PolicyID")
    policy_ddf.fillna(values={"PCApplicableDed": 0}, inplace=True)
    # TODO: Add location applicable deductible for locations not in conditions
    policy_ddf["OutLocApplicableDed"] = 0

    policy_ddf["TotalApplicableDed"] = policy_ddf.OutLocApplicableDed + policy_ddf.ApplicableDed

    # calculate effective policy deductible
    policy_ddf["PolicyEffectiveDed"] = policy_ddf.apply(get_policy_eff_ded, axis=1)

    policy_ddf["PCBackAllocFactor"] = policy_ddf.apply(get_pc_backalloc_factor, axis=1)
    # policy_ddf.fillna(values={"BackAllocFactor": 0}, inplace=True)

    return policy_ddf


def apply_policy_terms():
    pass


def get_policy_eff_ded(policy: Series):
    """get effective deductible for a single policy passed as a pandas.Series object

    Parameters
    ----------
    policy_series : Series
        _description_
    """
    if policy.DeductibleTypeCode == "N":
        return policy.TotalApplicableDed
    if policy.DeductibleTypeCode == "MI":
        return max(policy.Deductible1, policy.TotalApplicableDed)
    if policy.DeductibleTypeCode == "MA":
        return min(policy.Deductible1, policy.TotalApplicableDed)
    if policy.DeductibleTypeCode == "MM":
        return min(max(policy.Deductible1, policy.TotalApplicableDed), policy.Deductible2)
    return policy.Deductible1


def get_pc_backalloc_factor(policy):
    if policy.TotalApplicableDed != 0:
        return policy.EffectiveDeductible / policy.TotalApplicableDed
    return 1
