from pandas import DataFrame, merge, Series


def apply_policy_conditions_ded(pc_df: DataFrame,
                                pcc_df: DataFrame,
                                loc_dr_df: DataFrame):
    pc_df2 = pc_df.copy(deep=True)
    # pc_df2["PCApplicableGULoss"] = 0
    pc_df2["PCApplicableDed"] = 0
    # pc_df2["PCApplicableLimit"] = 0
    for i, row in pc_df2.iterrows():
        pcc_filtered = pcc_df[pcc_df["CONDITIONID"] == row.CONDITIONID]
        loc_filtered_df = merge(pcc_filtered, loc_dr_df, left_on="LOCID", right_on="Locid")

        # gu_loss = loc_filtered_df["SiteGULoss"].sum()

        # TODO: Add conditions for deductible type
        applicable_ded = max(loc_filtered_df["SiteApplicableDed"].sum(), row["DEDUCTIBLE1"])
        # TODO: Add conditions for limit type
        # applicable_loss = min(loc_filtered_df["SiteGULoss"].sum(), row["LIMIT1"])
        # pc_df2.at[i, "PCApplicableGULoss"] = gu_loss
        pc_df2.at[i, "PCApplicableDed"] = applicable_ded
        # pc_df2.at[i, "PCApplicableTotalLoss"] = applicable_loss
    return pc_df2


def apply_policy_conditions_back_ded(policy_ded_df: DataFrame,
                                     pc_ded_df: DataFrame):

    pc_ddf = merge(pc_ded_df,
                   policy_ded_df,
                   how="left",
                   left_on="POLICYID",
                   right_on="POLICYID")
    pc_ddf["PCEffectiveDed"] = pc_ddf["PCApplicableDed"] * pc_ddf["PCBackAllocFactor"]
    pc_ddf["PCEffectiveDed"] = pc_ddf.apply(get_pc_effective_ded, axis=1)
    return pc_ddf


def apply_policy_conditions_terms():
    pass


def get_pc_effective_ded(pc: Series):
    return 0
