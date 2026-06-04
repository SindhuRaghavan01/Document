from pandas import DataFrame


def apply_location_ded(loc_dr_df: DataFrame) -> DataFrame:
    loss_df = loc_dr_df.copy(deep=True)
    for field in ["Bldg", "ApptStruct", "Content", "BI"]:
        loss_field = field + "GULoss"
        value_field = field + "Value"
        ded_field = field + "DeductAmt"
        # limit_field = field + "Limit"
        appl_ded_field = field + "ApplicableDed"
        # appl_limit_field = field + "ApplicableLimit"
        loss_df[loss_field] = loss_df[value_field] * loss_df["DR"] / 100
        loss_df[appl_ded_field] = loss_df.apply(lambda x: min(x[ded_field], x[loss_field]), axis=1)
        # loss_df[appl_limit_field] = loss_df.apply(lambda x: min(x[limit_field], x[loss_field]), axis=1)

    for calc in ["GULoss", "ApplicableDed"]:
        loss_df["Combined" + calc] = (loss_df["Bldg" + calc] + loss_df["ApptStruct" + calc] +
                                      loss_df["Content" + calc])
        loss_df["Site" + calc] = (loss_df["Bldg" + calc] + loss_df["ApptStruct" + calc] +
                                  loss_df["Content" + calc] + loss_df["BI" + calc])
    return loss_df


def apply_location_back_ded():
    pass


def apply_location_terms():
    pass


def apply_location_back_gr():
    pass
