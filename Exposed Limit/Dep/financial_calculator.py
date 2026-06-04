import pandas

from DataLayer.db_utils import get_connection_config

from .db_data_get import (get_location_data,
                          get_policy_conditions_criteria_data,
                          get_policy_conditions_data, get_policy_data)
from .db_data_push import save_location_results, save_policy_results
from .exposed_limit_utils import get_exposed_limit_DR_df
from .location_calculations import (apply_location_back_ded,
                                    apply_location_back_gr, apply_location_ded,
                                    apply_location_terms)
from .location_intersection import get_location_dr
from .policy_calculations import apply_policy_ded, apply_policy_terms
from .policy_condition_calculations import (apply_policy_conditions_back_ded,
                                            apply_policy_conditions_ded,
                                            apply_policy_conditions_terms)


def financial_calculator(activity_id: int, company_id: int):
    """run the financial calculator for the activity id and save the results in the database

    Parameters
    ----------
    activity_id : int
        current activity id
    company_id : int
        company id
    """
    connection_config = get_connection_config()

    # ------ get damage ratio data
    damage_ratio_df = get_exposed_limit_DR_df()
    # ------ get base exposure information
    base_table_name = "Bound_201905_BHSI_Exposure_Bound_Surge_Only_May_2019"
    peril_id = 2

    # company_id = 1
    # ------ get location data
    filter_string = "StateCode in ('TX', 'AL', 'MS', 'LA')"
    location_df = get_location_data(base_table_name,
                                    peril_id,
                                    connection_config,
                                    company_id,
                                    filter_string)

    # ------ get policy conditions data
    pc_df = get_policy_conditions_data(base_table_name=base_table_name,
                                       peril_id=peril_id,
                                       connection_config=connection_config,
                                       company_id=company_id)

    # ------ get policy conditions criteria
    pcc_df = get_policy_conditions_criteria_data(base_table_name=base_table_name,
                                                 peril_id=peril_id,
                                                 connection_config=connection_config,
                                                 company_id=company_id)

    # ------ get policy data
    policy_df = get_policy_data(base_table_name=base_table_name,
                                peril_id=peril_id,
                                connection_config=connection_config,
                                company_id=company_id)

    # ------ apply damage ratio to the location data
    # gives GU losses
    loc_dr_df = get_location_dr(location_df=location_df, damage_ratio_df=damage_ratio_df)
    n_rows = loc_dr_df.shape[0]
    batch_size = 10000
    t = []

    # ------ apply location terms
    # gives location deductibles
    for i in range(n_rows // batch_size + 1):
        t.append(apply_location_ded(loc_dr_df.iloc[i * batch_size: (i + 1) * batch_size].reset_index(drop=True)))
    loc_ded_df = pandas.concat(t)

    # ------ apply policy conditions using locations, policy conditions and policy condition criteria
    # gives policy condition deductibles
    pcc_ded_df = apply_policy_conditions_ded(pc_df=pc_df,
                                             pcc_df=pcc_df,
                                             loc_dr_df=loc_ded_df)

    # ------ apply policy terms
    # gives policy deductibles
    policy_ded_df = apply_policy_ded(policy_df=policy_df,
                                     pcc_ded_df=pcc_ded_df,
                                     loc_dr_df=loc_dr_df)

    # ------ calculate policy condition effective deductible
    # gives policy condition effective deductible
    pcc_back_ded_df = apply_policy_conditions_back_ded(policy_ded_df=policy_ded_df,
                                                       pcc_ded_df=pcc_ded_df)

    # ------ calculate location effective deductible
    # back-allocate total location deductible to policy condition effective deductible
    loc_back_ded_df = apply_location_back_ded(pcc_back_ded_df=pcc_back_ded_df,
                                              loc_ded_df=loc_ded_df)

    # ------ calculate location GR' loss
    # apply loc limits and effective deductible
    loc_gr_df = apply_location_terms(loc_back_ded_df)

    # ------ calculate policy condition GR loss
    # add up location GR losses and apply condition limits and effective deductibles
    pcc_gr_df = apply_policy_conditions_terms(pcc_back_ded_df=pcc_back_ded_df,
                                              loc_gr_df=loc_gr_df)

    # ------ calculate policy final GR loss
    # add up condition and location GR losses to get policy GR loss
    # apply policy limits and effective deductibles
    policy_gr_df = apply_policy_terms(policy_ded_df=policy_ded_df,
                                      pcc_gr_df=pcc_gr_df)

    # ------ calculate location final GR loss
    # back-allocate policy GR loss to location loss
    loc_final_gr_df = apply_location_back_gr(policy_gr_df=policy_gr_df,
                                             loc_back_ded_df=loc_back_ded_df)

    # ------ save location final GR losses
    save_location_results(loc_gr_df=loc_final_gr_df,
                          activity_id=activity_id,
                          company_id=company_id,
                          connection_config=connection_config)

    # ------ save policy final GR losses
    save_policy_results(loc_gr_df=loc_final_gr_df,
                        activity_id=activity_id,
                        company_id=company_id,
                        connection_config=connection_config)
    # ------ create executive summary?
