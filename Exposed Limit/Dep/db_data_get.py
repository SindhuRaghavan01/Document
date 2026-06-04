from DataLayer.db_utils import get_data


def get_location_data(base_table_name: str,
                      peril_id: int,
                      connection_config: dict,
                      company_id: int,
                      filter_string: str = None):
    peril_table_name = f"{base_table_name}_{peril_id}"
    query = """select loc.Locid, loc.Accountid,
loc.CountryCode, loc.State, loc.StateCode, loc.CRESTA, loc.County, loc.CountyCode,
loc.PostalCode, loc.City, loc.StreetName, loc.Latitude, loc.Longitude,
loc.BldgClass, loc.OccType, loc.NumStories, loc.FLOORAREA,
loc.BldgValue, loc.ApptStructValue, loc.ContentValue, loc.BIValue, loc.Tiv,
locperil.LimitType, locperil.BldgLimit, locperil.ApptStructLimit, locperil.ContentLimit, locperil.BILimit,
locperil.DeductibleType, locperil.BldgDeductAmt, locperil.ApptStructDedAmt ApptStructDeductAmt,
locperil.ContentDeductAmt, locperil.BIDeductAmt
from [Data].dbo.{base_table_name} loc
inner join [Data].dbo.{peril_table_name} locperil
on loc.Locid = locperil.Locid
"""
    query = query.format(base_table_name=base_table_name, peril_table_name=peril_table_name)
    if filter_string:
        query = query + f" where {filter_string}"
    location_data, _ = get_data(query=query,
                                connection_config=connection_config,
                                as_type="DataFrame")
    return location_data


def get_policy_conditions_criteria_data(base_table_name: str,
                                        peril_id: int,
                                        connection_config: dict,
                                        company_id: int
                                        ):
    pcc_name = f"{base_table_name}_PolicyConditionsCriteria"
    peril_table_name = f"{base_table_name}_{peril_id}"
    query = """SELECT a.[LOCID], a.[CONDITIONID], a.[PerilID], a.[INCLUDED]
    FROM [Data].[dbo].{pcc_name} a
    inner join (select distinct a.locid
                from [Data].[dbo].{base_table_name} a
                inner join [Data].[dbo].{peril_table_name} b
                on a.locid = b.locid
    )b
    on a.locid = b.locid
    WHERE PerilID = {peril_id}"""
    query = query.format(pcc_name=pcc_name,
                         peril_id=peril_id,
                         base_table_name=base_table_name,
                         peril_table_name=peril_table_name)
    pcc_data, _ = get_data(query=query,
                           connection_config=connection_config,
                           as_type="DataFrame")
    return pcc_data


def get_policy_conditions_data(base_table_name: str,
                               peril_id: int,
                               connection_config: dict,
                               company_id: int
                               ):
    pc_name = f"{base_table_name}_PolicyConditions"
    pcc_name = f"{base_table_name}_PolicyConditionsCriteria"
    peril_table_name = f"{base_table_name}_{peril_id}"
    query = """SELECT a.[CONDITIONID], a.[POLICYID], a.[CONDITIONNAME], a.[OcclimitType],
              a.[LIMIT1], a.[LIMIT2], a.[LIMIT3], a.[LIMIT4], a.[DEDUCTIBLE1],
              a.[DEDUCTIBLE2], a.[DEDUCTIBLETYPE], a.[Attach1], a.[Attach2], a.[Attach3],
              a.[Attach4], a.[InuringSequenceNumber], a.[CONDITIONTYPE],
              a.[PARENTCONDITIONID], a.[POLICYCONDITION], a.[PerilID]
    FROM [Data].[dbo].[{pc_name}] a
    WHERE PerilID = {peril_id}
    """
    query = query.format(pc_name=pc_name,
                         peril_id=peril_id,
                         pcc_name=pcc_name,
                         base_table_name=base_table_name,
                         peril_table_name=peril_table_name)
    pc_data, _ = get_data(query=query,
                          connection_config=connection_config,
                          as_type="DataFrame")
    return pc_data


def get_policy_data(base_table_name: str,
                    peril_id: int,
                    connection_config: dict,
                    company_id: int):
    policy_table_name = f"{base_table_name}_Policy"
    query = """SELECT a.[POLICYID], a.[POLICYNUMBER], a.[INCEPTIONDATE],
      a.[EXPIRYDATE], a.[ATTACHMENTPOINT], a.[COMPANYSHARE], a.[LAYERAMOUNT],
        a.[DeductibleTypeCode], a.[Deductible1], a.[Deductible2], a.[Peril],
          a.[CurrencyCode], a.[AccountID]
    FROM [Data].[dbo].[{policy_table_name}] a
    WHERE PerilID = {peril_id}
    """
    query = query.format(policy_table_name=policy_table_name,
                         peril_id=peril_id)
    pc_data, _ = get_data(query=query,
                          connection_config=connection_config,
                          as_type="DataFrame")
    return pc_data
