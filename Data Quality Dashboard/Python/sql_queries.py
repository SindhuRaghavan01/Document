sql_exposureSetSID = """
select CAST(evw.exposuresetSID as nvarchar(20)) as ExposureSetSID
    FROM AIRProject.dbo.tExposureView ev
    inner join AIRProject.dbo.tExposureViewDefinition evw on ev.ExposureViewSID = evw.ExposureViewSID
    inner join AIRProject.dbo.tDataSource d on d.DataSourceSID = evw.ExposureSetDataSourceSID
    WHERE ExposureViewName = '{exposureViewName}' and d.DataSourceName = '{ExpDBName}'
    -- FOR XML PATH('')), 1, 1, '')
"""
queries = {}

# insert into #TIVByGeoCodeMatchLevel(SponsorName, ZIPCode, GeoMatchLevel, TIV, LocCount)
queries["TIVByGeoCodeMatchLevel"] = """
    select '{SponsorName}' SponsorName, l.PostalCode as ZIPCode, GeoMatchLevel,SUM(TotalReplacementValue /ExchangeRate) as TIV, COUNT(DISTINCT l.locationSID) as LocCount, SUM(l.GrossArea) as FloorArea
     FROM [{ExpDBName}]..tContract c
     inner join [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID and c.exposuresetSID = l.exposuresetSID
     inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
    inner join [AIRReference].[dbo].[tGeoMatchLevel] geo on geo.GeoMatchLevelCode = l.geomatchlevelcode
      WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND
      CurrencyExchangeRateSetSID = 5
     GROUP BY l.PostalCode, GeoMatchLevel
     ORDER BY 2
     """

# insert into #TIVByStateCountry(SponsorName, ZIPCode, CountryName, CountryCode, StateName, StateCode, County, TIV, LocCount)
queries["TIVByStateCountry"] = """select '{SponsorName}' SponsorName, l.PostalCode as ZIPCode, CountryName ,CountryCode,AreaName as Statename,AReaCode as statecode,SubAreaName as county, SUM(TotalReplacementValue /ExchangeRate) as TIV,COUNT(DISTINCT l.locationSID) as LocCount, SUM(l.GrossArea) as FloorArea
     FROM [{ExpDBName}]..tContract c
     inner join [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID and c.exposuresetSID = l.exposuresetSID
     inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
    inner join [AIRReference].[dbo].[tGeoMatchLevel] geo on geo.GeoMatchLevelCode = l.geomatchlevelcode
      WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND
      CurrencyExchangeRateSetSID = 5
    GROUP BY l.PostalCode, CountryName ,CountryCode ,AreaName,AReaCode,SubAreaName
      ORDER BY 6"""


# insert into #TIVByOccupancyCategory(SponsorName,  ZIPCode, OccupancyCategory, TIV, LocCount)
queries["TIVByOccupancyCategory"] = """
    SELECT ' {SponsorName}' SponsorName, 
                l.PostalCode as ZIPCode,
                OccupancyCategory,
                SUM(TotalReplacementValue /ExchangeRate) as TIV,
                COUNT(DISTINCT l.locationSID) as LocCount, SUM(l.GrossArea) as FloorArea
    from  [{ExpDBName}]..tContract c
    inner join [{ExpDBName}]..tLocation l 
        on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
        on curr.CurrencyCode = l.CurrencyCode
    inner join AIRReference..tAIROccupancy occ 
        on occ.AIROccupancyCode=l.UserOccupancyCode 
    WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
         CurrencyExchangeRateSetSID = 5 
    GROUP BY l.PostalCode, OccupancyCategory 
    ORDER BY 2,3"""

# insert into #TIVByConstructionCategory(SponsorName, ZIPCode, ConstructionCategory, TIV, LocCount)
 
queries["TIVByConstructionCategory"] = """
    SELECT ' {SponsorName} ' SponsorName,
            l.PostalCode as ZIPCode,
            ConstructionCategory,
            SUM(TotalReplacementValue /ExchangeRate) as TIV,
            COUNT(DISTINCT l.locationSID) as LocCount,
            SUM(l.GrossArea) as FloorArea
    from   [{ExpDBName}]..tContract c 
    inner join  [{ExpDBName}]..tLocation l 
        on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
        on curr.CurrencyCode = l.CurrencyCode
    inner join AIRReference..tAIRConstruction concode 
        on concode.AIRConstructionCode=l.UserConstructionCodeA
    WHERE c.exposureSetSID in ({ListofExposureSetSID})  
        AND CurrencyExchangeRateSetSID = 5 
    GROUP BY l.PostalCode, ConstructionCategory 
    ORDER BY 2"""

# insert into #TIVByYearBuiltRange(SponsorName, ZIPCode, YearBuiltRange, TIV, LocCount)
  
queries["TIVByYearBuiltRange"] = """
    select '{SponsorName}' SponsorName,
            PostalCode as ZIPCode,
            YearBuiltRange, 
            SUM(TIV) as TIV,
            SUM(LocCount) as LocCount,
            SUM(FloorArea) as FloorArea
    from
    (SELECT l.PostalCode, 
            CASE WHEN ISNULL(Yearbuilt,0) = 0 THEN 'Unknown'
                 WHEN Yearbuilt >0 AND YEARbuilt < 1995 THEN '<1995'
                 WHEN Yearbuilt >1995 AND YEARbuilt <=2001 THEN '1995-2001'
                 WHEN Yearbuilt >2001 AND YEARbuilt <=2012 THEN '2002-2012'
                 WHEN Yearbuilt >2013 AND YEARbuilt <=2050 THEN '>=2013'
                 ELSE 'Unknown' END as YearBuiltRange,
            TotalReplacementValue /ExchangeRate TIV,
            l.GrossArea as FloorArea,
            1 as LocCount
    FROM  [{ExpDBName}]..tContract c
    inner join  [{ExpDBName}]..tLocation l
        on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
        on curr.CurrencyCode = l.CurrencyCode
    WHERE c.exposureSetSID in ({ListofExposureSetSID}) AND 
            CurrencyExchangeRateSetSID = 5 )t1
    GROUP BY PostalCode, YearBuiltRange 
    ORDER BY 2 """


# insert into #TIVByNumberOfStories(SponsorName,  ZIPCode,NoOfStoriesRange, TIV, LocCount)

queries["TIVByNumberOfStories"] = """
    SELECT ' {SponsorName} ' SponsorName,
            PostalCode as ZIPCode,
            NoofStories,
            SUM(TIV) as TIV,
            SUM(LocCount) as LocCount,
            SUM(FloorArea) as FloorArea
    FROM
    (SELECT l.PostalCode, CASE WHEN ISNULL(Stories,0) = 0 THEN 'Unknown'
    WHEN Stories >0 AND Stories <=1 THEN '1 Storey'
    WHEN Stories >1 AND Stories <=3 THEN '2-3 Storey'
    WHEN Stories >3 AND Stories <=7 THEN '4-7 Storey'
    WHEN Stories >8 THEN '>=8 Storey'
    ELSE 'Unknown' END as NoofStories,
    TotalReplacementValue /ExchangeRate TIV,
    l.GrossArea as FloorArea,
    1 as LocCount
    FROM   [{ExpDBName}]..tContract c
    inner join  [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID  AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
    WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
    CurrencyExchangeRateSetSID = 5 
    )t1
    GROUP BY PostalCode, NoofStories 
    ORDER BY 2 """


# insert into #TIVByFloorArea(SponsorName, ZIPCode, AreaUnit, FloorAreaRange, TIV, LocCount)
    
queries["TIVByFloorArea"] = """    
    SELECT '{SponsorName}' SponsorName, PostalCode as ZIPCode,  Areaunit, FloorAreaRange, SUM(TIV) as TIV, SUM(LocCount) as LocCount, SUM(FloorArea) as FloorArea
    from
    (SELECT l.PostalCode, GrossAreaUnitCode AS Areaunit, 
    CASE WHEN ISNULL(GrossArea, 0) = 0 THEN 'Unknown'
    WHEN GrossArea >0 AND    GrossArea <=1500 THEN '1-1500'
    WHEN GrossArea >1500 AND GrossArea <=2500 THEN '1500-2500'
    WHEN GrossArea >2500 AND GrossArea <=5000 THEN '2500-5000'
    WHEN GrossArea >5000 AND GrossArea <=10000 THEN '5000-10000'
    WHEN GrossArea >10000 THEN '10000'
    END As FloorAreaRange,
    TotalReplacementValue /ExchangeRate TIV,
    l.GrossArea as FloorArea,
    1 as LocCount
    FROM   [{ExpDBName}]..tContract c
    inner join  [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID  AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
      WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
    CurrencyExchangeRateSetSID = 5 
    )t1
    GROUP BY PostalCode, Areaunit,FloorAreaRange
    ORDER BY 2,4"""


# insert into #BuildingValueperSQFTbyFLCounty(SponsorName, ZIPCode, FLCounty , BuildingValue , LocCount , FloorArea , BuildingValueperSQFT)
	
queries["BuildingValueperSQFTbyFLCounty"] = """
select '{SponsorName}' SponsorName,
        l.CountryName as Country,
        l.AreaName as AreaName,
        l.PostalCode as ZIPCode, 
        l.SubAreaName as county,
        SUM(TotalReplacementValue /ExchangeRate) as BuildingValue,
        COUNT(DISTINCT l.locationSID) as LocCount,
        SUM(l.GrossArea) as FloorArea,
        AVG(case when isnull(l.GrossArea, 0) =0 then 0 
                 else l.ReplacementValueA/ExchangeRate/l.GrossArea end)
            as BuildingValueperSQFT
    from   [{ExpDBName}]..tContract c
    inner join  [{ExpDBName}]..tLocation l 
    on l.ContractSID = c.ContractSID AND  l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
    on curr.CurrencyCode = l.CurrencyCode 
    inner join [AIRReference].[dbo].[tGeoMatchLevel] geo 
    on geo.GeoMatchLevelCode = l.geomatchlevelcode
    WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
    CurrencyExchangeRateSetSID = 5 
	GROUP BY l.PostalCode, CountryName ,CountryCode ,AreaName,AReaCode,SubAreaName
    ORDER BY 1"""

# insert into #TIVbyRoofAge(SponsorName, ZIPCode, State, RoofAge, TIV, LocCount)

queries["TIVbyRoofAge"] = """
     SELECT '{SponsorName}' SponsorName, PostalCode as ZIPCode, t.AreaName, t.RoofAge, SUM(t.TIV) as TIV,  sum(LocCount) as LocCount, sum(FloorArea) as FloorArea
     from  (
     SELECT l.PostalCode,
             AreaName, 
             TotalReplacementValue /ExchangeRate TIV,
             l.GrossArea as FloorArea,
             Case WHEN RoofYearBuilt = 0 THEN 'Unknown'
             WHEN (2022 - RoofYearBuilt) <= 5 THEN '0 to 5'
             WHEN (2022 - RoofYearBuilt) <= 10 AND (2022 - RoofYearBuilt) > 5  THEN  '5 to 10'
             WHEN (2022 - RoofYearBuilt) <= 15 AND  (2022 - RoofYearBuilt) > 10 THEN '10 to 15'
             WHEN (2022 - RoofYearBuilt) <= 20 AND  (2022 - RoofYearBuilt) > 15 THEN '15 to 20'
             WHEN (2022 - RoofYearBuilt) > 20 AND RoofYearBuilt <> 0 THEN '>20'
     	END as RoofAge, 1 as LocCount
      FROM  [{ExpDBName}]..tContract c
         inner join  [{ExpDBName}]..tLocation l 
         on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
         inner join  [{ExpDBName}]..tLocFeature lf on l.LocationSID = lf.LocationSID
         inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
         on curr.CurrencyCode = l.CurrencyCode
         WHERE c.exposureSetSID in ({ListofExposureSetSID})   AND 
         CurrencyExchangeRateSetSID = 5 
     )t
     group by PostalCode, t.AreaName, t.RoofAge
     order by t.AreaName, t.RoofAge DESC """



# insert into #TIVByOccupancyCode(SponsorName,ZIPCode, UserOccupancyCode, TIV, LocCount)
    
queries["TIVByOccupancyCode"] = """
    SELECT '{SponsorName}' SponsorName, 
            l.PostalCode as ZIPCode,
            l.UserOccupancyCode,
            SUM(TotalReplacementValue /ExchangeRate) as TIV, 
            COUNT(DISTINCT l.locationSID) as LocCount,
            SUM(l.GrossArea) as FloorArea
    from  [{ExpDBName}]..tContract c
    inner join [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
    inner join AIRReference..tAIROccupancy occ on occ.AIROccupancyCode=l.UserOccupancyCode 
    --inner join AIRUserSetting ..tCurrencyExchangeRateSet curr on curr.BaseCurrencyCode = l.CurrencyCode 
    -- inner join tLocTerm lc on lc.LocationSID = l.LocationSID
    -- inner join tLocFeature lf on lf.LocationSID = l.LocationSID
      WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
     CurrencyExchangeRateSetSID = 5 
     GROUP BY l.PostalCode, l.UserOccupancyCode  
     ORDER BY 2"""



# insert into #TIVByConstructionTypeCode(SponsorName, ZIPCode, UserConstructionType, TIV, LocCount)
queries["TIVByConstructionTypeCode"] = """SELECT '{SponsorName}' SponsorName, l.PostalCode as ZIPCode, l.UserConstructionCodeA , SUM(TotalReplacementValue /ExchangeRate) as TIV, COUNT(DISTINCT l.locationSID) as LocCount, sum(l.GrossArea) as FloorArea
    from   [{ExpDBName}]..tContract c 
    inner join  [{ExpDBName}]..tLocation l on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
    inner join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr on curr.CurrencyCode = l.CurrencyCode
    inner join AIRReference..tAIRConstruction concode on concode.AIRConstructionCode=l.UserConstructionCodeA

    --inner join AIRUserSetting ..tCurrencyExchangeRateSet curr on curr.BaseCurrencyCode = l.CurrencyCode 
    -- inner join tLocTerm lc on lc.LocationSID = l.LocationSID
    -- inner join tLocFeature lf on lf.LocationSID = l.LocationSID
      WHERE c.exposureSetSID in ({ListofExposureSetSID})  AND 
     CurrencyExchangeRateSetSID = 5 
     GROUP BY l.PostalCode, l.UserConstructionCodeA 
     ORDER BY 2"""


#insert into #TotalTIV(SponsorName, TotalTIV, TotalBuildingValue, TotalLocCount)
queries["TotalTIV"] = """SELECT '{SponsorName}' SponsorName, SUM(TotalReplacementValue /ExchangeRate) as TIV,
                         SUM(ReplacementValueA / ExchangeRate) as BuildingValue,
                         COUNT(DISTINCT l.LocationSID) as LocCount,
                         SUM(l.GrossArea) as FloorArea
                         FROM [{ExpDBName}]..tContract c 
                         inner join  [{ExpDBName}]..tLocation l 
                         on l.ContractSID = c.ContractSID AND l.ExposureSetSID = c.ExposureSetSID
                         left join [AIRUserSetting].[dbo].[tCurrencyExchangeRateSetConversion] curr 
                         on curr.CurrencyCode = l.CurrencyCode
                         WHERE c.exposureSetSID in ({ListofExposureSetSID})  
                         AND CurrencyExchangeRateSetSID = 5"""


