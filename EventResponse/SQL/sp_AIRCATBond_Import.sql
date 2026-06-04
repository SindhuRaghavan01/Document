

USE ARASystem
GO

/* 

The stored procedure will be used to transfer the CAT Bond data from Touchstone Re into
UI.

The user is expected to run analysis in TS re and create a new ILS instrument in UI before
transferring the data.

Inputs needed for the stored proc:

@ILSName = This should be the exact name used in the UI ILS Instrument
@CompanyName = Company name used ini TsRe UI after adding/importing bond.
@AnalysisName = Analysis name used in TsRe while running analysis. This will also be
available ini results section of TsRe UI.
@LinkedServer = Default LinkedServer where TsRe Production SQL Server
@Vendor = 'AIR' if it is default view without adjustments;
          'ARA View' if it is with adjustments
@Template = 'WSST' -- For default view;
            'STD -- if we do a long-term view for any analysis
@Perspective = 'Net' -- Default Perspective
               'Ground up'
               'Gross'
@CusipID = CusipID if the bond is already registered or available for trade.
           This will be used to retrieve market spreads from different brokers.


We need to add company or programID to identify the anlalysis.

*/


Create PROCEDURE [dbo].[sp_AIRCATBond_Import]
-- declare
@ActivityID int,
@LinkedServer NVARCHAR(200),
@Debug int = 0
as 
begin

Declare
@ILSName NVARCHAR(255),
@CompanyName NVARCHAR(500),
@AnalysisName NVARCHAR(500),
@Vendor NVARCHAR(50),
@Template NVARCHAR(50),
@Perspective NVARCHAR(50),
@CusipID NVARCHAR(50),
@ReProgramSID int = 0 -- Optional parameter. If this value is passed, that ProgramSID will be used.

Declare @CurrentAIRVersion NVARCHAR(10)
SET @CurrentAIRVersion = '10.0'
SET @Template = 'World All Perils (10k Warm SST Hurr, Time Dep Hybrid)'
SET @Perspective ='Net Loss'

DEclare 
@ILSInstrumentID int, 
-- @ReprogramSID int,
@AnalysisSID int,
@CipherID NVARCHAR(100),
@AnalysisTbl NVARCHAR(500),
@addExpense int,
@ResultSID int,
@Author nvarchar(50),
@CurrencyCode NVARCHAR(50),
@Sponsor_Name NVARCHAR(100),
@SponsorID Int,
@CoveredPerils nvarchar(100), --->>>>>
@Modeling_Agent NVARCHAR(100),
@RiskModelingFirm nvarchar(100),
@Collateral NVARCHAR(100),
@ReferenceRate NVARCHAR(100),
@TriggerType NVARCHAR(100),
@CoveredPeril NVARCHAR(100), --->>>>
@LossBasis NVARCHAR(100),
@OpenQuery NVARCHAR(500),
@Query NVARCHAR(max),
@InceptionDate Date,
@ExpirationDate Date,
@DateOfIndication Date,
@UpdateRiskSeasons Float,
@Today Date,
@rs FLOAT,
@rp Float,
@rrp Float,
@CatBondImportActivityID int,
@CurrencyExchangeFactor FLOAT

Declare @ILSPricingSummaryTableName NVARCHAR(510)
DECLARE @ExposureTableName NVARCHAR(510)

select @AnalysisName = a.AnalysisName,
       @CompanyName = a.SourceDbName,
       @Vendor = a.Vendor,
       @ILSInstrumentID = a.InstrumentID,
       @ILSName = b.InstrumentName,
       @ReProgramSID = AccPortSrNo
from ARAData_1.dbo.ActivityMonitor A
INNER Join ARAData_1.dbo.Instrument b
on a.InstrumentID = b.InstrumentID
where a.ActivityMonitorID = @ActivityID 

Set @ReProgramSID = Isnull(@ReProgramSID, 0)

Set @Query = ' Select 
      @AnalysisSID = AnalysisSID,
      @ResultSID = ResultSID,
      @Author = Author,
      @CurrencyCode = CurrencyCode'
    + ' FROM OpenQuery( ' 
    + QUOTENAME(@LinkedServer) 
    + ','
    + '''SELECT AnalysisSID,
                 ResultSID,
                 Author,
                 CurrencyCode'
    + ' FROM '
    + ' AIRProject.dbo.TAnalysis a'
    + ' Where AnalysisName = '''''
    + @AnalysisName
    + ''''''')'

IF @Debug = 1 Print @Query
EXECUTE sp_Executesql @Query,
N'@AnalysisSID int output, @ResultSID int output,
@Author nvarchar(255) Output, @CurrencyCode nvarchar(255) output',
@AnalysisSID = @AnalysisSID Output,
@ResultSID = @ResultSID Output,
@Author = @Author Output,
@CurrencyCode = @CurrencyCode output

IF @Debug = 1 Print @Query

SELECT @CurrencyExchangeFactor = a.ExcahngeFactor
FROM ARAData_1.dbo.CurrencyTable A 
Where A.CurrencyCode = @CurrencyCode


IF @Debug  = 1 PRINT 'AnalysisSID = ' + Cast(@AnalysisSID as nvarchar(255))

IF @ReProgramSID = 0 or @ReProgramSID is NULL
BEGIN
    SET @Query = ' Select @ReProgramSID = ReProgramSID '
               + ' FROM ' + QUOTENAME(@LinkedServer)
               + '.AIRResultRe.dbo.t'
               + Cast(@AnalysisSID as nvarchar(50))
               + '_Loss_ByReProgram'

    IF @DEBUG = 1 Print @Query

    EXEC sp_ExecuteSQL @Query, N'@ReProgramSID nvarchar(50) output',
                @ReProgramSID = @ReProgramSID Output
            
END

BEGIN TRY

-- This script creates a new activityID and JobID 

-- INSERT into ARAData_1.dbo.ActivityMonitor
-- (JobID,
-- UserId,
-- ARADbName,
-- AcctName,
-- AnalysisType,
-- ShapeId,
-- PEril,
-- JobStatus,
-- CurrencyCode,
-- SubmissionDateTime,
-- LastUpdationDateTime,
-- AnalysisName,
-- NumberOfTrials,
-- Vendor,
-- IsDefaultActivity,
-- [Version],
-- SourceDbName
-- )
-- select @ResultSID as JobID,
--        @Author as USerID,
--        '' as ARADbName,
--        '' as AcctName,
--        InstrumentID as AcctGrpID,
--        12 as AnalysisType,
--        0 as ShapeId,
--        0 as Peril,
--        128 as JobStatus,
--        @CurrencyCode as CurrencyCode,
--        GETDATE() as SubmissionDateTime, 
--        GETDate() as LastUpdationDateTime,
--        SUBSTRING(@AnalysisName, 1, 400) as AnalysisName,
--        10000 NumberOfTrials,
--        @Vendor as Vendor,
--        1 as IsDefaultAnalysis,
--        @CurrentAIRVersion as Version,
--        @LinkedServer as SourceDbName
-- from ARAData_1.dbo.Instrument
-- WHERE InstrumentID = @ILSInstrumentID


-- select * 
-- from ARAData_1.dbo.ActivityMonitor
-- where @ILSInstrumentID = AcctGrpID

-- SET @ActivityID = @@IDENTITY

IF @Debug = 1 Print ' Activity ID ' + Cast(@ActivityID as nvarchar(255)) + ' ----- Instrument ID = '
+ Cast(@ILSInstrumentID as nvarchar(255)) + ' ---- ReprogramSID = ' + Cast(@ReprogramSID as nvarchar(255))

-- Update ILSData
-- This query will update CATbond attribute from Tre(UI) Risk information section into
-- ILSData table Extention_Information_I, II, III is updated as 250,50,10 basis points
-- respectively for all CATbonds.

--  Select @ILSInstrumentID = Max(InstrumentID)
--  from ARAData_1.dbo.Instrument
--  where InstrumentName = @ILSName


drop table if exists #tmpAttribute
CReate Table #tmpAttribute
(
    REFERENCERate NVARCHAR(500),
    CoveredPerils NVARCHAR(500),
    Modeling_Agent NVARCHAR(500),
    RiskModelingFirm NVARCHAR(500),
    Collateral NVARCHAR(500),
    Sponsor_Name NVARCHAR(500),
    TriggerType NVARCHAR(500),
    LossBasis NVARCHAR(500),
    Extension_Information_I float,
    Extension_Information_II float,
    Extension_Information_III Float,
    Modeled_AAL_Base Float,
    Modeled_AAL_Sensitivity float,
    Modeled_attachProb_Base float,
    Modeled_attachProb_Sensitivty float,
    Modeled_ExhaustProb_Base float,
    Modeled_ExhaustProb_Sensitivity Float,
    Placement_Agent NVARCHAR(500)
)

Set @Query = 
' insert into #tmpAttribute 
(
    REFERENCERate,
    CoveredPerils,
    Modeling_Agent,
    RiskModelingFirm,
    Collateral,
    Sponsor_Name,
    TriggerType,
    LossBasis,
    Extension_Information_I,
    Extension_Information_II ,
    Extension_Information_III ,
    Modeled_AAL_Base ,
    Modeled_AAL_Sensitivity ,
    Modeled_attachProb_Base ,
    Modeled_attachProb_Sensitivty ,
    Modeled_ExhaustProb_Base ,
    Modeled_ExhaustProb_Sensitivity ,
    Placement_Agent 
)

select MAX(Case when CATBondAttributeName = ''Reference Rate''
                then CATBondAttributeValue END ) [ReferenceRate],
    MAX(Case when CATBondAttributeName = ''Covered Perils / Areas''
                then CATBondAttributeValue END ) [CoveredPerils],
    MAX(Case when CATBondAttributeName = ''Modeling Firm''
                then CATBondAttributeValue END ) [Modeling_Agent],
   MAX(Case when CATBondAttributeName = ''Modeling Firm''
                then CATBondAttributeValue END ) [RiskModelingFirm],
   MAX(Case when CATBondAttributeName = ''Collateral Account''
                then CATBondAttributeValue END ) [Collateral],
   MAX(Case when CATBondAttributeName = ''Sponsor''
                then CATBondAttributeValue END ) [Sponsor_Name],
   MAX(Case when CATBondAttributeName = ''Trigger Type''
                then CATBondAttributeValue END ) [TriggerType],
   MAX(Case when CATBondAttributeName = ''Cover Type''
                then CATBondAttributeValue END ) [LossBasis],
  250 Extension_Information_I,
  50 Extension_Information_II,
  10 Extension_Information_III,
  MAX(CASE When CATBondAttributeName = ''       Offering Expected Loss1 %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_AAL_Base ,
  MAX(CASE When CATBondAttributeName = ''       Offering Expected Loss2 %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_AAL_Sensitivity ,
  MAX(CASE When CATBondAttributeName = ''       Offering Attach Prob1. %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_attachProb_Base ,
  MAX(CASE When CATBondAttributeName = ''       Offering Attach Prob2. %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_attachProb_Sensitivty ,
  MAX(CASE When CATBondAttributeName = ''       Offering Exhaust Prob1. %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_ExhaustProb_Base ,
  MAX(CASE When CATBondAttributeName = ''       Offering Exhaust Prob2. %'' THEN
      CAST( CATBondAttributeValue as Float)/100 END) Modeled_ExhaustProb_Sensitivity ,
  MAX(CASE When CATBondAttributeName = ''Manager'' 
               THEN CATBondAttributeValue END) Placement_Agent            
from ' + QUOTENAME(@LinkedServer) +'.AIRUserSetting.dbo.tCatBondAttribute a 
Inner Join ' + QUOTENAME(@LinkedServer) +'.AIRExposureRe.dbo.tCatBondAttributeValue b 
on a.catBondAttributeSID = b.CatBondAtrributeSID 
where b.ReinsuranceProgramSID = ' + Cast(@ReProgramSID as nvarchar(50))

If @Debug = 1 PRINT @query

-- TODO = Fix This
-- TS re has incorrect mappings for CATBondAttribute table

Exec sp_ExecuteSQL @Query

-- Update Additional data from Touchstone Re
Update ARAData_1.dbo.Instrument
Set [Status] = 'On Hold',
-- SPonsorID = @SponsorID,
-- LossBasis = a.Lossbasis,
-- CoveredPerils = a.CoveredPerils,
ModelingAgent = a.Modeling_Agent,
RiskModelingFirm = a.Modeling_Agent,
Collateral = a.Collateral,
ReferenceRate = a.REFERENCERate,
-- TriggerType = a.TriggerType,
-- CUSIP_ID = @CusipID,
ExtensionInformationI = a.Extension_Information_I,
ExtensionInformationII = a.Extension_Information_II,
ExtensionInformationIII = a.Extension_Information_III,
ModeledAALBase = a.Modeled_AAL_Base,
ModeledAALSensitivity = a.Modeled_AAL_Sensitivity,
ModeledAttachProbBase = a.Modeled_attachProb_Base,
ModeledAttachProbSensitivity = a.Modeled_attachProb_Sensitivty,
ModeledExhaustProbBase = a.Modeled_ExhaustProb_Base,
ModeledExhaustProbSensitivity = a.Modeled_ExhaustProb_Sensitivity,
PlacementAgent = a.Placement_Agent
from #tmpAttribute a 
where InstrumentID = @ILSInstrumentID

-- Update CUSIP on UI from Ts Re if it's not entered
IF ltrim(rtrim(isnull(@CusipID, ''))) <> ''
BEGIN
    Update ARAData_1.dbo.Instrument 
    Set CUSIP = @CusipID
    where InstrumentID = @ILSInstrumentID
END
ELSE
BEGIN
    Select @CusipID = CUSIP
    FROM ARAData_1.dbo.Instrument
    where InstrumentID = @ILSInstrumentID
END

-- Populate Layer Loss Set
-- This query populates the YELT table using the ActivityID in ARA_Analysis DB as Analysis_ActivityID

Drop Table if exists #tmpLossSet
Declare @LossTbl NVARCHAR(500)
DEClare @AAL FLOAT

set @LossTbl = ' ARA_Analysis.dbo.Analysis_' + cast(@ActivityID as nvarchar(50))

set @Query = ' IF object_ID(''' + @LossBasis + ''', ''U'') is not null 
drop table ' + @LossTbl + ' ;Create Table ' + @LossTbl
+ '( [EventID] bigint, [Year] int, [Day] int, [Loss] Float, [GrossLoss] Float)'

IF @DEBUG = 1 Print @Query
Execute sp_executeSQL @Query
SET @Query = 
' insert into ' + @LossTbl + ' (EventID ,
Year ,
Day , 
Loss ,
GrossLoss)'
 + ' select b.EventID as EventID, YearID as Year, b.Day,
    TreatyFinalNetLossGrossParticipation / @CurrencyExchangeFactor as Loss, 0'
 + ' from ' + QUOTENAME(@LinkedServer) + '.AIRResultRe.dbo.t'
 + Cast(@AnalysisSID as nvarchar(50)) + '_Loss_ByReProgram a'
 + ' Inner Join ' + QUOTENAME(@LinkedServer)
 + '.AIREventInfo.dbo.vAIREventinfo_default b on a.ModelCode*1e7 + a.EventID = b.EVentID'
 + ' where a.CatalogTypeCode = ' + '''STC'''
 + ' and ReprogramSID = ' + Cast(@ReProgramSID as nvarchar(50))

 IF @DEbug = 1 select @LossTbl, @ActivityID, @AnalysisSID, @LinkedServer

 IF @Debug = 1 Print @Query 
 Execute sp_executeSQL @Query, N'@CurrencyExchangeFactor FLOAT',
 @CurrencyExchangeFactor = @CurrencyExchangeFactor


 Set @Query = 'SELETC @AAL = SUm(Loss)/1e4 from ' + @LossTbl
 Execute sp_executeSQL @Query, N'@AAL float Output', @AAL= @AAL Output

 If @AAL < 1
 Begin
    SET @Query = ' TRUNCATE TABLE ' + @LossTbl
    EXEC sp_ExecuteSQL @Query

    SET @Query = 
    ' insert into ' + @LossTbl + ' (EventID ,
    Year ,
    Day , 
    Loss ,
    GrossLoss)'
    + ' select b.EventID as EventID, YearID as Year, b.Day,
        TreatyFinalNetLoss100Participation / @CurrencyExchangeFactor as Loss, 0'
    + ' from ' + QUOTENAME(@LinkedServer) + '.AIRResultRe.dbo.t'
    + Cast(@AnalysisSID as nvarchar(50)) + '_Loss_ByReProgram a'
    + ' Inner Join ' + QUOTENAME(@LinkedServer)
    + '.AIREventInfo.dbo.vAIREventinfo_default b on a.ModelCode*1e7 + a.EventID = b.EVentID'
    + ' where a.CatalogTypeCode = ' + '''STC'''
    + ' and ReprogramSID = ' + Cast(@ReProgramSID as nvarchar(50))

    EXECute sp_ExecuteSQL @Query, N'@CurrencyExchangeFactor Float',
    @CurrencyExchangeFactor = @CurrencyExchangeFactor

End

-- Add AAL Summary by LOB, SubArea, Peril
-- No need to catch the error is caught within the sp and logged to activityErrorLog

Execute ARAData_1.dbo.arapl_populate_AALSummary
@ActivityID, @AnalysisSID, @LinkedServer, @ReProgramSID, @Debug = 0

Begin 

-- Update ExposureData if the data has not been imported before 
-- This Script returns aggregate exposure by Exposure Set, Peril, Year, Geography and LOB for given company.

SET @ExposureTableName = 'ARAData_1.dbo.[Instrument_' + Cast(@ILSInstrumentID as nvarchar(25))
+ '_' + ARAData_1.dbo.unfnGetAccountName(@ILSName)

IF OBJECT_ID(@ExposureTableName, 'U') is NULL
Begin
    Execute Update_CATBondExposure_Data @ILSInstrumentID, @ILSName, @CompanyName, @LinkedServer

if @Debug = 1 print  'Updated Company Exposure Data'

END

-- FOR ARAView
-- Create Terms with AnalysisSID = NULL, 
-- Delete the old ILS Terms and Insert them again

-- IF not Exists(Select * from ARAData_1.dbo.Terms a where a.InstrumentID = @ILSInstrumentID
-- and a.ActivityMonitorID is null)
-- Begin
DELETE ARAData_1.dbo.Terms
from ARAData_1.dbo.Terms
where InstrumentID = @ILSInstrumentID and 
@AnalysisSID is NULL

-- Setting @ActivityID = Null Since ActivityID is inserted as AnalysisSID in ILSTerms

EXECUTE ARApl_Populate_ILSTerms
@ILSInstrumentID = @ILSInstrumentID,
@ReProgramSID = @ReProgramSID,
@AnalysisSID = @AnalysisSID,
@LinkedServer = @LinkedServer,
@ActivityID = Null

IF @Debug = 1 Print 'Updated Terms with AnalysisSID = Null'
-- END

-- Update Terms
-- Information used to create a CAT Bond Layer structure is used to update Layer Terms.


EXECUTE ARApl_Populate_ILSTerms
@ILSInstrumentID = @ILSInstrumentID,
@ReProgramSID = @ReProgramSID,
@AnalysisSID = @AnalysisSID,
@LinkedServer = @LinkedServer,
@ActivityID  = @ActivityID

IF @Debug = 1 Print 'Updated Terms with AnalysisSID = ActivityID'


-- Update ILSResults
-- This will populate Risk Statistics using the above populated YELT

Execute ARAPL_Populate_ILSresults 
@ActivityID,
@Vendor,
[@Template],
@Perspective,
@addExpense

IF @Debug = 1 PRINT 'Updated ILSResults'

-- Updated ILSHistorical
-- This Script Returns the triggered historical Events using STC events available in Loss Sets.


select @ActivityID

Execure ARAPL_Populate_ILSHistorical 
@ActivityID

IF @Debug = 1 Print 'Updated Historical Events'

IF @Vendor = 'ARA'
Begin 

-- Update Risk Seasons/ Risk Period/Remaining Risk Period in ILSData
-- using sp ARApl_getRemainingAAL


select @DateOfIndication = Max(a.DateOfIndication)
from ARAData_1.dbo.CATBond_BrokerPricing_BrokerName A
where a.CUSIP = @CusipID

select @InceptionDate = InceptionDate,
       @ExpirationDate = ExpirationDate
from ARAData_1.dbo.Instrument a
Inner Join ARAData_1.dbo.Terms b
on a.InstrumentID = b.InstrumentID
and b.InuringSequenceNumber = 1
where a.InstrumentID = @ILSInstrumentID


SEt @Today = Case When @DateOfIndication > @InceptionDate
THEN @DateOfIndication else @InceptionDate END


DRop Table if exists #tmpSeasons
Create Table #tmpSeasons
(RiskSeasons float, remainingRiskPeriod float)

insert into #tmpSeasons(RiskSeasons, remainingRiskPeriod)
Execute ARApl_getRemainingAAL @Today, @ExpirationDate, @ActivityID

select @rs = isnull(riskSeasons, 0),
@rrp = remainingRiskPeriod
from #tmpSeasons


-- Get the risk Period of the deal 
-- we use this complicated expression to report the correct risk 
-- period to account for the leap year

Set @rp = 
datediff(year, 
        @InceptionDate,
         @ExpirationDate) 
- case when (DATEADD(year,
                     @InceptionDate,
                     @ExpirationDate),
            @InceptionDate) > @ExpirationDate 
 then 1 else 0 end + 
 cast(DATEDIFF(day,
               DATEADD(year,
                      datediff(year,
                                 @InceptionDate, @ExpirationDate) - 
 case when dateAdd(year, DATEDIFF(year, @InceptionDate, @ExpirationDate)
 > @ExpirationDate then 1 else 0 end , @InceptionDate), @ExpirationDate
 ) as float) / cast(datediff(day, dateadd(year, -1, @ExpirationDate),
  @ExpirationDate) as float)



  Update ARAData_1.dbo.Instrument
  SET RiskSeasons = @rs,
  RemainingRiskPeriod = @rrp, 
  RiskPeriod = @rp
  Where InstrumentID = @ILSInstrumentID
  
  IF @DEbug = 1 Print ' Updated Risk Period, Risk Seasons and Remaining Risk Period to ILSData'


-- Update Seasonally Adjusted Spread and Bid-Ask Spread in the broker Tables
-- This Query will calculate and update Seasonally Adjusted Spread and Bid-Ask Spread for all the weeks (available)



-- Create AAL By Month
Execute ARAPl_ANalysis.dbo.ARAPL_GenerateAALbyMonthtable
@ActivityID,
@ILSInstrumentID

IF @DEbug = 1 Print 'AAL by Month Table Generated'

End

-- Create ILSPricingSummaryTable
Set @ILSPricingSummaryTableName = 
'ARAPl_Analysis.dbo.[ILSPricingSummaryTableName_' + Cast(@activityID as nvarchar(255))
+ ']'

Set @Query = ' IF Object_ID(' + 
@ILSPricingSummaryTableName + ',''U'') is not null '
+ ' Drop table ' + @ILSPricingSummaryTableName

EXECUTE ARAPl_Analysis.dbo.ARApl_createISLPricingSummaryTable
@ILSPricingSummaryTableName


Update ARAData_1.dbo.ActivityMonitor 
SET IsDefaultActivity = 0 
where InstrumentID = @ILSInstrumentID
and ANalysisType = 12
and Vendor = @Vendor
and ActivityMonitorID <> @ActivityID

END TRY

BEGIN CATCH
UPdate ARAData_1.dbo.ActivityMonitor 
SET JobStatus = 256 
where ActivityMonitorID = @ActivityID


DECLARE @Message NVARCHAR(2000)
SET @Message = 'Could not import CAT Bond for ' + QUOTENAME(@ILSName, '''')
+ '(' + QUOTENAME(@CompanyName, '''') + ')'

EXEC ARADAta_1.dbo.usp_GetErrorInfo
@QueryName = 'update_Crown_CatBond',
@ActivityID = @ActivityID,
@Message = @Message

End CAtch



end