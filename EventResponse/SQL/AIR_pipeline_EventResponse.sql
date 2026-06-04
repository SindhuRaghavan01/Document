	
Declare @InputParamtbl nvarchar(255)
Declare @guid nvarchar(255)
declare @EventID_byEventParam nvarchar(255)
declare @EventIDtbl nvarchar(255)
declare @IndustryLosstbl nvarchar(MAX)
declare @tblAIRCatalog nvarchar(MAX)
declare @portfolioLossTableName NVARCHAR(255)
declare @dealLossTableName NVARCHAR(255)
declare @ReportID int
declare @Fund int 
declare @Vendor nvarchar(255)
declare @ShapeTableName nvarchar(255)
declare @EventIDFiltertbl nvarchar(255)
declare @trackTableName nvarchar(255)
declare @ModelCode nvarchar(50)
declare @Query nvarchar(MAX)
declare @SelectQuery nvarchar(500) 
declare @GroupbyQuery nvarchar(500)
declare @id int
declare @max_id int
declare @parameter nvarchar(255)
declare @isTrackAnalysis BIT
declare @isModel27 BIT

DECLARE @Latitudes NVARCHAR(2048)
DECLARE @Longitudes NVARCHAR(2048)
DECLARE @Radii NVARCHAR(2048)

--- Track Analysis Inputs


SET @Latitudes = '19.72,23.56,27.67,32.04' --'25.6,27.6,29.9'
SET @Longitudes = '-82.97,-84.08,-83.45,-82.88' --'-83.4,-80.6,-77.3'
SET @Radii = '100,150,400,450'
Set @guid = 'DryRun_IanAdv013A_06162023' -- Unique identifier text
SET @isTrackAnalysis = 1 -- set 1 for track analysis 


Set @InputParamtbl = '[ARAPL_LiveEventData].[dbo].[AIR_inputParameterTable]' -- Input Parameter Table

SET @isModel27 = case when exists(select * from ARAPL_LiveEventData.dbo.AIR_inputParameterTable where Param = 'ModelCode' and ModelCode = 27)
 then 1 else 0 end

-- *******************************************
--- Track Analysis
-- *******************************************

-- Inputs: lat, long and other shape parameter
-- Filter Table: one of two possible tables (output table of Event params stored Procedure or RMS/AIR recommended EventId's)
-- Outputs: distinct Event ID's and Event Description filtered based on Tracks or shape params
-- Output table: save the outputs in table name containing shapes or tracks

if @isTrackAnalysis = 1
begin
	-- Create shapes
	EXEC [ARAPL_LiveEventData].[dbo].[AIR_CreateShapes_Circle] @Latitudes=@Latitudes, @Longitudes=@Longitudes, @Radii=@Radii, @guid=@guid
	SET @ShapeTableName = '[ARAPL_LiveEventData].[dbo].[AIR_Shapes_' + @guid + ']'

	-- EventID's by Tracks filtered by shapes
	-- run track analysis on track output
	exec [ARAPL_LiveEventData].[dbo].[AIR_GetLikeEventTracks] @InputShapeTableName=@ShapeTableName, @EventIDFilterTable=@EventID_byEventParam, @guid=@guid  --INPUT
	set @trackTableName = '[ARAPL_LiveEventData].[dbo].[AIR_LikeEvents_Tracks_' + @guid + ']'  --OUTPUT
	
	SET @Query = 'SELECT * FROM ' + @ShapeTableName
	EXEC (@Query)
end
else
begin
	set @trackTableName = @EventID_byEventParam
end

-- ****************************************
--- Event Params stored Procedure
-- ****************************************

-- Inputs: default input table with ModelCode 27
-- Filter Table: one of two possible tables (output table of Tracks Stored Procedure or RMS/AIR recommended EventId's)
-- Outputs: distinct Event ID's and Event Description filterd based on input params
-- Output Table: save the outputs in table name containing Eventparams or like events

EXEC [ARAPL_LiveEventData].[dbo].[AIR_GetLikeEventIds_byEventParams] @InputParamTableName = @InputParamtbl, @EventIDFiltertbl=@trackTableName, @guid = @guid
Set @EventID_byEventParam = '[ARAPL_LiveEventData].[dbo].[AIR_Filtered_EventId_' + @guid + ']'

--- AIR/RMS EventID's

-- Inputs:

-----------------------------
----- Distinct EventID's
-------------------------------

-- Input: one of three possibe tables
--		1. output table of Event params stoed Procedure
--		2. output table of Tracks Stored Procedure 
--		3. RMS/AIR recommended EventId's

set @EventIDtbl = @EventID_byEventParam --Input for the three following Queries


-- Questions:
--   1. How do you fetch the industryloss table name ?
--				-- AreaName
--				-- SubAreaName
--  
--  fetch 'ModelCode' from input Parameter table for Industry Loss table name

set @Query = ' select @ModelCode  = Value ' 
+ ' from ' + @InputParamtbl
+ ' where Param = ''ModelCode'''

EXEC sp_executeSQL @Query, N'@ModelCode INT OUT', @ModelCode=@ModelCode OUT

select @IndustryLosstbl = PIMCO_IndustryLossTable , @tblAIRCatalog = AIRCataLogTable
from AIREventInfo..TablesbyModelCode
where ModelCode = @ModelCode

print @IndustryLosstbl
print @tblAIRCatalog


-----------------------------------------------------------------------------------------------

--  2.  How do you include the Event Params for Events ?
--				-- will it come from input parameter table? No
				-- Will it come from input lookup parameter tabel? Yes


drop table if exists EventParams

create table EventParams( ID INT IDENTITY(1,1), Parameter nvarchar(255))
insert into EventParams ( Parameter)
Select  Param
from ARAPL_Configuration..LiveEventAnalysisOptions
where ModelCode = @ModelCode and Param_type = 'Param'

Set @SelectQuery = ''
Set @GroupbyQuery = ''

SET @id = 1
SELECT @max_id = max(ID) FROM EventParams

WHILE @id <= @max_id
BEGIN			
	select @parameter = Parameter
	from EventParams
	where ID = @id
	
	set @SelectQuery = @SelectQuery + ' , cat.' + cast(@parameter as nvarchar(255))
	set @GroupbyQuery = @GroupbyQuery + ' , cat.' + cast(@parameter as nvarchar(255))
	
	if @id = 1
		begin 
			set @SelectQuery = @SelectQuery + ' , Sum(IndustryLoss) as IndustryLoss'
		end	
	set @id = @id + 1 
END

set @SelectQuery = SUBSTRING(@SelectQuery, 4, 2000000)
set @GroupbyQuery = SUBSTRING(@GroupbyQuery, 4, 2000000)

print @SelectQuery
print @GroupbyQuery



-- ********************
--  Queries
-- ********************

--  (1) EventID, Event Description, Industry Loss

print(  ' select ' + @SelectQuery
	+ ' from ' + @EventIDtbl + ' a '
	+ ' inner join [AIREventInfo].dbo.' + @IndustryLosstbl + ' b '
	+ ' on a.EventID = b.EventID '
	+ ' inner join [AIREventInfo].dbo.' + @tblAIRCatalog + ' cat '
	+ ' on a.EventID = cat.EventID '
	+ ' Group by ' + @GroupbyQuery
	+ ' order by cat.EventID ')

exec(  ' select ' + @SelectQuery
	+ ' from ' + @EventIDtbl + ' a '
	+ ' inner join [AIREventInfo].dbo.' + @IndustryLosstbl + ' b '
	+ ' on a.EventID = b.EventID '
	+ ' inner join [AIREventInfo].dbo.' + @tblAIRCatalog + ' cat '
	+ ' on a.EventID = cat.EventID '
	+ ' Group by ' + @GroupbyQuery
	+ ' order by cat.EventID ')


--  (2) EventID, Event Description, AreaName, SubAreaName, Industry Loss, Event Params


Print( ' select  a.EventID, b.IndustryLoss, b.CountryName, b.AreaName, b.SubAreaName'
	+ ' from ' + @EventIDtbl + ' a '
	+ ' inner join [AIREventInfo].dbo.' + @IndustryLosstbl + ' b '
	+ ' on a.EventID = b.EventID ')

EXEC( ' select  a.EventID, b.IndustryLoss, b.CountryName, b.AreaName, b.SubAreaName'
	+ ' from ' + @EventIDtbl + ' a '
	+ ' inner join [AIREventInfo].dbo.' + @IndustryLosstbl + ' b '
	+ ' on a.EventID = b.EventID ')


--  (3) Portfolio Level Loss and Deal Level Loss

drop table if exists ReportID_byFund
create table ReportID_byFund( ID INT IDENTITY(1,1), ReportID int , Fund int , Vendor nvarchar(255), ReportDate Date)

insert into ReportID_byFund(ReportID, Fund, Vendor)
select ID, portfolioid, Vendor
from ARAPL_Configuration..BoundPortfolio 
where IsDefaultPortfolio = 1 and IsDeleted = 0 
and PortfolioID in (1,2,3,4) and Vendor = 'PIMCOVIEW'

set @id = 1 
select @max_id = Max(ID) from ReportID_byFund


while @id <= @max_id 
begin 

	select @ReportID = ReportID, @Fund= Fund, @Vendor = Vendor
	from ReportID_byFund
	where ID = @id

	print @Reportid	
	print @Fund
	print @Vendor

	exec [ARAPL_LiveEventData].dbo.AIR_GetDealLevelLosses @ReportID= @ReportID, @LikeEventsTableName =@EventIDtbl, @PortfolioLossTableName=@portfolioLossTableName  OUTPUT, @DealLevelLossTableName =@dealLossTableName OUTPUT, @guid =@guid
	set @portfolioLossTableName = '[ARAPL_LiveEventData].dbo.[AIR_PortfolioLoss_'+ @Vendor + '_Fund_' + cast(@Fund as nvarchar)  + '_' + @guid + ']'
	set @dealLossTableName = '[ARAPL_LiveEventData].dbo.[AIR_DealLevelLoss_' + @Vendor + '_Fund_' + cast(@Fund as nvarchar) + '_' + @guid + ']'

	
	exec('select  * from ' + @portfolioLossTableName )

	exec (' Select * from ' + @dealLossTableName )


	set @id = @id + 1
end


