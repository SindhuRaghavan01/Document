SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO 
ALTER PROCEDURE [dbo].[AIR_GetLikeEventsbyEventParams] @InputParamTable nvarchar(2555),
    @eventFilterTbl NVARCHAR(2555),
    @guid nvarchar(255) AS Begin
SET NOCOUNT ON;
Declare @IndustryLossTbl nvarchar(255)
Declare @Query nvarchar(max)
Declare @tblAIRCatalog nvarchar(255)
Declare @ModelCode nvarchar(50)
Declare @OutputTable nvarchar(255)
SET @OutputTable = '[EventRepsonse].[dbo].[AIR_Filtered_EventId_' + @guid + ']'
Declare @fieldName nvarchar(255)
Declare @minValue float
Declare @maxValue float
Declare @value nvarchar(2550)
Declare @param_data_type NVARCHAR(255)
Declare @param_id INT
Declare @param_max_id INT
Declare @display_name NVARCHAR(255)
Declare @IndustryLoss NVARCHAR(255)
Declare @minIndustryLoss float
Declare @maxIndustryLoss FLOAT
DECLARE @IndustryArea nvarchar(255)
Declare @IndustrySubArea nvarchar(255)
Declare @IndustryFilter nvarchar(max) = ''
Declare @IndustryGrpFilter nvarchar(max) = ''
Declare @IndustrySelectFilter nvarchar(max) = ''
Declare @paramFilterString NVARCHAR(max) = '' 

-- Input Table 
drop table if EXISTS #inputParams
create table #inputParams(ID int, 
                          ModelCode nvarchar(255), 
                          Param nvarchar(255), 
                          Min Float,
                          Max Float,
                          Value nvarchar(255),
                          param_Type nvarchar(255),
                          param_data_type nvarchar(255))

set @Query = 'insert into #inputparams(ID,
                                        ModelCode,
                                        Param,
                                        Min,
                                        Max,
                                        Value,
                                        Param_type,
                                        Param_data_type)
    select ID, modelCode, param,
           min, max, value,
           param_type, param_data_type
    from ' + @InputParamTable

print @Query

EXEC sp_Executesql @Query

-- select modelcode

select @Modelcode = a.value
from #inputParams a
where a.Param_type = 'Model' and a.Param = 'ModelCode'

print @ModelCode

-- select industry loss table and event catalog table

select @IndustryLossTbl = industryLossTable, @tblAIRCatalog = AIRCatalogTable
from EventCatalog.dbo.TablesbyModelCode

-- Extract parameter table from input table

drop table if exists #tmpEventParams
select ROW_NUMBER() over (order by ID) as ID, a.Param as FieldName, a.Min as minValue, a.Max as maxValue, a.[Value], a.param_data_type
into #tmpEventParams
from #inputParams a
where a.param_Type = 'Param' and a.ModelCode = @ModelCode

-- Set ID and MaxID for Loop

set @param_id = 1
set @param_max_id = (select max(ID) from #tmpEventParams)

-- Loop

BEGIN

    select @FieldName = FieldName,
           @minValue = minValue,
           @maxValue = maxValue,
           @param_data_type = param_data_type,
            @Value = [Value]
    from #tmpEventParams
    where ID = @param_id

    IF @fieldName is not NULL
    BEGIN
         IF @param_data_type = 'List' and @value is not null
         Begin
             set @value = '''' + REPLACE(@value, ',', ''',''') + ''''
             set @paramFilterString = @paramFilterString + ' AND cat.' + 
                QUOTENAME(@fieldName) + ' in (' + cast(@value as nvarchar(2000)) + ')'
         End
         IF @param_data_type = 'Text' and @value is not null
         Begin
             set @value = '''' + REPLACE(@value, ',', ''',''') + ''''
             set @paramFilterString = @paramFilterString + ' AND cat.' + 
                QUOTENAME(@fieldName) + ' Like %' + cast(@value as nvarchar(2000)) + '%'
        END
        
    END 


END



END
GO