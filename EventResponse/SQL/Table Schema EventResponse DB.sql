use EventResponse

-- AIR_InputParamTable
drop table if exists EventResponse.dbo.AIR_InputParamTable
Create Table 
EventResponse.dbo.AIR_InputParamTable
(ID int identity(1,1),
modelCode int,
param nvarchar(255),
min float,
max float,
value nvarchar(max),
param_type nvarchar(255),
param_data_type nvarchar(255))

select * 
from AIR_InputParamTable


-- Live Event Analysis Options

Create Table
EventResponse.dbo.LiveEventAnalysisOption
(ID int identity(1,1),
modelCode int,
param nvarchar(255),
min float,
max float,
value nvarchar(max),
param_type nvarchar(255),
param_data_type nvarchar(255),
displayName nvarchar(255))


select * 
from LiveEventAnalysisOption