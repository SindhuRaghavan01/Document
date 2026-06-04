
/*

Input Parameter Table

Based on param_data_type column value
- Range = insert values in Min and Max column
- not Range = insert values in Value column

For List param_data_type isert value in single quotes with comma seperated
e.g., 'Florida,Texas'

If default param value is not changed, insert values as suggested in comments

*/




drop table if exists [ARAPL_LiveEventData].[dbo].[AIR_inputParameterTable]
GO

create table [ARAPL_LiveEventData].[dbo].[AIR_inputParameterTable](
	ID INT Identity(1,1), 
	ModelCode nvarchar(255),
	Param nvarchar(255),
	Min float, 
	Max float, 
	Value nvarchar(max),
	Param_type nvarchar(255), 
	Param_data_type nvarchar(255))
GO
Insert Into [ARAPL_LiveEventData].[dbo].[AIR_inputParameterTable]
( ModelCode, Param, Min, Max, Value, Param_type, Param_data_type) Values 
('',	'ModelCode',	NULL,	NULL,	'27',	'Model',	'nvarchar'),
('',	'IndustryLoss',	0,	10e9,	NULL,	'IndustryLoss',	'Range'), -- default IndustryLoss Range 0-10e10
('',	'AreaName',	NULL,	NULL,	null,	'IndustryLoss',	'List'),
('',	'SubAreaName',	NULL,	NULL,	NULL,	'IndustryLoss',	'List'),
('27',	'EventID',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'ModelID',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'Event',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'Year',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'Day',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'EventDescription',	NULL,	NULL,	NULL,	'Param',	'Text'),
('27',	'Landfall',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'CHI',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'SaffSimpson',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'SegmentID',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'CenPress',	NULL,	NULL,	NULL,	'Param',	'Range'),
('27',	'MaxWind',	0,	800,	NULL,	'Param',	'Range'), -- default range 0-800 for MaxWind
('27',	'RadMax',	0,	200,	NULL,	'Param',	'Range'), -- default range 0-200 for RadMax
('27',	'Speed',	40,	500,	NULL,	'Param',	'Range'), -- default range 0-500 for speed
('27',	'Angle',	-180,	180,	NULL,	'Param',	'Range'), -- default range -180 to 180 for angle
('27',	'Longitude',	-180,	180,	NULL,	'Param',	'Range'), -- default longitude -180 to 180 
('27',	'Latitude',	-90,	90,	NULL,	'Param',	'Range'), -- default latitude -90 to 90
('27',	'Area',	null, null,'Alabama,Mississippi,Florida,Tennesse',	'Param',	'List'),
('27',	'SubArea',	NULL,	NULL,	NULL,	'Param',	'List'),
('27',	'LandfallNumber',	NULL,	NULL,	NULL,	'Param',	'List')




--Nagasaki,Oita,Saga,Kumamoto,Fukuoka,Kagoshima,Miyazaki

select * 
from [ARAPL_LiveEventData].[dbo].[AIR_inputParameterTable]

--select *--min(radmax),max(radmax),min(speed), max(speed),min(cenpress), max(cenpress) 
--from AIREventInfo.dbo.TblModel27_WSST_Flat a
--where a.SubArea = 'Lee'

--select *
--from AIREventInfo.dbo.Industry_Loss_2021_US_Wind_wAdj_BySubarea a
--where a.EventID = 270005223


