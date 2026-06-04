
-- select *
-- from sys.databases

USE EventCatalog

-- Industry loss info Table

Create Table EventCatalog.dbo.IndustryLossTableInfo
(ID Int Identity(1,1), 
lossTableName nvarchar(255),
region nvarchar(25),
subRegion nvarchar(25),
peril nvarchar(25),
isAdjusted int,
isSubArea int,
isLOB int,
modelCode int)


-- select * 
-- from EventCatalog.dbo.IndustryLossTableInfo


-- Tables by Model Codes

CREATE TABLE
EventCatalog.dbo.TablesbyModelCode
(ID Int identity(1,1),
modelCode int,
industryLossTable nvarchar(255),
clientIndustryLossTable nvarchar(255),
AIRCatalogTable nvarchar(255))

select * 
from EventCatalog.dbo.TablesbyModelCode 

-- vAIR Event info table

CREATE TABLE
EventCatalog.dbo.vAIREventInfo_default
(ID int identity(1,1),
catalogTypeCode nvarchar(25),
modelCode int,
catalogTagTypeCode int,
eventID int,
day int,
year int,
eventNumber int,
eventDescription nvarchar(max))

select * 
from EventCatalog.dbo.vAIREventInfo_default

-- Industry Loss Table

CREATE TABLE
EventCatalog.DBo.Industry_Loss_Year_Region_Peril_SubRegion_Vendor_YELT
(ID int identity(1,1),
eventID int,
yearID int,
countryName nvarchar(255),
areaName nvarchar(255),
subAreaName nvarchar(255),
industryLoss float)

select * 
from EventCatalog.DBo.Industry_Loss_Year_Region_Peril_SubRegion_Vendor_YELT

-- Event Catalog Tables

-- AIR Model Code = 27 = US HU

CREATE table
EventCatalog.dbo.tblModel27_WSST
(ID int identity(1,1),
eventID int,
modelID int,
event int,
year int,
day int,
eventDescription nvarchar(2550),
landfall1 nvarchar(2),
CHI1 float,
SaffSimpson1 int,
SegmentID1 int,
CenPress1 float, 
MaxWind1 float,
RadMax1 float,
Speed1 float,
Angle1 float,
Longitude1 float,
Latitude1 float,
Area1 float,
SubArea1 float,
landfall2 nvarchar(2),
CHI2 float,
SaffSimpson2 int,
SegmentID2 int,
CenPress2 float, 
MaxWind2 float,
RadMax2 float,
Speed2 float,
Angle2 float,
Longitude2 float,
Latitude2 float,
Area2 float,
SubArea2 float,
landfall3 nvarchar(2),
CHI3 float,
SaffSimpson3 int,
SegmentID3 int,
CenPress3 float, 
MaxWind3 float,
RadMax3 float,
Speed3 float,
Angle3 float,
Longitude3 float,
Latitude3 float,
Area3 float,
SubArea3 float,
landfall4 nvarchar(2),
CHI4 float,
SaffSimpson4 int,
SegmentID4 int,
CenPress4 float, 
MaxWind4 float,
RadMax4 float,
Speed4 float,
Angle4 float,
Longitude4 float,
Latitude4 float,
Area4 float,
SubArea4 float,
landfall5 nvarchar(2),
CHI5 float,
SaffSimpson5 int,
SegmentID5 int,
CenPress5 float, 
MaxWind5 float,
RadMax5 float,
Speed5 float,
Angle5 float,
Longitude5 float,
Latitude5 float,
Area5 float,
SubArea5 float,
landfall6 nvarchar(2),
CHI6 float,
SaffSimpson6 int,
SegmentID6 int,
CenPress6 float, 
MaxWind6 float,
RadMax6 float,
Speed6 float,
Angle6 float,
Longitude6 float,
Latitude6 float,
Area6 float,
SubArea6 float,
landfall7 nvarchar(2),
CHI7 float,
SaffSimpson7 int,
SegmentID7 int,
CenPress7 float, 
MaxWind7 float,
RadMax7 float,
Speed7 float,
Angle7 float,
Longitude7 float,
Latitude7 float,
Area7 float,
SubArea7 float,
landfall8 nvarchar(2),
CHI8 float,
SaffSimpson8 int,
SegmentID8 int,
CenPress8 float, 
MaxWind8 float,
RadMax8 float,
Speed8 float,
Angle8 float,
Longitude8 float,
Latitude8 float,
Area8 float,
SubArea8 float,
)


select  * 
from EventCatalog.dbo.tblModel27_WSST

-- AIR Model Code = 11 = US EQ

CREATE TABLE
EventCatalog.dbo.tblModel11_TD
(ID int identity(1,1),
eventID int,
modelID int,
event int,
year int,
day int,
eventDescription nvarchar(2550),
EQ_Depth float,
EQ_Source_Type nvarchar(255),
EQ_Fault_Name nvarchar(255),
Fault_Type nvarchar(2),
Rupture_Length float,
Rupture_Width float,
Dip int,
Dip_Azimuth int,
Fault_Segments int,
Epicenter nvarchar(25),
Magnitude float,
Longitude float,
latitude float,
areaName nvarchar(255),
subAreaName nvarchar(255),
tsunamiEvent nvarchar(2),
inducedEvent nvarchar(2))


select * 
from tblModel11_TD
-- AIR Model Code = 08 = US FL

CREATE Table
EVentCatalog.dbo.tblModel08
(ID int identity(1,1),
eventID int,
modelID int,
event int,
year int,
day int,
eventDescription nvarchar(2550),
HourofDay int,
duration_Hours int,
timeToFirstRunoff_HourofYear int,
timeToLastRunoff_HourofYear int,
timeToFirstPeakFlow_HourofYear int,
timeToLastPeakFlow_HourofYear int,
numberCatchmentsFlooded int,
numberLinksFlooded int,
AreaCatchmentsFlooded_KmSq float,
LengthLinksFlooded_Km float,
Min_Latitude_Boundary float,
Max_Latitude_Boundary float,
Min_Longitude_Boundary float,
Max_Longitude_Boundary float
)

select * 
from tblModel08
-- AIR Model Code = 05 = US WF

Create Table 
eventCatalog.dbo.tblMode05
(ID int identity(1,1),
eventID int,
modelID int,
event int,
year int,
day int, 
eventDescription nvarchar(2550),
BurnedAreaSqMiles float,
Longitude float,
Latitude float,
AreaName nvarchar(25),
subAreaName nvarchar(255),
FiresInCluster int)

select *
from tblMode05

