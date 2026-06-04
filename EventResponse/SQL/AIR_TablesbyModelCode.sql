

drop table if exists AIREventInfo..TablesbyModelCode


create table AIREventInfo..TablesbyModelCode(ID INT Identity(1,1), ModelCode nvarchar(255),AIR_IndustryLossTable nvarchar(255), PIMCO_IndustryLossTable nvarchar(255), AIRCataLogTable nvarchar(255))

INSERT INTO AIREventInfo..TablesbyModelCode( ModelCode, AIR_IndustryLossTable, PIMCO_IndustryLossTable, AIRCataLogTable) 
VALUES ('5',	'Industry_Loss_2021_US_WF_wotAdj_bySubarea',	'Industry_Loss_2021_US_WF_wAdj_bySubarea',	'TblModel05'),
('6',	'Industry_Loss_2021_AUS_ANP_wotAdj_BySubarea',	'Industry_Loss_2021_AUS_ANP_wAdj_BySubarea',	'TblModel06'),
('8',	'Industry_Loss_2021_US_FL_wotAdj_bySubarea',	'Industry_Loss_2021_US_FL_wAdj_bySubarea',	'TblModel08'),
('11',	'Industry_Loss_2021_US_EQ_wotAdj_bySubarea',	'Industry_Loss_2021_US_EQ_wAdj_bySubarea',	'TblModel11_TD'),
('13',	'Industry_Loss_2021_US_EQ_wotAdj_bySubarea',	'Industry_Loss_2021_US_EQ_wAdj_bySubarea',	'TblModel13'),
('14',	'Industry_Loss_2021_US_EQ_wotAdj_bySubarea',	'Industry_Loss_2021_US_EQ_wAdj_bySubarea',	'TblModel14'),
('15',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel15'),
('18',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel18'),
('20',	'Industry_Loss_2021_US_SCS_wotAdj_bySubarea',	'Industry_Loss_2021_US_SCS_wAdj_bySubarea',	'TblModel20'),
('23',	'Industry_Loss_2022_US_Wind_SubArea_AIR',	'Industry_Loss_2022_US_Wind_SubArea_PIMCOView',	'TblModel23_Flat'),
('27',	'Industry_Loss_2022_US_Wind_SubArea_AIR',	'Industry_Loss_2022_US_Wind_SubArea_PIMCOView',	'TblModel27_WSST_Flat'),
('28',	'Industry_Loss_2021_US_WT_wotAdj_bySubarea',	'Industry_Loss_2021_US_WT_wAdj_bySubarea',	'TblModel28'),
('22',	'Industry_Loss_2022_US_SCS_SubArea_AIR',	'Industry_Loss_2022_US_SCS_SubArea_PIMCOView',	'TblModel20'),
('30',	'Industry_Loss_2021_Canada_Wind_wotAdj_bySubarea',	'Industry_Loss_2021_Canada_Wind_wAdj_bySubarea',	'TblModel30'),
('31',	'Industry_Loss_2022_EUR_EQ_SubArea_AIR',	'Industry_Loss_2022_EUR_EQ_SubArea_PIMCOView',	'TblModel31'),
('33',	'Industry_Loss_2021_EUR_EQ_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_EQ_wAdj_BySubarea',	'TblModel33'),
('41',	'Industry_Loss_2021_EUR_Wind_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_Wind_wAdj_BySubarea',	'TblModel41'),
('42',	'Industry_Loss_2021_Canada_ANP_wotAdj_bySubarea',	'Industry_Loss_2021_Canada_ANP_wAdj_bySubarea',	'TblModel42'),
('43',	'Industry_Loss_2021_EUR_SCS_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_SCS_wAdj_BySubarea',	'TblModel43'),
('44',	'Industry_Loss_2021_AUS_ANP_wotAdj_BySubarea',	'Industry_Loss_2021_AUS_ANP_wAdj_BySubarea',	'TblModel44'),
('51',	'Industry_Loss_2021_AUS_EQ_wotAdj_BySubarea',	'Industry_Loss_2021_AUS_EQ_wAdj_BySubarea',	'TblModel51'),
('52',	'Industry_Loss_2021_JP_EQ_wotAdj_BySubarea',	'Industry_Loss_2021_JP_EQ_wAdj_BySubarea',	'TblModel52_AIRTD'),
('54',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel54'),
('55',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel55'),
('58',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel58'),
('60',	'Industry_Loss_2021_JP_Wind_wotAdj_BySubarea',	'Industry_Loss_2021_JP_Wind_wAdj_BySubarea',	'TblModel60_Flat'),
('61',	'Industry_Loss_2021_AUS_Wind_wotAdj_BySubarea',	'Industry_Loss_2021_AUS_Wind_wAdj_BySubarea',	'TblModel61_Flat'),
('68',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel68_Flat'),
('70',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel70'),
('72',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel72'),
('76',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'Industry_Loss_2021_World_ANP_wotAdjwDS',	'TblModel76'),
('90',	'Industry_Loss_2021_EUR_FL_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_FL_wAdj_BySubarea',	'TblModel90'),
('92',	'Industry_Loss_2021_EUR_FL_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_FL_wAdj_BySubarea',	'TblModel92'),
('94',	'Industry_Loss_2021_EUR_FL_wotAdj_BySubarea',	'Industry_Loss_2021_EUR_FL_wAdj_BySubarea',	'TblModel94')



select * 
from AIREventInfo..TablesbyModelCode