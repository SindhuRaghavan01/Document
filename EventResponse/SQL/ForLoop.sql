
declare @ID int
declare @MaxID int
declare @ActivityMonitorID int
declare @Name NVARCHAR(255)
declare @sponsorName nvarchar(255)
declare @sqlQuery NVARCHAR(255)

drop table if exists #tmp1

-- create table #tmp1(ID int, ActivityMonitorID int)
-- insert into #tmp1
-- select ROW_NUMBER() over (order by ActivityMonitorID), ActivityMonitorID
-- from ARAData_1.dbo.ActivityMonitor
-- where ActivityMonitorID not in (1,2,38,39,40,41,58) and TermsID is null


use Analysis
select ROW_NUMBER() over (order by TABLE_NAME desc) as ID, TABLE_NAME
into #tmp1
from INFORMATION_SCHEMA.TABLES
where TABLE_NAME like 'Analysis_%'

select *
from #tmp1

set @ID = 1

set @MaxID = (select max(ID) from #tmp1)

select @MaxID

while @ID <= @MAxID
Begin

print @ID

select @Name = Table_Name
from #tmp1
where ID = @ID



-- select distinct @sponsorName = Sponsor
-- from ARAData_1.dbo.Instrument
-- Inner join ARAData_1.dbo.Sponsor
-- on Sponsor.SponsorLookUpID = Instrument.SponsorLookUpID
-- where Instrument.InstrumentID = @ActivityMonitorID



set @sqlQuery = ' execute sp_rename  ''Analysis.dbo.'+@Name+''', '''+@Name+'_backup_06202023'''
--  + '
-- select *
-- into  DAta.dbo.YELT_'+ cast(@ActivityMonitorID as nvarchar(25)) +
-- + ' from [Claimsmanager].dbo.YELT_'+ cast(@ActivityMonitorID as nvarchar(25))
-- -- '
-- update ARAData_1.dbo.Instrument
-- set IssuerName = '''+ cast(@sponsorName as nvarchar(255))
-- + ''' where InstrumentID = ' + cast(@ActivityMonitorID as nvarchar(25))

print @sqlQuery

-- Execute sp_ExecuteSQL @sqlQuery

SEt @ID = @ID + 1

end


print 'Completed'


