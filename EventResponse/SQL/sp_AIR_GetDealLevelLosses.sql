

USE EventResponse
Go

SET ANSI_NUlls On
Go
Set QUOTED_IDENTIFIER on
Go


Create PROCEDURE
dbo.AIR_GetDealLevelLosses
@ReportID int,
@likeEventTableName nvarchar(255),
@perspective nvarchar(25),
@portfolioLossTableName nvarchar(255) output,
@dealLevelLossTableName NVARCHAR(255) output,
@guid nvarchar(255),
@debug int = 0

AS
Begin

    print 'Hello stored procedue'
End 



