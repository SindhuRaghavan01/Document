

USE EventResponse
Go

SET ANSI_NUlls On
Go
Set QUOTED_IDENTIFIER on
Go


Create PROCEDURE
dbo.AIR_GetLikeEventTracks
@InputShapeTableName nvarchar(2555),
@EventIDFilterTable nvarchar(2555),
@guid nvarchar(2555),
@debug nvarchar(255)
AS
Begin

    print 'Hello stored procedue'
End 
