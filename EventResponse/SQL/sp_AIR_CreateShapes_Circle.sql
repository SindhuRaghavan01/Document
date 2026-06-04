USE EventResponse
Go

SET ANSI_NUlls On
Go
Set QUOTED_IDENTIFIER on
Go


Create PROCEDURE
dbo.AIR_CreateShapes_Circle
@Latitudes nvarchar(2555),
@Longitudes nvarchar(2555),
@Radii nvarchar(2555),
@guid nvarchar(255)
AS
Begin

    print 'Hello stored procedue'
End 
