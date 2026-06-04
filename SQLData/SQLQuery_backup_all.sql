DECLARE @DatabaseName NVARCHAR(128);
DECLARE @BackupPath NVARCHAR(256) = 'F:\Manage\SQLData.PrimarySQL2022\'; -- Specify your backup path
DECLARE @ShrinkTargetPercent INT = 10; -- Target free space percentage after shrinking

-- Create the backup directory if it doesn't exist (SQL Server needs permissions)
-- This is a basic example; for production, ensure robust error handling and permissions.
--EXEC master.dbo.xp_cmdshell 'IF NOT EXIST "C:\SQLBackups\" mkdir "C:\SQLBackups\"';

DECLARE db_cursor CURSOR FOR
SELECT name
FROM sys.databases
WHERE state = 0 -- ONLINE databases
AND is_in_standby = 0 -- Not in standby mode
AND is_read_only = 0 -- Not read-only
AND name NOT IN ('master', 'tempdb', 'model', 'msdb'
,'AIRCompanyLoss'
,'AIRGlobalSetting'
,'AIRProject'
,'AIRReinsurance'
,'AIRSecurity'
,'AIRUserSetting'
,'AIRWork'
,'AIRResult'
,'AIRExposure'
,'AIRExposureSummary'
,'AIRExposureRe'
,'AIRResultRe'
,'Z_DBAdmin'
) -- Exclude system databases
ORDER BY Name 

OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @DatabaseName;

WHILE @@FETCH_STATUS = 0
BEGIN
    PRINT 'Processing database: ' + @DatabaseName;

    -- 1. Shrink the database
    -- This shrinks data and log files, aiming for @ShrinkTargetPercent free space.
    -- Consider DBCC SHRINKFILE for more granular control over individual files.
    EXEC ('DBCC SHRINKDATABASE(' + @DatabaseName + ', ' + @ShrinkTargetPercent + ');');
    PRINT 'Database ' + @DatabaseName + ' shrunk.';

	IF @DatabaseName LIKE '%ARAPL%'
	BEGIN
    -- 2. Backup the database
    DECLARE @BackupFileName NVARCHAR(256);
    SET @BackupFileName = @BackupPath + @DatabaseName + '_' + FORMAT(GETDATE(), 'yyyyMMdd_HHmmss') + '.bak';

    EXEC ('BACKUP DATABASE ' + @DatabaseName + ' TO DISK = ''' + @BackupFileName + ''' WITH COMPRESSION, INIT, STATS = 10;');
    PRINT 'Database ' + @DatabaseName + ' backed up to ' + @BackupFileName;
	END

    FETCH NEXT FROM db_cursor INTO @DatabaseName;
END;

CLOSE db_cursor;
DEALLOCATE db_cursor;

PRINT 'All specified databases processed.';