



-- declare @ActivityMonitorID int 
-- declare @YELTType nvarchar(255) 
-- declare @sqlQuery nvarchar(2555)


-- set @ActivityMonitorID = 11
-- set @YELTType = 'Gross'    -- Gross or Net


-- IF @YELTType = 'Gross'
-- begin 
-- set @sqlQuery = 
-- ' select * 
-- from Data.dbo.YELT_' + cast(@ActivityMonitorID as nvarchar(25))

-- print @sqlQuery

-- execute sp_executeSQL @sqlQuery
-- end


-- IF @YELTType = 'Net'
-- Begin
-- print 'i'
-- set @sqlQuery = 
-- ' select * 
-- from Analysis.dbo.Analysis_' + cast(@ActivityMonitorID as nvarchar(25))

-- print @sqlQuery

-- execute sp_executeSQL @sqlQuery
-- end



USE ARAData_1;  
GO  
CREATE PROCEDURE GetYELTfromActivityMonitorID  
     @ActivityMonitorID int,
     @YELTType nvarchar(255),
     @Debug int = 0
AS   

    SET NOCOUNT ON;  

    declare @sqlQuery nvarchar(2555)
    
    IF @YELTType = 'Gross'
    begin 
    set @sqlQuery = 
    ' select * 
    from Data.dbo.YELT_' + cast(@ActivityMonitorID as nvarchar(25))

    if @Debug = 1
    begin
    print @sqlQuery
    end

    execute sp_executeSQL @sqlQuery
    end


    IF @YELTType = 'Net'
    Begin
    print 'i'
    set @sqlQuery = 
    ' select * 
    from Analysis.dbo.Analysis_' + cast(@ActivityMonitorID as nvarchar(25))

    
    if @Debug = 1
    begin
    print @sqlQuery
    end

    execute sp_executeSQL @sqlQuery
    end


GO  
