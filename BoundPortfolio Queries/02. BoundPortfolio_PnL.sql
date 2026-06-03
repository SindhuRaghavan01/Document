CREATE TABLE dbo.UI_PnL_Results
(
    ExportID NVARCHAR(25),
    ColName NVARCHAR(100),

    Expected FLOAT,
    Yr1_11 FLOAT,
    Yr2 FLOAT,
    Yr5 FLOAT,
    Yr10 FLOAT,
    Yr25 FLOAT,
    Yr50 FLOAT,
    Yr100 FLOAT,
    Yr250 FLOAT,
    Yr500 FLOAT,
    Yr1000 FLOAT,

    CreatedDate DATETIME DEFAULT GETDATE(),

    CONSTRAINT PK_UI_PnL 
    PRIMARY KEY (ExportID, ColName)
);

EXEC dbo.YourStoredProcedure '20862';

SELECT ColName,Expected,Yr1_11,Yr2,Yr5,Yr10,Yr25,Yr50,Yr100,Yr250,Yr500,Yr1000
FROM dbo.UI_PnL_Results
WHERE ExportID = '20862'
ORDER BY ColName;
