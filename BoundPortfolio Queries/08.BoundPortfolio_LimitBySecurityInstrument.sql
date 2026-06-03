CREATE TABLE dbo.UI_BoundPortfolio_LimitBySecurityInstrument
(
    ExportID NVARCHAR(25),
    SecurityInstrument NVARCHAR(100),
    ExposedLimit FLOAT,
    ActualPctOfAUM FLOAT,
    Guideline FLOAT,
    GuidelineCheck NVARCHAR(20),
    EL_Pct FLOAT,
    AUM_Pct FLOAT,
    Premium_Pct FLOAT,
    Premium_Dollar FLOAT,
    CreatedDate DATETIME DEFAULT GETDATE(),

    CONSTRAINT PK_UI_BoundPortfolio 
    PRIMARY KEY (ExportID, SecurityInstrument)
);


EXEC arapl_getBoundPortfolio_LimitBySecurityInstrument_UI   20862

SELECT SecurityInstrument,ExposedLimit,ActualPctOfAUM,GuidelineCheck,EL_Pct,AUM_Pct,Premium_Pct,Premium_Dollar,CreatedDate
FROM dbo.UI_BoundPortfolio_LimitBySecurityInstrument
WHERE ExportID = '20862'
ORDER BY ActualPctOfAUM DESC
