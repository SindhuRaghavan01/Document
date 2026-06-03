CREATE TABLE dbo.UI_PortfolioMetrics
(
    ExportID NVARCHAR(25),
    Metric NVARCHAR(255),
    USD FLOAT,
    [%NAV] FLOAT,
    CreatedDate DATETIME DEFAULT GETDATE(),

    CONSTRAINT PK_UI_Results 
    PRIMARY KEY (ExportID, Metric)
);


EXEC [dbo].[arapl_getBoundPortfolio_PortfolioMetrics_UI] 20862 ;

SELECT Metric,USD,[%NAV]
FROM dbo.UI_PortfolioMetrics
WHERE ExportID = '20862';
