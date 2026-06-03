CREATE TABLE  dbo.UI_EPCurve_Results
(
    ExportID INT,
    AggLoss FLOAT,
    [Rank] INT,
    [Probability] FLOAT,
    [UW P&L] FLOAT,
    [Total P&L] FLOAT,
    [Excess Return(% Total)] FLOAT,
    [Total Return(% Total)] FLOAT,
    [Excess Return(% Deployed)] FLOAT,
    [Total Return(% Deployed)] FLOAT
)


EXEC arapl_getBoundPortfolio_EPCurve_UI              20862

SELECT * FROM UI_EPCurve