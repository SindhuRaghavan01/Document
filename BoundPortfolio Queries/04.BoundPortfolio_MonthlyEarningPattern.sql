CREATE TABLE dbo.UI_ILS_Monthly_AAL
(
    ExportID NVARCHAR(255),
    ILSName  NVARCHAR(255),
    Jan  DECIMAL(18,6),
    Feb  DECIMAL(18,6),
    Mar  DECIMAL(18,6),
    Apr  DECIMAL(18,6),
    May  DECIMAL(18,6),
    June DECIMAL(18,6),
    July DECIMAL(18,6),
    Aug  DECIMAL(18,6),
    Sept DECIMAL(18,6),
    Oct  DECIMAL(18,6),
    Nov  DECIMAL(18,6),
    Dec  DECIMAL(18,6)
)

SELECT  

ILSName,Jan,Feb,Mar,Apr,May,June,July,Aug,Sept,Oct,Nov,Dec

FROM dbo.UI_ILS_Monthly_AAL

EXEC arapl_getBoundPortfolio_MonthlyEarningPattern_UI 20862