CREATE TABLE dbo.UI_ILS_Monthly_PremiumBreakup
(
    ExportID      NVARCHAR(255),
    MonthNo       INT,
    GrossPremium  DECIMAL(18,2),
    PremiumShare  DECIMAL(18,8)
)

SELECT 
    
    LEFT(DATENAME(MONTH, DATEADD(MONTH, MonthNo - 1, 0)), 3) AS Month,
    GrossPremium
    
FROM dbo.UI_ILS_Monthly_PremiumBreakup
ORDER BY MonthNo;