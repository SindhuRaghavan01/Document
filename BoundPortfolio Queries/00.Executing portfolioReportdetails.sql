DECLARE @ExportID NVARCHAR(25)
 
exec [arapl_getPortfolioReportDetails]        1,'05/12/2026','AIR Worldwide','all','All','05/12/2026',0,0,@ExportID,1
 
PRINT @ExportID

--558028 
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioDetails_148132
 
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortDetailsDealYELT_148132
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_148132

   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortYELT_148132
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortDetailsYELT_148132
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_byHistoricalEvents_148132
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_AALByMonth148132
    
   -- SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_byRegionPerilPortfolio_148132
   --select * from ARAPL_Works.dbo.tmpBoundPortfolioDetails_148132
   --select * from ARAPL_Analysis.dbo.BoundPortDetails_1
   --SELECT * FROM ARAPL_Analysis.dbo.BoundPort_1
   --SELECT * FROM ARAPL_Analysis.dbo.BoundPortfolioSummary_1

     
     SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_byRegionPerilDeal_558028
     SELECT * FROM ARAPL_Works.dbo.tmpBoundPortfolioSummary_byRegionPerilILSType_558028



   --1. arapl_getBoundPortfolio_RegionPeril
   --exec [arapl_getBoundPortfolio_RegionPeril_UI] 148132
   --select * from ARAPL_Works.dbo.BoundPortfolio_RegionPeril_UI_148132


    --2. arapl_getBoundPortfolio_PnLReport
    --EXEC dbo.[arapl_getBoundPortfolio_PnLReport_UI]148132

    --SELECT ColName,Expected,Yr1_11,Yr2,Yr5,Yr10,Yr25,Yr50,Yr100,Yr250,Yr500,Yr1000
    --FROM dbo.UI_PnL_Results
    --WHERE ExportID =148132
    --ORDER BY ColName;

    
    --3.getBoundPortfolio_PortfolioMetrics
    --EXEC [dbo].[arapl_getBoundPortfolio_PortfolioMetrics_UI] 148132 ;

    --SELECT Metric,USD,[%NAV] FROM dbo.UI_PortfolioMetics
    --WHERE ExportID = '148132';

    --4. getBoundPortfolio_MonthlyEarningPattern
     --EXEC arapl_getBoundPortfolio_MonthlyEarningPattern_UI 148132
    
     SELECT  ILSName,Jan,Feb,Mar,Apr,May,June,July,Aug,Sept,Oct,Nov,Dec
     FROM dbo.UI_ILS_Monthly_AAL

     --5. exec [arapl_getBoundPortfolio_MonthlyPremiumBreakup_UI]
     --EXEC arapl_getBoundPortfolio_MonthlyPremiumBreakup_UI 148132
     
     --SELECT LEFT(DATENAME(MONTH, DATEADD(MONTH, MonthNo - 1, 0)), 3) AS Month,
     --GrossPremium FROM dbo.UI_ILS_Monthly_PremiumBreakup
     --ORDER BY MonthNo;


       -- 6.getBoundPortfolio_RegionPeril
       --Exec dbo.[arapl_getBoundPortfolio_RegionPeril_UI] 148132

        --select ID,RegionPerilDescription,AAL,OtherAAL,VAR_100,TVAR_100,CoTVAR_100,AEP_100_OtherAAL,PnL 
        --from ARAPL_Works.dbo.BoundPortfolio_RegionPeril_UI_148132
        --order by id asc


       --7. EP CUrve

       --EXEC arapl_getBoundPortfolio_EPCurve_UI 148132
       --SELECT * FROM UI_EPCurve


       --8. LimitBySecurityInstrument

       --EXEC arapl_getBoundPortfolio_LimitBySecurityInstrument_UI   148132

        --SELECT SecurityInstrument,ExposedLimit,ActualPctOfAUM,GuidelineCheck,EL_Pct,AUM_Pct,Premium_Pct,Premium_Dollar,CreatedDate
        --FROM dbo.UI_BoundPortfolio_LimitBySecurityInstrument
        --WHERE ExportID = '148132'
        --ORDER BY ActualPctOfAUM DESC

        --9.[arapl_getBoundPortfolio_LimitbyIssuer_UI]
        --EXEC [dbo].[arapl_getBoundPortfolio_LimitbyIssuer_UI] 148132
        
        --SELECT Issuer,[Exposed Limit] ,[Actual % of AUM],Guideline,[Guideline Check] FROM IssuerGuidelineOutput

     

