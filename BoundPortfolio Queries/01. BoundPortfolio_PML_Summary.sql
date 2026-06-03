CREATE TABLE dbo.UI_PerilRegion_PML
(
    ExportID NVARCHAR(25),
    PerilRegion NVARCHAR(100),

    Actual_100yr_Occ_PML_Dollar FLOAT,
    Actual_500yr_Agg_PML_Dollar FLOAT,

    Actual_100yr_Occ_PML_Pct FLOAT,
    Actual_500yr_Agg_PML_Pct FLOAT,

    Guideline_100yr_Occ_PML_Pct FLOAT,
    Guideline_500yr_Agg_PML_Pct FLOAT,

    GuidelineCheck_100yr_Occ_PML NVARCHAR(20),
    GuidelineCheck_500yr_Agg_PML NVARCHAR(20),

    CreatedDate DATETIME DEFAULT GETDATE(),

    CONSTRAINT PK_UI_PerilRegion 
    PRIMARY KEY (ExportID, PerilRegion)
);

EXEC dbo.arapl_getPerilRegion_PML_UI '20862';

SELECT PerilRegion,Actual_100yr_Occ_PML_Dollar,Actual_100yr_Occ_PML_Pct,Actual_500yr_Agg_PML_Dollar,Actual_500yr_Agg_PML_Pct,GuidelineCheck_100yr_Occ_PML,GuidelineCheck_500yr_Agg_PML
FROM dbo.UI_PerilRegion_PML
WHERE ExportID = '20862'
ORDER BY PerilRegion;


