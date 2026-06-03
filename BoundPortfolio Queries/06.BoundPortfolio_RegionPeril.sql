CREATE TABLE ARAPL_Works.dbo.BoundPortfolio_RegionPeril_UI (
    ID INT,
    RegionPerilDescription NVARCHAR(255),
    AAL FLOAT,
    OtherAAL FLOAT,
    VAR_100 FLOAT,
    TVAR_100 FLOAT,
    CoTVAR_100 FLOAT,
    AEP_100_OtherAAL FLOAT,
    PnL FLOAT,
    Allianz FLOAT,
    [CAT Bond] FLOAT,
    [ILW Long] FLOAT,
    [ILW Short] FLOAT,
    Reinsurance FLOAT,
    [Retro QS] FLOAT,
    [Retro XOL] FLOAT,
    [Side Car] FLOAT
);


select ID,RegionPerilDescription,AAL,OtherAAL,VAR_100,TVAR_100,CoTVAR_100,AEP_100_OtherAAL,PnL from ARAPL_Works.dbo.BoundPortfolio_RegionPeril_UI_20862
order by id asc
