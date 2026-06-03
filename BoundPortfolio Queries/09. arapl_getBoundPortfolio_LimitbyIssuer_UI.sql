CREATE TABLE IssuerGuidelineOutput
(
    Issuer VARCHAR(255),
    [Exposed Limit] FLOAT,
    [Actual % of AUM] FLOAT,
    Guideline FLOAT,
    [Guideline Check] VARCHAR(50)
)

SELECT Issuer,[Exposed Limit] ,[Actual % of AUM],Guideline,[Guideline Check]   FROM IssuerGuidelineOutput