"""optimization_queries contains all raw queries used in optimization
"""
optimization_queries = {
    "optimization_input_query": """
select t1.ILSInstrumentID, t1.Updated_LayerID, t1.description, t1.CipherID, t1.OptimizationLayerID, t1.AUM, t1.Premium, t1.Initial_Fees, t1.Min_Share, t1.Max_Share, t1.Qscore, t1.layer_group, t1.Enable_Exclude, t1.Rein_Retro, t1.Occ_Agg_QS, t1.Status, t1.OptimizationID, t1.Participation, t2.Sponsor, t1.CrownActivityID, t1.AnalyzeReActivityID,
isnull(t2.Reinstatement, 0) Reinstatement, isnull(t2.ManagementFees, 0) ManagementFees, isnull(t2.ProfitCommission, 0) ProfitCommission, isnull(t2.Hurdle, 0) Hurdle, isnull(t2.UWOverride, 0) UWOverride, isnull(t2.Origination, 0) Origination, isnull(t2.FET, 0) FET, isnull(t2.Setup, 0) Setup, isnull(t2.Brokerage, 0) Brokerage,
case when left(t1.layer_group, 7) = 'g600xxx' then 0
        when (isnull(t2.Reinstatement, 0)+ isnull(t2.ManagementFees, 0)+ isnull(t2.ProfitCommission, 0)+ isnull(t2.Hurdle, 0)+ isnull(t2.UWOverride, 0)+ isnull(t2.Origination, 0)+ isnull(t2.FET, 0)+ isnull(t2.Setup, 0)) = 0 then 0
        else 1 end as calcExpenses
from(
select b.ILSInstrumentID, b.Updated_LayerID, b.description, b.CipherID, b.OptimizationLayerID, b.AUM, b.Premium, b.Initial_Fees, b.Min_Share, b.Max_Share, b.Qscore, b.layer_group, b.Enable_Exclude, b.Rein_Retro, b.Occ_Agg_QS, b.Status, b.OptimizationID, b.Participation,
max(case when a.AnalysisType=12 then a.ActivityID else -100 end) CrownActivityID, max(case when a.AnalysisType=24 then a.ActivityID else -100 end) AnalyzeReActivityID
from ARAPL_Configuration.dbo.ILSActivityMonitor a
inner join ARAPL_Optimization.dbo.OptimizationInput b
on a.AcctGrpId = b.ILSInstrumentID and b.OptimizationID = {OptimizationID}
and a.Vendor = 'PIMCOView'
and a.IsDefaultAnalysis = 1
where b.Min_Share <> b.Max_Share or b.Min_share <> 0 or b.Min_share <> 0
group by b.ILSInstrumentID, b.Updated_LayerID, b.description, b.CipherID, b.OptimizationLayerID, b.AUM, b.Premium, b.Initial_Fees, b.Min_Share, b.Max_Share, b.Qscore, b.layer_group, b.Enable_Exclude, b.Rein_Retro, b.Occ_Agg_QS, b.Status, b.OptimizationID, b.Participation
)t1
inner join ARAPL_Configuration.dbo.AnalyzeReMetadata t2
on t1.AnalyzeReActivityID = t2.analyzereActivityID
""",
    "optimization_parameters_query": """SELECT [ParameterID],[Discretization],[Cardinality],[CorrelateGroups],[Solver],[MutationRate],[SeedingOverride],[PopulationSize],[Iterations],[OptimizationID] FROM [ARAPL_Optimization].[dbo].[OptimizationParameters] WHERE OptimizationID={OptimizationID}""",
    "optimization_objectives_query": """SELECT [ObjConsID],[OptimizationID],[Criterion],[Functions],[Names],[Types],[Min],[Max],[Return_Period],[Loss_Perspective],[Threshold],[IsInclusive],[Fields],[ProportionateTo],[Defaults],[WeightedAvg],[ReturnPeriodUpper],[ReturnPeriodLower],[Layer_Groups],[PrimaryPerspective],[PrimaryLayerGroups],[ComponentPerspective],[ComponentLayerGroups],[LossFilter]
  FROM [ARAPL_Optimization].[dbo].[Optimization_Objectives_Constraints] a
  where a.Criterion = 'Objective' and OptimizationID={OptimizationID}""",
    "optimization_constraints_query": """SELECT [ObjConsID],[OptimizationID],[Criterion],[Functions],[Names],[Types],[Min],[Max],[Return_Period],[Loss_Perspective],[Threshold],[IsInclusive],[Fields],[ProportionateTo],[Defaults],[WeightedAvg],[ReturnPeriodUpper],[ReturnPeriodLower],[Layer_Groups],[PrimaryPerspective],[PrimaryLayerGroups],[ComponentPerspective],[ComponentLayerGroups],[LossFilter]
  FROM [ARAPL_Optimization].[dbo].[Optimization_Objectives_Constraints] a
  where a.Criterion like 'Constraint%' and OptimizationID={OptimizationID}""",
    "fitness_query": """exec arapl_optimization.dbo.sp_FitnessEvaluation {optimization_id}, {iteration}, {save_to_db}""",
    "create_fitness_table_query": """if object_id('{fitness_table_name}', 'U') is not null drop table {fitness_table_name};  create table {fitness_table_name}(solution_number int, layer_number int, ILSInstrumentID INT, share float)""",
    "loss_filter_peril_query": """select distinct LossFilter, isnull(Peril, '') Peril, isnull(Region, '') Region, isnull(SubRegion, '') SubRegion, isnull(Country, '') Country
from ARAPL_Configuration.dbo.PerilRegion_Bound a
where a.LossFilter in ({loss_filter_quoted_list})""",
    "get_ylt_query": """exec ARAPL_Optimization.dbo.arapl_getYELTtoYLT 'arapl_analysis.dbo.analysis_{crown_activity_id}', 'AEP',@region='{region}', @Peril='{peril}', @Subregion='{subregion}', @Country='{country}'
    """
}
