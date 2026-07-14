# Cat Bond Pricing (Buy/Sell) — SQL Test Automation: Setup Guide

## Files
- `CatBond_Automation_Framework.xlsx` — ReadMe / Config / ParamMap / Lookup_* / TestCases / Results
- `modConnection.bas` — opens the ADO connection from Config
- `modUtils.bas` — parameter binding, lookups, dynamic Results columns
- `modTestRunner.bas` — the engine: `RunAllTests`, `RunSelectedTest`

## What it does
For each enabled TestCases row, VBA calls `sp_RiskAnalyzer_Catbond_Pricing_Buy` (or `_Sell`,
based on the `Buy or Sale` column) with that row's inputs, translated to your real SQL
parameters, and pastes **everything SQL returns** — output parameters, return value, and/or a
result set — into new `Out_<name>` columns on the Results sheet automatically. You don't need
to know the output schema in advance; whatever the stored procedure hands back becomes a column.

## Why ParamMap exists
Your TestCases sheet uses business labels (`Transaction Price`, `Activity ID / ILS Name`).
Your stored procedure wants specific parameter names and, in some cases, translated values:

| TestCases column | -> | SQL parameter | Translation |
|---|---|---|---|
| Transaction Date | -> | `@BuyDate` (Buy rows) / `@SellDate` (Sell rows) | the other is sent NULL |
| Transaction Price | -> | `@BuyPrice` | — |
| Transaction Yield | -> | `@DesiredYield` | — |
| Accrued Interest | -> | `@isAccruedInterest` | Yes/No -> 1/0, via `Lookup_YesNo` |
| Activity ID / ILS Name | -> | `@ILSInstrumentID` | bond name -> ID, via `Lookup_ILSInstrument` |
| Calculation Type | -> | `@CalculationType` | "Yield to Price" -> `'YieldGivenPrice'`, via `Lookup_CalculationType` |
| Portfolio Name | -> | `@Fund` | portfolio name -> Fund ID, via `Lookup_PortfolioFund` |
| Modelling Agent (auto) | -> | `@Vendor` | falls back to `Config!DefaultVendor` if blank |
| (auto-generated) | -> | `@Guid` | unique value per call |
| `Config!DebugFlag` | -> | `@DEBUG` | fixed, from Config |

This is all declared on the **ParamMap** sheet (`SQLParam | SourceType | Source | LookupSheet |
ApplyWhenColumn | ApplyWhenValue | Notes`) — nothing is hardcoded in VBA, so you can retarget or
extend it without touching code.

## One-time setup
1. Save the workbook as **.xlsm**.
2. Alt+F11 → **File → Import File** → import all three `.bas` files. No reference needs
   enabling (late-bound ADO via `CreateObject`).
3. `Config!ConnectionString` → point at your SQL Server.
4. Fill in `Lookup_ILSInstrument`, `Lookup_PortfolioFund`, `Lookup_CalculationType` with your
   real reference data (one sample row each, matching your worked example, is pre-filled).
5. Add/replace rows on `TestCases`.
6. Alt+F8 → `RunAllTests`.

## TestCases columns
`Test Case ID | Enabled | Activity ID / ILS Name | Sponsor | View Selected | Portfolio Name |
Transaction Date | Buy or Sale | Calculation Type | Transaction Price | Transaction Yield |
Sell Quantity | Accrued Interest | Transactional Spread (%) | Company Share (%) |
Modelling Agent (auto) | Coupon Frequency (auto) | Risk Free Rate (auto) |
Stored Procedure (optional) | Expected Result`

- `Enabled = Y` to include the row in `RunAllTests`.
- `Stored Procedure (optional)` — leave blank to auto-derive from `Config!StoredProcPattern`
  (`sp_RiskAnalyzer_Catbond_Pricing_{BuyOrSale}`) using the row's `Buy or Sale` value.
- `Expected Result` — `PASS` marks the row as passing when the SP executes without a SQL error;
  `FAIL` flips that (for negative tests that should error); leave blank to just log output with
  no verdict.

## Results sheet
Fixed columns: `TestID | StoredProcedure | RunTimestamp | DurationMs | Status | ErrorMessage |
ExpectedResult`, followed by dynamic `Out_<name>` columns — one per output parameter, the return
value, and every column of any result set the SP returns. `RunAllTests` clears these dynamic
columns at the start of each run so old columns from a previous SP version don't linger.

## Extending later
- **New SP**: add its parameters as new ParamMap rows, point `Stored Procedure (optional)` at
  it (or extend `StoredProcPattern`), add any new `Lookup_*` sheet it needs. No VBA changes.
- **Numeric pass/fail against an expected price/yield** (rather than just logging output): add
  an `ExpectedValue`/`Tolerance` column pair and a comparison step — happy to add this once you
  see what fields the SP actually returns from a first live run.
