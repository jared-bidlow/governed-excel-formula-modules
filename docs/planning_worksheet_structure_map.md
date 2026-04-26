# Planning Worksheet Structure Map (Reference Parse)

This is a public-safe structure reference derived from a local workbook parse and normalized to the public starter workbook contract.

No workbook binary, private workbook name, source path, person, vendor, or real project data is stored in this repository.

## Reference Parse Shape

| Item | Observed shape |
|---|---|
| Worksheet name | Omitted for public safety |
| Used range | `A1:BM98` |
| Header row | `2` |
| Main data rows | `3:98` |
| Header filter range | `A2:BI97` |
| Explicit list validation | `O3:O98` with allowed values `Y,N` |

The reference parse showed one explicit `Y,N` validation range. In the public starter, that old position is not treated as a stable column contract. The add-in finds the current `Chargeable` column by header name and applies the `Y,N` validation through row `2000`.

## Public Starter Column Bands

| Columns | Role |
|---|---|
| `A:L` | Project identity, source/job keys, status, ownership, and classification fields |
| `M:N` | Chargeability and carry-over inputs |
| `O:AZ` | Annual finance fields and twelve monthly projected/actual/budget triplets |
| `BA:BL` | Comments, planning stage/maturity, internal readiness inputs, resource fields, future tier override, and cancellation flag |

Finance triplets repeat by month:

- `<Month> Projected`
- `<Month> Actuals`
- `<Month>`

## Planning Table Column Map

The public starter `Planning Table` is a 64-column table from `A:BL`. The add-in validates header order from row `2`; formulas should resolve fields by header name unless they are deliberately reading the contiguous finance block from `Annual Projected` through `December`.

| Column | Header | Band | Validation / format | Primary dependencies |
|---|---|---|---|---|
| `A` | `Composite Cat` | Identity / rollup | Manual pre-formula helper for Excel sort, dedupe, and Data > Subtotal workflows | `get.GetIsRollupMask`, `get.GetHasDetailIdentityMask`, `kind.CapConsumeMask`, `CapitalPlanning.BuildGroupedRows`, `Ready.InternalJobs_Export` rollup exclusion. |
| `B` | `Project Description` | Identity / rollup | Required input | `get.GetHasDetailIdentityMask`, `kind.CapConsumeMask`, `Search.Projects_Health`, `Analysis.PM_SPEND_REPORT`, `Analysis.WORKING_BUDGET_SCREEN`, burndown detail, `Ready.InternalJobs_Export`. |
| `C` | `Category` | Identity / grouping | Required input | `get.GetHasDetailIdentityMask`, `kind.CapConsumeMask`, main report detail rows, PM spend grouping fallback, working-budget category grouping, reforecast/detail outputs. |
| `D` | `Source ID` | Source key | Optional identifier | `get.GetHasDetailIdentityMask`, `kind.CapConsumeMask`, `Ready.InternalJobs_Export`. |
| `E` | `Job ID` | Source key | Required for chargeable rows by health check | `get.GetHasDetailIdentityMask`, `kind.CapConsumeMask`, `Search.Projects_Health`, `Ready.InternalJobs_Export`. |
| `F` | `Status` | Workflow | Dropdown from `Validation Lists[Status]`; required header fill | `kind.CapConsumeMask`, `kind.HideClosedMask`, `kind.CloseMetricPack`, main report, PM spend, working budget, scorecard, burndown, defer selectors. |
| `G` | `BU` | Organization / cap | Required header fill | `kind.CapByBU`, `kind.PortfolioCap`, main report BU grouping and cap remaining, BU scorecard, burndown, defer selector, Ready export. |
| `H` | `Revised Group` | Organization / grouping | Dropdown-compatible group field | `PM_Filter_Dropdowns`, `kind.FutureTierAuto`, main report grouping, scorecard, burndown, Ready export. |
| `I` | `Region` | Organization / grouping | Dropdown-compatible group field | `PM_Filter_Dropdowns`, `Search.Projects_Health`, working-budget screen, burndown detail, Ready export. |
| `J` | `Site` | Organization / grouping | Dropdown-compatible group field | `PM_Filter_Dropdowns`, `Search.Projects_Health`, working-budget screen, burndown detail, Ready export. |
| `K` | `PM` | Owner / grouping | Dropdown-compatible group field | `PM_Filter_Dropdowns`, `Search.Projects_Health`, `Analysis.PM_SPEND_REPORT`, working-budget screen, burndown detail, Ready export. |
| `L` | `Type` | Classification | Optional text | Ready export and starter review context. |
| `M` | `Chargeable` | Chargeability | `Y,N` dropdown rows `3:2000` | `Ready.ChargeableFlag`, `Ready.InternalReady3`, `Ready.InternalJobs_Export`, `Search.Projects_Health`, main report grouped value pack. |
| `N` | `Carry-over` | Planning scope | Text/status input, not Yes/No | `kind.FutureTierAuto`, main report future filters, BU scorecard, burndown, Ready export. |
| `O` | `Annual Projected` | Finance annual | Currency format; required header fill | `get.GetAnnualProjectedPos`, `get.GetProjections`, `kind.CapConsumeMask`, main report totals, all Analysis screens, defer selectors. |
| `P` | `Current Authorized Amount` | Finance annual | Currency format; required header fill | `get.GetCurrentAuthorized`, `kind.CapConsumeMask`, working-budget and defer remaining-authority calculations. |
| `Q` | `January Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `R` | `January Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `S` | `January` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `T` | `February Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `U` | `February Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `V` | `February` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `W` | `March Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `X` | `March Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `Y` | `March` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `Z` | `April Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AA` | `April Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AB` | `April` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AC` | `May Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AD` | `May Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AE` | `May` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AF` | `June Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AG` | `June Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AH` | `June` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AI` | `July Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AJ` | `July Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AK` | `July` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AL` | `August Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AM` | `August Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AN` | `August` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AO` | `September Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AP` | `September Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AQ` | `September` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AR` | `October Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AS` | `October Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AT` | `October` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AU` | `November Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AV` | `November Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AW` | `November` | Finance monthly budget | Currency format | `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `AX` | `December Projected` | Finance monthly projected | Currency format | `get.GetProj12Raw`, `get.GetFinalProj12`, phasing helpers, defer schedule remaining, spend/burndown analysis. |
| `AY` | `December Actuals` | Finance monthly actuals | Currency format | `get.GetActuals12`, YTD actuals, cap consumption, spend reports, scorecard, burndown, defer selectors. |
| `AZ` | `December` | Finance monthly budget | Currency format | `get.GetFinanceEndPos`, `get.GetBudget12`, budget/YTD comparisons, spend signal packs, analysis screens. |
| `BA` | `Comments` | Operator notes | Free text | Ready export context. |
| `BB` | `Stage` | Readiness workflow | Free text | `Ready.InternalReady3`, `Ready.InternalJobs_Export`, `Search.Projects_Health`. |
| `BC` | `Planning Maturity` | Readiness workflow | Free text | `Ready.InternalReady3`, `Ready.InternalJobs_Export`, `Search.Projects_Health`. |
| `BD` | `Planning Notes` | Operator notes | Free text | Ready export context. |
| `BE` | `Internal Eligible` | Internal readiness | `Y,N` dropdown rows `3:2000`; required header fill | `Ready.InternalEligible`, `Ready.InternalJobs_Export` keep mask and output. |
| `BF` | `Earliest Start (Internal)` | Internal scheduling | Month/list-like text | Main report, reforecast queue, defer selector, burndown, Ready export. |
| `BG` | `Blocked Reason (Internal)` | Internal scheduling | Free text | Ready export progress signal and context. |
| `BH` | `Timeline` | Planning policy | Free text/list-like status | `defer.PolicyExcludeFlag`, Ready export progress signal and context. |
| `BI` | `Person` | Resource context | Free text | Ready export and resource-oriented examples. |
| `BJ` | `Hours` | Resource context | Numeric format | Ready export and resource-oriented examples. |
| `BK` | `Future Tier Override` | Future-scope control | Free text/list-like override | `kind.FutureTierFinal`, main report future filters, scorecard, burndown, Ready export. |
| `BL` | `Canceled` | Exclusion flag | `Y,N` dropdown rows `3:2000` | `kind.CapConsumeMask`, main report, all Analysis screens, defer selectors, `Ready.InternalJobs_Export`. |

## Yes/No Columns

The public starter Yes/No source list is `Validation Lists[Yes No]` with values `Y,N`. The Office.js add-in applies these rules from row `3` through row `2000` by matching header names on row `2`, not by hardcoded column letters.

| Column | Header | Input meaning | Primary dependencies |
|---|---|---|---|
| `M` | `Chargeable` | Whether the row is chargeable/internal-labor relevant. | `Ready.ChargeableFlag`, `Ready.InternalReady3`, `Ready.InternalJobs_Export`, `Search.Projects_Health`, and the main report grouped value pack. |
| `BE` | `Internal Eligible` | Internal-work eligibility flag. | `Ready.InternalEligible` resolves this field directly; `Ready.InternalJobs_Export` uses it for the export `Internal Eligible` column and for its keep mask. |
| `BL` | `Canceled` | Explicit cancellation flag separate from `Status`. | `kind.CapConsumeMask` excludes canceled rows with no actuals from the accounting universe; the main report, analysis screens, and defer helpers pass this field into that mask; `Ready.InternalJobs_Export` filters canceled rows out. |

## Dependency Notes

- `Chargeable` is the canonical chargeability field. Do not reintroduce `JobFlag`; readiness and search logic use `Chargeable` by header name.
- `Internal Eligible` is the canonical readiness eligibility field. Do not add a separate visible `Eligible` fallback column to new starter workbooks.
- `Composite Cat` is kept as a manual pre-formula planning-table helper. It is suitable for Excel's built-in sort, remove-duplicates, and Data > Subtotal workflows, and formulas only treat rollup-looking labels such as `Grand Total` or `*Total` as exclusions.
- `Ready.InternalReady3` requires a positive eligibility flag, a ready maturity, an execution-stage value, and `Chargeable = Y` to return its strict ready result.
- `Ready.InternalJobs_Export` keeps non-rollup, non-canceled, eligible rows that have some internal-readiness progress signal, then emits computed `Internal Ready Final` from eligibility, maturity, stage, and chargeability.
- `Search.Projects_Health` reports missing `Chargeable` and flags `Chargeable = Y` when `Job ID` is blank.
- `Canceled` is accepted as `Y`, `yes`, `true`, `x`, `canceled`, or `cancelled` in readiness/accounting masks, but the starter dropdown contract remains `Y,N`.
- `Carry-over` and `Future Tier Override` are not Yes/No columns in the public starter. They are separate planning inputs used by future-filter and analysis logic.

## Formula Footprint Seen In The Reference Parse

Formula patterns observed in the source worksheet:

| Area | Pattern | Public-template note |
|---|---|---|
| Visible row helper | Running visible index with `SUBTOTAL` | Not required by the starter add-in. |
| Composite category | Row composite helper over classification fields | Public starter carries `Composite Cat` as a manual pre-formula field for operator sorting, dedupe, and Excel Data > Subtotal workflows. |
| Monthly actual lookups | Actuals-by-key formulas by month | Public starter stores monthly actual columns and formulas read them by header/finance block. |
| Internal readiness | `Ready.InternalReady3(...)` | Public starter exposes readiness through computed `Internal Ready Final` in `Ready.InternalJobs_Export`; there is no source-table `Internal Ready` override column. |
| Resource output | Resource row spill helper | Public starter carries `Person` and `Hours` fields for resource-oriented examples. |

## Practical Contract

- Keep Yes/No validation header-driven. Do not bind rules to fixed letters except as documentation of the current starter layout.
- Preserve the public starter's `A:BL` width unless the formula modules, add-in validation, docs, and audit checks change together.
- Keep `Chargeable`, `Internal Eligible`, and `Canceled` as the complete source-table Yes/No field set until a later workbook parse proves another Yes/No input should be promoted.
- Keep `Composite Cat` manual and pre-formula; it can support Excel built-in Subtotal workflows without becoming add-in formula output.
- AutoFilter header dropdown behavior from the reference parse remains deferred; this map documents it but does not require the add-in to create filter dropdowns.

