# Scenario Matrix

Use this as a lightweight validation checklist after formula changes.

## Main Report

| Scenario | Expected behavior |
|---|---|
| Cap-consuming active work exists | The row remains in the report universe. |
| In-service row has no projected or actual dollars | The row is excluded as deadweight. |
| In-service row has actual spend above projection | Hidden burn contributes to BU cap context. |
| BU subtotal exceeds effective cap | Over-cap flag and shadow price appear on the subtotal row. |
| `Cap Setup` value changes for a BU | BU cap remaining and cap posture update without editing formula modules. |
| Detail row has normal active spend within projection | Detail flags remain line-level and no reforecast action is created by that fact alone. |

## PM Spend Report

| Scenario | Expected behavior |
|---|---|
| `[groupBy]` is omitted | `Analysis.PM_SPEND_REPORT()` groups by `Category`. |
| `[groupBy]` is a valid non-PM header such as `Region` | Group totals, PM summary rows, and job detail use that selected group. |
| `[groupBy]` is `PM` | The screen falls back to `Category` because PM is already the second summary key. |
| A projected job has zero YTD spend | The detail row shows `Pot Skip` and carries the projected amount into `Pot Skip Projected`. |

## Working Budget Screen

| Scenario | Expected behavior |
|---|---|
| Projected amount is positive and YTD spend is zero | The row is marked `Pot Skip Review`. |
| Authorized amount is positive while projection and spend are zero | The row is marked `Authorized Hold Review`. |
| Projection is zero and YTD spend is positive | The row is marked `Unplanned Spend Review`. |
| YTD spend exceeds projection | The row is marked `Reforecast / Over Projected`. |
| No working-budget candidates exist | The screen returns `No working-budget candidates found in the active accounting universe`. |

## BU Cap Scorecard

| Scenario | Expected behavior |
|---|---|
| BU has cap-consuming rows | `Analysis.BU_CAP_SCORECARD()` returns one row for that BU plus a grand total row. |
| Future filter excludes future rows | Future projected dollars move out of `In-Scope Projected` and into `Hidden by Future`. |
| Closed-row control hides closed rows | Closed projected dollars move out of `In-Scope Projected` and into `Hidden by Closed`. |
| BU in-scope projected dollars exceed effective BU cap | `Cap Posture` is `Over Cap Plan` and `Plan Cap Remaining` is negative. |
| BU in-scope projected dollars use at least 95% of effective BU cap | `Cap Posture` is `Near Cap Plan` unless already over cap. |
| YTD spend exceeds YTD budget | `Spend Posture` is `Ahead of YTD Budget` and `Spend vs YTD Budget` is positive. |
| YTD budget is zero and YTD spend is positive | `Spend Posture` is `Unplanned Spend`. |

## Reforecast Queue

| Scenario | Expected behavior |
|---|---|
| Positive projection with report flag evidence containing `Over Projected` | `Analysis.REFORECAST_QUEUE([groupBy])` classifies the row as `Raise Forecast`. |
| Zero projection with positive YTD spend | The row is classified as `Add Forecast`, not `Raise Forecast`. |
| Positive projection with zero YTD spend | The row is classified as `Cut / Hold Review`. |
| Positive authorization with zero projection and zero YTD spend | The row is classified as `Release / Hold Auth`. |
| Normal active spend within projection | The row is not promoted into the reforecast queue. |
| Reforecast candidates exist | `Reforecast Queue Totals` starts with a `Grand Total` decision-dollar row before action totals. |
| Multiple candidate jobs exist under one selected group | `Reforecast Queue by <Group>, Job, and Action` shows job rows plus a built-in group subtotal. |
| Reforecast detail has multiple candidates | Detail rows sort by priority, then largest decision dollars, then group and project description. |

## Burndown Screen

| Scenario | Expected behavior |
|---|---|
| `[groupBy]` is omitted | `Analysis.BURNDOWN_SCREEN()` groups by `BU`. |
| Rows survive the active meeting filters | The screen starts with `Burndown Controls In Effect`, then shows the group summary and ranked detail. |
| No rows survive the active meeting filters | The screen returns `No cap-consuming rows found for meeting burndown`. |
| Future or closed-row controls hide projected work | Hidden dollars appear in the controls section and the group summary. |
| A cut target is entered | Candidate rows are labeled against the cut target without changing source data. |

## Dropdown Application Data

| Scenario | Expected behavior |
|---|---|
| Starter setup runs | `Planning Review!B2:D2` get dropdowns from the centralized `applicationData` model. |
| Default controls are unchanged | Baseline behavior stays `BU`, `All`, `SHOW`, cut target `0`, and month `Mar`. |
| Group, future, or closed-row controls change | The report and Analysis screens react through the sheet-linked workbook names. |
| `Planning Table` contains a `Chargeable` header | Rows `3:2000` receive a `Y,N` dropdown by header-driven validation. |
| The `Chargeable` header is missing | Workbook validation fails before treating the starter layout as valid. |

## Ready And Row Flags

| Scenario | Expected behavior |
|---|---|
| `Chargeable` is set to `Y` | `Search` can require a job id, and `Ready.InternalJobs_Export` uses `Chargeable` as the chargeability input for its example internal-ready calculation. |
| `Ready.ChargeableFlag` is evaluated | The helper finds the `Chargeable` column by header name instead of assuming column `O`. |
| `Ready.InternalEligible` is evaluated | The helper finds `Internal Eligible` by header name; no separate visible `Eligible` fallback column is present in the starter. |
| The starter workbook is created | No `JobFlag` column and no source-table `Internal Ready` column are present; `Chargeable` is the only starter chargeability flag. |
| Demo outputs are inserted | The `Internal Jobs` sheet spills `=Ready.InternalJobs_Export()` at `A4` for readiness review. |
| Internal jobs are exported | `Ready.InternalJobs_Export` emits computed `Internal Ready Final` from eligibility, maturity, stage, and chargeability instead of reading a manual override column. |
| Operators use Excel Data > Subtotal | `Composite Cat` remains available as a manual pre-formula helper for sort, dedupe, and subtotal workflows without becoming formula output. |
| Readiness behavior is changed later | Formula logic, this scenario section, and audit checks are updated in the same pass before `Ready` is treated as production-validated. |

## Public Safety

| Scenario | Expected behavior |
|---|---|
| A real company name, local path, or workbook filename appears in docs | Static audit fails. |
| A workbook binary is added to the repo | Static audit fails. |
| Formula modules contain unbalanced brackets or quotes | Formula lint fails. |
