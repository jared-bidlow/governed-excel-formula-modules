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

## Public Safety

| Scenario | Expected behavior |
|---|---|
| A real company name, local path, or workbook filename appears in docs | Static audit fails. |
| A workbook binary is added to the repo | Static audit fails. |
| Formula modules contain unbalanced brackets or quotes | Formula lint fails. |
