# Planning Plugins

This page is the public-template menu for optional `Analysis` screens that sit beside a core capital-planning report.

Planning plugins are formula-only views. They do not edit workbook data, approve decisions, or replace an operator's review.

## Plugin Menu

| Plugin | Status | Entry point | Decision it supports |
|---|---|---|---|
| PM Spend Report | Implemented | `Analysis.PM_SPEND_REPORT([groupBy])` | Which existing jobs, PMs, and groups are carrying projected, spent, remaining, and zero-spend work. |
| Working Budget Screen | Implemented | `Analysis.WORKING_BUDGET_SCREEN()` | Which current jobs need carry, cut, authorization, or reforecast review before a working budget is drafted. |
| BU Cap Scorecard | Implemented | `Analysis.BU_CAP_SCORECARD()` | How each business unit is measuring against cap, projected work, YTD budget, and YTD spend. |
| Reforecast Queue | Implemented | `Analysis.REFORECAST_QUEUE([groupBy])` | Which jobs need forecast changes, hold review, or authorization cleanup. |
| Burndown Screen | Implemented | `Analysis.BURNDOWN_SCREEN([groupBy])` | How much current-year work remains under the active meeting filters and which jobs explain it. |
| Cut/Deferral Pack | Future candidate | Future `Analysis` wrapper | Which low-spend or low-lock jobs can reduce near-term burn. |
| Carryover Draft | Future candidate | Future `Analysis` screen | Which remaining work should seed the next planning cycle. |

## PM Spend Report

`Analysis.PM_SPEND_REPORT([groupBy])` is an existing-work summary and detail screen.

Default group:

- `Category`

Fallback behavior:

- If `[groupBy]` is omitted, use `Category`.
- If `[groupBy]` does not match a budget header, fall back to `Category`.
- If `[groupBy]` is `PM`, use `Category` because PM is already the second summary key.

The screen has three sections:

- `Group Totals by <Group>`
- `PM Summary by <Group>`
- `Jobs by <Group>`

Use it when the next planning task needs a job-level source list with projected dollars, YTD spend, remaining dollars, status, and `Pot Skip` projected dollars.

## Working Budget Screen

`Analysis.WORKING_BUDGET_SCREEN()` is a formula-only screening view for current jobs before a separate working-budget table is drafted.

The screen has two sections:

- `Working Budget Totals by Planning Hint`
- `Working Budget Detail`

It classifies rows into planning hints such as `Pot Skip Review`, `Authorized Hold Review`, `Unplanned Spend Review`, `Reforecast / Over Projected`, `Carry Forecast Review`, or `Monitor`.

Use it to choose what needs human review before building a monthly plan. It does not create or update the planning table.

## BU Cap Scorecard

`Analysis.BU_CAP_SCORECARD()` is a BU-level cap and spend posture screen.

It answers:

- what each BU cap is,
- how much in-scope projected work is sitting against that cap,
- how much hidden burn reduces effective BU cap,
- how much projected work is hidden by future or closed-row filters,
- how actual spend compares to YTD budget,
- whether the BU is within, near, or over its cap plan.

The output includes cap, projected, spend, remaining-burn, hidden-work, job-count, cap-posture, and spend-posture columns.

BU caps come from the workbook's `Cap Setup` sheet. Update that table to change caps; do not edit formula modules for cap changes.

## Reforecast Queue

`Analysis.REFORECAST_QUEUE([groupBy])` is a formula-only planning queue.

Default group:

- `BU`

Fallback behavior:

- If `[groupBy]` is omitted, use `BU`.
- If `[groupBy]` does not match a budget header, fall back to `BU`.

The screen has three sections:

- `Reforecast Queue Totals`
- `Reforecast Queue by <Group>, Job, and Action`
- `Reforecast Queue Detail`

`Reforecast Queue Totals` starts with a `Grand Total` decision-dollar row, then lists action-level totals sorted by decision dollars.

`Reforecast Queue by <Group>, Job, and Action` is a `PIVOTBY` matrix. It uses the selected group as the first row key, a synthetic `Site | PM | Project Description` job key as the second row key, action as the column key, and decision dollars as the value. The pivot shows built-in group subtotals and a grand total.

`Reforecast Queue Detail` remains the ranked evidence queue, sorted by priority first, then largest decision dollars, then selected group and project description.

Candidate actions:

| Action | Trigger | Intended decision |
|---|---|---|
| `Raise Forecast` | Report flag evidence contains `Over Projected` | Raise the forecast enough to cover actual spend already incurred. |
| `Add Forecast` | Projected amount is zero and YTD spend is positive | Add a forecast for unplanned actual spend. |
| `Cut / Hold Review` | Projected amount is positive and YTD spend is zero | Decide whether zero-spend projected work should remain, move, or be cut. |
| `Release / Hold Auth` | Authorized amount is positive while projection and spend are zero | Decide whether unused authorization should remain available or be released. |

## Burndown Screen

`Analysis.BURNDOWN_SCREEN([groupBy])` is a meeting view for remaining current-year burn.

Default group:

- `BU`

The screen shows:

- controls in effect, including group, as-of month, future filter, closed-row filter, cut target, jobs in scope, and hidden dollars;
- a director summary by selected group;
- ranked job detail for the rows driving remaining burn.

Use `docs/burndown_screen_runbook.md` for the operator reading guide.

## Interpretation Boundary

The planning plugins are decision-support screens only.

- They do not write back to workbook source tables.
- They do not change main report totals.
- They do not change subtotal flags.
- They do not change cap remaining values.
- A workbook user must still make and apply final planning decisions through their normal controlled process.
