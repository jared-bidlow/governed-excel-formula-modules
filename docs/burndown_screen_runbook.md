# Burndown Screen Runbook

Use `Analysis.BURNDOWN_SCREEN([groupBy])` as the meeting burndown view when a director or operations lead needs to understand how much 2026 work remains, which filters are currently in effect, and which jobs are driving the remaining burn.

## What This Screen Is For

The screen is a summarized meeting view, not the job-based defer selector and not the full cap-feasibility report. It answers:

- How much projected work is in scope.
- How much has burned through the selected as-of month.
- How much remains by the selected group.
- How much work is hidden because of future or closed-row controls.
- Which largest remaining jobs explain the burn.

## Controls In Effect

The first section, `Burndown Controls In Effect`, is there so the meeting does not start with hidden assumptions.

- `Group` is the active grouping passed to `BURNDOWN_SCREEN([groupBy])`; default is `BU`.
- `As Of Month` comes from `'Planning Review'!$M$2`.
- `Future Filter` shows the current `Future_Filter_Mode` setting.
- `Closed Rows` shows the current `HideClosed_Status` setting.
- `Cut Target` shows `Burndown_Cut_Target`; blank or invalid input is treated as `0`.
- `Jobs in Scope` is the count of rows that survived the current meeting filters.
- `Hidden Dollars` and `Hidden Dollars %` show projected dollars hidden by future or closed controls.

## How To Read The Director Summary

The `Burndown by ...` section is fixed at 15 columns so it stays readable for non-technical users.

- Start with `Remaining to Burn`, `% Burned`, and `Jobs in Scope`.
- Use `Current Remaining`, `F1 Committed Remaining`, `F2 Future Remaining`, and `F3/Other Future Remaining` to explain the future-tier mix.
- Use `Hidden by Future $` and `Hidden by Closed $` to explain why visible scope is smaller than the full cap-consuming universe.
- Treat `Remaining Share %` as the group's share of visible remaining burn, not as a cap-feasibility test.

## When To Use The Analyst Helpers

The wide helper outputs remain available for diagnosis and power-user review:

- `BURNDOWN_WS_HORIZON([groupBy])` shows the packed `PIVOTBY` view by future tier.
- `BURNDOWN_WS_STATUS([groupBy])` shows the packed `PIVOTBY` view by status.
- `BURNDOWN_WS_SIGNAL([groupBy])` shows the packed `PIVOTBY` view by scope state.

Use those helpers when someone needs to investigate a bucket, not as the default meeting surface.

## What Not To Over-Interpret

- This screen does not replace `CAPEX_REPORT` for cap-feasibility decisions.
- Hidden dollars explain meeting filters; they are not automatically removed from cap accountability.
- `Cut Candidate` is a meeting aid for a target cut amount; it is not an approval workflow.
- Future-tier buckets depend on the current future-tier helper logic and any `Future Tier Override` values in the workbook.
