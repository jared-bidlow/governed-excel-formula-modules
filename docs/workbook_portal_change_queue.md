# Workbook Change Queue

Use this file as a generic queue for workbook-side changes that cannot be represented only by text formula modules.

## Scope

- Repo-safe changes: formula modules, docs, scripts, and audit rules.
- Workbook-side changes: data validation, workbook-only named formulas, table formulas, worksheet layout, list ranges, and other Excel state.
- Do not commit workbook binaries to this repo.

## Packet Standard

Each workbook-side change should be captured before it is applied:

- Status
- Workbook placeholder name
- Sheet or table
- Target range or named formula
- Current state
- Proposed change
- Reason
- Performance impact
- Usability impact
- Expected report impact
- Validation check
- Rollback

## Example Packet

### Add controlled group dropdown to planning rows

Status: `Example`

Workbook: `SamplePlanningWorkbook.xlsx`

Sheet: `Planning Table`

Target: `ExampleGroupColumn`

Proposed change:

- Add list validation that points to a small source spill on a setup sheet.

Expected report impact:

- Visible report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

Validation check:

- Confirm existing values remain valid.
- Confirm dropdown selection does not create noticeable recalculation delay.
