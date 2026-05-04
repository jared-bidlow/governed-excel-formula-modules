# Notes Apply Design

This document is the design checkpoint for the notes-apply workflow introduced at commit `3703af6dd6aa4108de8d71fdb99720629faf7ab3`.

## Release Boundary

`v0.2.0` adds a controlled notes, status, and timeline apply workflow.

It does not include external graph export, external validation, semantic graph publication, asset workflow design, generated workbook artifacts, or workbook binaries.

## Workflow Contract

The workflow moves operator note edits through a staged table before any live planning row is changed:

```text
Planning Review O:R inputs -> tblDecisionStaging -> office-scripts/apply_notes.ts -> Planning Table
```

`Planning Review` is the operator-facing review surface. Columns `O:R` hold proposed edits until they are staged.

`tblDecisionStaging` is the review and safety boundary. It must contain the source `ReviewRow`, row key, proposed values, match status, apply status, and operator-ready fields needed by the script.

`office-scripts/apply_notes.ts` is the controlled apply step. It reads only staged rows, validates row matching and readiness, and writes approved updates to `Planning Table`.

`Planning Table` remains the live workbook source table for planning notes, timeline, comments, and status.

## Controlled Inputs

`Planning Review` columns `O:R` provide the controlled input surface:

- `ExistingMeetingNotes`
- `NewPlanningNotes`
- `NewTimeline`
- `NewStatus`

## Two-Pass State Machine

The apply script is intentionally staged.

Run 1 prepares rows:

- Refresh or recalculate upstream formulas.
- Inspect each `Planning Review!P:R` row with proposed updates.
- Record `ReviewRow` so formula-backed staging columns resolve from the exact source row instead of table row position.
- Write `ApplyStatus = Prepared` for rows with staged input.
- Write `BudgetRowFound` and `ApplyMessage` so the operator can inspect match state before commit.
- Do not write live `Planning Table` values on the prepare pass.

Run 2 applies prepared rows:

- Re-check each row already marked `Prepared`.
- Apply only rows that still satisfy the refusal rules.
- Write `AppliedOn`, `ApplyStatus`, `ApplyMessage`, and `BudgetRowFound`.
- Leave rows prepared when the row is not safe to apply.
- Clear raw proposed inputs only after a successful apply.

| Current row state | Condition | Script action |
| --- | --- | --- |
| Blank / Pending | `ApplyReady` true, valid target, new values exist | Mark `Prepared`; write intended values into `*_New`; do not update `Planning Table` |
| Prepared | Target still uniquely matches | Write to `Planning Table`; mark `Applied`; set timestamp and message |
| Prepared | Target no longer uniquely matches | Mark `Blocked`; do not write |
| Any | More than one staged row targets the same `Planning Table` row | Mark `Blocked`; do not write |
| Any | `BudgetMatchCount` is not `1` | Mark `Blocked`; do not write |
| Any | Blank key field | Mark `Blocked`; do not write |
| Any | No new note, status, or timeline values | Mark `Skipped`; do not write |
| Applied | No changed input | Leave alone |

## Required Staging Fields

`tblDecisionStaging` must expose these fields for the notes-apply contract:

- `ReviewRow`: records the exact `Planning Review` worksheet row that produced the staged update.
- `ApplyAction`: declares which write targets are intended for the row.
- `ApplyReady`: confirms the row has settled and is ready for the script.
- `ApplyStatus`: records staged state such as `Prepared`, `Applied`, or `Error`.
- `AppliedOn`: records the apply timestamp after a successful or failed apply attempt.
- `ApplyMessage`: records the operator-facing reason for prepared, skipped, applied, or error state.
- `BudgetMatchCount`: records how many live `Planning Table` rows match the staged key.
- `BudgetRowFound`: records the matched live row number when exactly one row is found.

## Write Targets

The script may write only these `Planning Table` targets:

- `Planning Notes`
- `Timeline`
- `Comments`
- `Status`

`Comments` may be used to preserve prior note context when `Planning Notes` is replaced. The script must not write back into resolved helper columns in the staging table.

## Refusal Rules

The script must refuse to apply a staged row when any of these conditions is true:

- `BudgetMatchCount` is not `1`.
- The key or project field is blank.
- The row contains no proposed update for the selected `ApplyAction`.
- The live row has not been matched explicitly.
- More than one staged row would write to the same live `Planning Table` row in one apply batch.

The script must not blindly overwrite live values. Every write to `Planning Table` must be tied to a validated staged row and a single matched live row.

## Pass Rule

A user can type notes, status, or timeline updates beside the capital planning report, stage them, run `ApplyNotes` once to prepare, run it again to apply, and see `Planning Table` updated without manual copy/paste.

This design document is a source-control checkpoint. It documents the workflow contract only.
