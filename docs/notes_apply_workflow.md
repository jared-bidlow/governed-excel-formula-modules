# Notes Apply Workflow

The notes workflow keeps meeting edits visible in Excel while making writeback explicit and controlled.

## Planning Review Notes Columns

`Setup Notes Workflow` adds a notes block next to the main report on `Planning Review`.

| Column | Header | Purpose |
|---|---|---|
| `O` | `ExistingMeetingNotes` | Helper/read-only context from the current report note state. |
| `P` | `NewPlanningNotes` | Operator input for the next planning-note value. |
| `Q` | `NewTimeline` | Operator input for the next timeline value. |
| `R` | `NewStatus` | Operator input for the next status value, dropdown-backed when validation lists are available. |

`ExistingMeetingNotes` is a helper column. It is formatted differently from the input columns and is not the place to type new decisions. Operators type new values in `NewPlanningNotes`, `NewTimeline`, and `NewStatus`.

## Decision Staging

The add-in creates `Decision Staging` and refreshes `tblDecisionStaging` with the staging columns expected by the Office Script.

The table carries report context, proposed updates, readiness fields, and apply status fields. The workflow is usable without manually copying values from `Planning Review`; the setup creates the staging table shape and the notes formulas create the review queue.

## ApplyNotes Two-Pass Behavior

The maintained script is `office-scripts/apply_notes.ts`.

Run 1 is prepare:

- refreshes workbook data connections,
- recalculates the workbook,
- marks rows with raw note inputs as `Prepared`,
- records match status in `ApplyStatus`, `AppliedOn`, `ApplyMessage`, and `BudgetRowFound`.

Run 2 is apply:

- applies only rows already marked `Prepared`,
- skips rows where the budget match is not exactly one,
- skips blank updates,
- writes status and messages back to `tblDecisionStaging`.

The script writes only to controlled target fields on `Planning Table`:

- `Planning Notes`
- `Timeline`
- `Comments`
- `Status`

After a successful apply, the script clears or marks the raw new-note inputs so the applied state is inspectable.

## Boundary

Formula modules produce the review and staging data. The Office Script performs the controlled writeback. This release does not add RDF export, SHACL validation, or an external data bridge.
