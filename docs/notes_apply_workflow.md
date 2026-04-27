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

The table carries report context, proposed updates, readiness fields, and apply status fields. The workflow is usable without manually copying values from `Planning Review`; the setup creates the staging table shape, and `ApplyNotes` run 1 resizes the table from the current `Planning Review!P:R` inputs while preserving formula-backed context columns.

`ApplyNotes` reads the report rows and `Planning Review!O:R`, filters to rows with `NewPlanningNotes`, `NewTimeline`, or `NewStatus`, records the source worksheet row in `ReviewRow`, then prepares matching `tblDecisionStaging` rows. The review/context/helper columns remain formulas keyed by `ReviewRow` so the table continues to show the current existing notes and resolved target values for the exact Planning Review row that created the staged row.

Fresh setup seeds `Planning Review!P5:R5` when those cells are blank. That public-safe smoke input targets `Sample over-projected work`, which exists in `Planning Table`. Run `ApplyNotes` once to mark the staged row `Prepared`, then run it again to update `Planning Notes`, `Timeline`, `Comments`, and `Status`.

For the smoke path, `Planning Review!P5:R5` is read by `ApplyNotes` run 1 and written into `tblDecisionStaging`.

## ApplyNotes Two-Pass Behavior

The maintained script is `office-scripts/apply_notes.ts`.

Run 1 is prepare:

- refreshes workbook data connections,
- recalculates the workbook,
- scans `Planning Review!P5:R200` for raw note/timeline/status inputs,
- resizes `tblDecisionStaging` to the matching rows,
- records `ReviewRow` and restores formulas for review/context/helper columns,
- marks matching rows `Prepared`,
- records match status in `ApplyStatus`, `AppliedOn`, `ApplyMessage`, and `BudgetRowFound`.

Run 2 is apply:

- applies only rows already marked `Prepared`,
- blocks rows where the budget match is not exactly one,
- blocks duplicate staged rows that would write the same `Planning Table` row in one apply batch,
- skips blank updates,
- writes status and messages back to `tblDecisionStaging`.

Status meanings:

- `Prepared`: row matched exactly one `Planning Table` row and is ready for the second run.
- `Applied`: row was written to `Planning Table`; matching `Planning Review!P:R` inputs were cleared when the source row still matched.
- `Blocked`: row was not eligible to write, usually because the project description was blank, did not match exactly one `Planning Table` row, or duplicated another staged row targeting the same `Planning Table` row.
- `Skipped`: no non-empty target values were available to write.
- `Error`: the script hit a row-specific exception while applying.
- blank status: no current `Planning Review!P:R` input is staged.

A later run with no current `Planning Review!P:R` inputs resets stale non-prepared staging rows to one blank formula-backed row. This keeps `Decision Staging` from showing old applied or blocked rows as current work.

The script writes only to controlled target fields on `Planning Table`:

- `Planning Notes`
- `Timeline`
- `Comments`
- `Status`

After a successful apply, the script clears the matching `Planning Review!P:R` source inputs. Column `O` then surfaces the refreshed existing-note context from `Planning Table`.

## Boundary

Formula modules produce the review data. The Office Script stages the current review inputs and performs the controlled writeback. This release does not add RDF export, SHACL validation, or an external data bridge.
