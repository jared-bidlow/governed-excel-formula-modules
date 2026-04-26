# Current Work Addendum

## Current Checkpoint

Current commit: `3703af6dd6aa4108de8d71fdb99720629faf7ab3`

Checkpoint name: notes-apply design checkpoint.

Design sentence: this release makes notes, status, and timeline updates operable through a governed workbook workflow: `Planning Review` inputs -> `tblDecisionStaging` -> `ApplyNotes` two-pass writeback -> `Planning Table`.

## Release Boundary

This checkpoint is documentation-only. It records the `ApplyNotes` workflow contract after the add-in handoff work and before any script implementation pass.

`v0.2.0` is scoped to the controlled notes, status, and timeline apply workflow.

Included:

- Notes-apply design contract.
- Current-work pointer for the active release boundary.

Excluded:

- Office Script edits.
- Add-in runtime edits.
- Formula module edits.
- TSV/sample additions.
- Workbook binaries or generated artifacts.
- Asset workflow design.
- RDF/export scope.
- SHACL validation or semantic graph publication.

## Next Recommended Implementation Task

Implement the notes-apply script against `docs/notes_apply_design.md` in a separate task.

That task should update `office-scripts/apply_notes.ts`, add or update focused static checks for the two-pass state machine and refusal rules, and leave add-in files, formula modules, samples, asset mapping, and export scope out of the change.
