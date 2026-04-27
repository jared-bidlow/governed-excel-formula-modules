# Asset Finance Model Modules

This is the v0.4 working slice for turning asset evidence into workbook-native finance outputs.

## Status

Started on branch `codex/asset-finance-model-modules`.

The first change in this slice is an `Automation Setup` worksheet in the generated starter template. It closes the release usability gap for the optional `ApplyNotes.ts` Office Script by documenting the import path inside the workbook artifact.

## Planned Model Modules

The next implementation targets are:

- depreciation,
- funding requirements,
- totals,
- chart-ready feeds.

These should use `qAssetEvidence_ModelInputs` as the Power Query bridge. The formulas should consume evidence rows that have already preserved the mapped-vs-classified distinction:

- mapped structural hints can support review queues and context,
- true classified evidence requires classifier metadata,
- `PresentWithClassifiedEvidence` should not be set by folder/context mapping alone.

## Source Boundary

Workbook binaries remain generated release artifacts. The source of truth stays in text:

- formula modules under `modules/`,
- Power Query templates under `samples/power-query/asset-evidence/`,
- starter TSVs under `samples/`,
- generated-workbook scripts under `tools/`,
- docs and audit checks in the repo.
