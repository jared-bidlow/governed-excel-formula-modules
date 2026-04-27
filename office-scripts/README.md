# Office Scripts

This folder contains controlled workbook writeback actions for the public starter workbook.

The formula modules and Office.js add-in prepare review queues, staging tables, and validation. Office Scripts are the only artifacts in this repo that write staged workbook changes back into source tables.

Scripts:

- `apply_notes.ts` stages note edits from `Planning Review!P:R` into formula-backed `Decision Staging`, then applies prepared rows into `Planning Table`.
- `apply_asset_mappings.ts` applies accepted asset setup rows into asset mapping, change, and state-history tables.
- Formula modules create review queues; Office Scripts perform controlled writes.
- RDF/export is not part of this release.

Operator boundary:

- Run setup from the add-in before running scripts.
- Review staged rows before applying them.
- Keep workbook binaries and private workbook data out of this repo.
- Treat script output as workbook state, not source-controlled data.
