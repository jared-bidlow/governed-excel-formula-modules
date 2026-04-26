# Power PnP guidance proposal

This repository supports PowerPnPGuidanceHub issue #110:

Govern operational Excel workbooks with source-controlled formula modules

The repo is supporting evidence for a guidance pattern, not the proposed publication artifact itself.

The proposed guidance describes how operational Excel workbooks can be governed with:

- source-controlled formula modules
- workbook contract documentation
- static validation checks
- an Office.js setup/validation add-in
- sanitized starter data
- a public/private data boundary

The workbook remains the operator-facing surface. The add-in installs and validates native Excel workbook names; it does not replace Excel as the calculation engine.
