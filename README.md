# Governed Excel Formula Modules

This repository is a public template for treating complex Excel workbook logic as source code.

It keeps workbook formulas in plain-text modules, pairs them with scenario documentation, and validates the module set with static audit tools. The pattern is designed for capital planning and forecast-review workbooks, but the repo does not include any real workbook, company data, or generated Excel artifacts.

## What This Demonstrates

- Excel `LAMBDA` / `LET` modules tracked as text.
- Dynamic-array planning screens built with functions such as `GROUPBY`, `PIVOTBY`, `FILTER`, `SORTBY`, `HSTACK`, and `VSTACK`.
- A workbook-binary boundary: formulas, docs, and tests are versioned; `.xlsx` files are not.
- Static checks for formula balance, named-formula size, public-safety strings, and important planning-screen contracts.
- Operator-facing docs that explain how a workbook user should import, review, and validate formula modules.

## Layout

```text
governed-excel-formula-modules/
+-- AGENTS.md
+-- README.md
+-- modules/
|   +-- get.formula.txt
|   +-- kind.formula.txt
|   +-- Capex_Tracker_2026.formula.txt
|   +-- analysis.formula.txt
|   \-- supporting workbook modules
+-- docs/
|   +-- operating_contract.md
|   +-- planning_plugins.md
|   +-- public_release_checklist.md
|   +-- scenario_matrix.md
|   +-- starter_workbook.md
|   +-- workbook_import_map.md
|   \-- change_log.md
+-- samples/
|   \-- planning_table_starter.tsv
\-- tools/
    +-- audit_capex_module.py
    \-- lint_formulas.py
```

## Quick Checks

From the repo root:

```bash
python tools/audit_capex_module.py
python tools/lint_formulas.py modules/*.formula.txt
```

The audit is intentionally text-only. It does not open Excel, edit workbook binaries, or require workbook data.

## Current Readiness

This repo works as a public source-code template:

- the text audit and formula lint pass from a clean checkout,
- the formula modules are importable plain-text examples,
- the starter table gives a blank workbook the expected source-table shape,
- public-safety checks block private labels, paths, workbook binaries, and old sample codes.

This repo is not a turnkey workbook:

- it does not ship an `.xlsx` file,
- it does not automate Excel Name Manager import,
- it does not prove runtime recalculation inside every Excel tenant,
- a real workbook owner still needs to map their own table names, headers, caps, and review process.

## Start From A Blank Workbook

For a first local trial, create a blank workbook and follow `docs/starter_workbook.md`.

The paste-ready starter table is in:

```text
samples/planning_table_starter.tsv
```

Paste it into `Planning Table!A2`, set `Planning Review!M2` to a month abbreviation such as `Mar`, then import the formula modules.

## Core Pattern

The example workbook logic is split into modules:

- `get` owns workbook range extraction helpers.
- `kind` owns shared calculation and display helpers.
- `Capex_Tracker_2026` owns the main cap-feasibility report formula.
- `Analysis` owns optional planning screens such as `BU_CAP_SCORECARD()` and `REFORECAST_QUEUE([groupBy])`.

The important implementation idea is the boundary, not the sample vocabulary: keep complex Excel logic in importable text modules, keep workbook binaries out of Git, and make formula behavior reviewable with docs and static checks.
