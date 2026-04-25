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
|   +-- capital_planning_report.formula.txt
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
|   +-- planning_table_starter.tsv
|   \-- cap_setup_starter.tsv
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

To validate, commit, rebase, and push the public repo in one local command:

```powershell
.\tools\push_public.ps1 -Message "Update public formula template"
```

## Why Excel

Excel is the right runtime for this pattern when the real work is planning, review, and decision support by people who already live in workbooks.

The goal is not to pretend that a workbook is a database or a web application. The goal is to make workbook logic governable:

- planners can inspect formulas, source rows, pivots, and exceptions without leaving Excel,
- dynamic arrays can produce live review screens without a separate service layer,
- `LET` and `LAMBDA` allow complex logic to be named, split, and reused,
- Git can track formula modules as source code while excluding workbook binaries,
- text audit can enforce size limits, public-safety rules, and key formula contracts.

A different stack is better when the primary need is multi-user transactions, permissions, APIs, durable storage, or application workflows. This repo is for the space where Excel is already the operating surface, and the missing discipline is source control, reviewability, and repeatable validation.

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
samples/cap_setup_starter.tsv
```

Paste the planning table into `Planning Table!A2`, paste the cap table into `Cap Setup!A2`, set `Planning Review!M2` to a month abbreviation such as `Mar`, then import the formula modules.

## Core Pattern

The example workbook logic is split into modules:

- `get` owns workbook range extraction helpers.
- `kind` owns shared calculation, cap lookup, grouping, flag, and display helpers.
- `CapitalPlanning` owns the main `CAPITAL_PLANNING_REPORT()` formula.
- `Analysis` owns optional planning screens such as `BU_CAP_SCORECARD()` and `REFORECAST_QUEUE([groupBy])`.

Cap limits are workbook inputs. Create a `Cap Setup` sheet, paste `samples/cap_setup_starter.tsv` into `Cap Setup!A2`, and replace the fake caps with your own planning limits. `kind.CapByBU(...)` maps the BU code before any colon in `Planning Table[BU]` to that cap table, and `kind.PortfolioCap` sums the table.

The important implementation idea is the boundary, not the sample vocabulary: keep complex Excel logic in importable text modules, keep workbook binaries out of Git, and make formula behavior reviewable with docs and static checks.
