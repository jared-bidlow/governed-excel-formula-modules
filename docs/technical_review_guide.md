# Technical Review Guide

This guide is for a technical reviewer who wants to understand what this repository demonstrates without opening a private workbook or seeing production data.

The short version: this repo treats workbook logic as governed source code. The calculation surface is still Excel, but the formulas, contracts, sample data, validation checks, and installation helper live in files that can be reviewed, diffed, audited, and restarted.

## Review Path

1. Read `README.md` for the repository shape.
2. Read `docs/operating_contract.md` for the workbook logic contract.
3. Inspect `modules/capital_planning_report.formula.txt` and `modules/analysis.formula.txt` for the main report and planning screens.
4. Inspect `tools/audit_capex_module.py` for public-safety checks and formula-contract checks.
5. Inspect `docs/scenario_matrix.md` for behavior coverage.
6. Inspect `docs/starter_workbook.md` and `samples/*.tsv` for the blank-workbook trial path.
7. Inspect `addin/` and `docs/office_addin.md` for the setup and installation layer.

## Systems Pattern

The repo demonstrates a practical pattern for complex workbook environments:

- keep durable logic in text modules instead of hidden workbook state;
- preserve workbook binaries and generated artifacts outside the repo;
- write the operating contract down before optimizing behavior;
- enforce public-safety and formula-contract rules with static checks;
- give operators a starter workbook path instead of only developer notes;
- keep JavaScript as installation and validation tooling, not as the calculation engine.

## Evidence In The Repo

The formula modules show how large workbook logic can be split into named responsibilities:

- `get` reads workbook table and range inputs.
- `kind` owns reusable calculation helpers, cap lookup, flags, and display logic.
- `CapitalPlanning` owns the main report formula.
- `Analysis` owns optional planning screens and review matrices.

The documentation shows the surrounding operating system:

- `docs/operating_contract.md` records the behavior contract.
- `docs/workbook_import_map.md` maps text modules back to workbook names.
- `docs/planning_plugins.md` explains review screens.
- `docs/scenario_matrix.md` records scenario coverage.
- `docs/public_release_checklist.md` preserves the public/private boundary.

The tools keep the repo honest:

- `tools/audit_capex_module.py` checks public safety, required formulas, named-formula budgets, add-in contracts, sample-table shape, and documentation coverage.
- `tools/lint_formulas.py` checks formula files without opening Excel.
- `tools/push_public.ps1` runs validation before committing and pushing.

## Public Boundary

This repo intentionally does not include:

- workbook binaries;
- production data;
- private workbook names;
- company, employee, vendor, customer, or project identifiers;
- local machine paths;
- generated workbook packages.

The public sample data is fictional. The important artifact is the governance pattern, not the sample business vocabulary.

## What To Run

From the repo root:

```powershell
python tools\audit_capex_module.py
python tools\lint_formulas.py modules\*.formula.txt
```

Those checks are text-only. They validate the source-controlled logic and documentation boundary without opening Excel.

For a workbook-facing trial, follow `docs/starter_workbook.md`. For the Office.js installer path, follow `docs/office_addin.md`.
