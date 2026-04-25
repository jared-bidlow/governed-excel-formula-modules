# Change Log

## 2026-04-25 - Search helper label cleanup

Semantic change:

- Removed legacy source/job shorthand wording from the search helper and replaced it with public `Job ID` wording.

Minimal diff summary:

- Updated `Search.Projects_Health` messages and local variable names.
- Added audit coverage to forbid legacy source/job shorthand wording.

Visible impact:

- Workbook behavior: health-message wording changed only.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Final public BU cleanup

Semantic change:

- Replaced work-looking BU codes and cap amounts with fictional public sample values.
- Added audit coverage so the removed BU codes and cap amounts cannot return silently.

Minimal diff summary:

- Updated `kind.CapByBU_Keys`, `kind.CapByBU_Vals`, and `kind.CapExCap`.
- Updated starter table BU sample values to `BU-A: Sample Unit` and `BU-B: Sample Unit`.
- Extended public-safety audit checks.

Visible impact:

- Workbook behavior: public sample cap values changed to fictional placeholders.
- Main report totals: can change when using only the public sample workbook data.
- Subtotal flags: can change when using only the public sample workbook data.
- Cap remaining values: can change when using only the public sample workbook data.
- No private workbook should import this public sample cap table without replacing the fictional placeholders.

## 2026-04-25 - Starter workbook table added

Semantic change:

- Added a paste-ready public starter table so new users can create a blank workbook trial without inventing the source-table shape.
- Documented why the current finance block needs three columns per month.

Minimal diff summary:

- Added `samples/planning_table_starter.tsv`.
- Added `docs/starter_workbook.md`.
- Updated README and workbook import map with the starter flow.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- New users get a concrete `Planning Table` shape for local testing.

## 2026-04-25 - Public release hardening second pass

Semantic change:

- Generalized remaining workbook-specific labels to public template names across formula modules, the staged-decision script, docs, and validation tooling.
- Added a public release checklist and strengthened the audit so old private workbook labels, local paths, URLs, email addresses, workbook binaries, and generated artifacts are blocked before export.

Minimal diff summary:

- Replaced old sheet/table/header vocabulary with `Planning Table`, `Planning Review`, `Decision Staging`, `Source ID`, `Job ID`, `Planning Notes`, and `Timeline`.
- Added `docs/public_release_checklist.md`.
- Extended `tools/audit_capex_module.py` with broader public-safety checks.

Visible impact:

- Workbook behavior: label contract changes only for the public template.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- Public release readiness is now checked directly by the audit.

## 2026-04-25 - Public template sanitization started

Semantic change:

- Converted the repo presentation from a private workbook workspace into a public-safe Excel formula-module template.
- Kept formula modules available as examples while removing public docs that named real workbook paths, workbook files, or organization-specific process details.

Minimal diff summary:

- Rewrote the README, operating contract, import map, planning-plugin menu, scenario matrix, and change log around the generic governed-formula-module pattern.
- Replaced workbook-specific lineage and inventory docs with generic public-template guidance.
- Replaced the static audit with a public-safety and formula-contract audit.

Visible impact:

- Workbook behavior: no intended change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- Public-template docs now describe the reusable pattern rather than a private workbook implementation.
