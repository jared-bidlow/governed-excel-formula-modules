# Change Log

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
