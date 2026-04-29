# Codex And ChatGPT Durable Contract

This repository should make review state explicit instead of relying on a reviewer to rediscover it from a zip.

## Required Release-Handoff Report

Every Codex implementation or release-readiness response should state:

```text
Feature status:
- Built:
- Scaffolded:
- Missing:

Validation:
- audit:
- formula lint:
- feature status:
- review packet:
- git diff --check:

Changed:
- files touched

Not changed:
- main report math
- existing AssetFinance calculation semantics
- workbook binaries
- private data
- real database credentials

Known limitations:
- database import scaffold only, unless a task explicitly implements tenant-specific import
- external database writeback not implemented
- workbook build validation depends on local Excel COM availability
```

## Review Packet

Use:

```bash
python -S tools/build_review_packet.py
```

The generated packet under `release_artifacts/review_packet/` is ignored by Git and can be shared for review. It summarizes branch state, feature status, audit output, formula lint, workbook manifest, Power Query adapters, asset workflow status, and changed files when Git metadata is available.

## Feature Status

Use:

```bash
python -S tools/report_feature_status.py
```

Feature status uses source-controlled evidence in `docs/feature_status.tsv` and separates:

- `Built`: implemented and expected to remain present.
- `Scaffolded`: public-safe template or partial implementation.
- `Missing`: intentionally not built yet.
- `Mismatch`: evidence exists where the expected status says missing, or the evidence pattern is invalid.

The command exits nonzero only when a feature marked `Built` lacks its expected evidence.
