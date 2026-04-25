# Query Playbook Starter Cards

## Purpose

Use these starter cards when a workbook needs operational coverage quickly and the exact per-query cards do not exist yet.

These are family-level defaults. They are not workbook-specific truth and should be replaced by exact query cards once an owner documents the real workflow.

## How To Use These Cards

- Copy the closest family card into a workbook-local playbook.
- Replace the placeholders with the real owner, trigger, load behavior, and thresholds.
- Keep the family card only as a temporary operating fallback.
- Do not treat these defaults as signoff rules until the workbook owner confirms them.

## Starter Card: `qQA_*`

Use this card for QA, control-gate, and release-blocking queries.

```md
### Query Family: qQA_*

- Business Purpose: Surface exceptions, control failures, and release-blocking conditions before a report, deck, or review package is treated as ready.
- Owner: <QA owner or workbook operator>
- Backup Owner: <reporting lead or process owner>
- Expected Load Type: <Worksheet-loaded when used as a visible review gate; connection-only when supporting an upstream gate>
- Refresh Cadence: Weekly during active operations and mandatory at month-close
- Trigger: After source refresh, before signoff, or when an exception threshold is questioned
- Upstream Dependencies: Source intake queries, parameter tables, mapping or dictionary queries, and any upstream staging queries
- Success Check: Output exists in the expected destination, exception counts are explainable, and no blocking control has changed state unexpectedly
- Safe Operator Action: Refresh once, refresh immediate dependencies, validate source path and parameter values, compare counts to the last clean run, capture error text, and log the issue
- Escalate When: A blocking exception remains after one safe retry, the query is missing from its expected load target, a variance exceeds threshold without explanation, or the issue cannot be summarized clearly by the operator
- Downstream Consumer: Signoff meeting, control review, exception queue, or release gate
```

### Common Failure Patterns For `qQA_*`

| Symptom | Likely Cause | Safe Operator Action | Escalate When |
| --- | --- | --- | --- |
| Exception count spikes | Upstream schema drift, duplicate source rows, or stale rules | Compare to prior run and inspect upstream row counts | Spike remains unexplained after one retry |
| Query missing from expected worksheet | Load destination drift | Recheck expected load type and destination | Query still lands incorrectly after settings are corrected |
| Empty output on an active review cycle | Source not loaded, bad filter window, or broken dependency | Validate source snapshot and dependency refresh order | Source is present but output stays empty |
| Refresh error blocks signoff | Path, credential, or schema problem | Capture error text and verify path or parameters | Error persists after one safe retry |

## Starter Card: `qDict_*`

Use this card for rule, mapping, and dictionary queries.

```md
### Query Family: qDict_*

- Business Purpose: Provide approved rule tables, mappings, overrides, or exclusions that downstream classification and QA queries rely on
- Owner: <rules owner or process SME>
- Backup Owner: <workbook operator or reporting lead>
- Expected Load Type: <Connection-only unless operators are expected to review the rule table directly>
- Refresh Cadence: On source change, before controlled refreshes, and at month-close if rule changes affect signoff outputs
- Trigger: Rule workbook updated, override table edited, or downstream classification behavior changes unexpectedly
- Upstream Dependencies: Rule source workbook, local override tables, parameter paths, and any staging queries that normalize the source
- Success Check: Latest rule source is present, row count is in expected range, required key columns are populated, and downstream queries pick up the newest rules
- Safe Operator Action: Refresh once, confirm source path and file timestamp, compare row count to the last clean run, verify required columns exist, and capture a sample of changed rows
- Escalate When: Required keys disappear, downstream classification shifts materially, row count change is unexplained, or a rule change affects release-blocking outputs
- Downstream Consumer: Classification queries, QA gates, review outputs, and published reporting tables
```

### Common Failure Patterns For `qDict_*`

| Symptom | Likely Cause | Safe Operator Action | Escalate When |
| --- | --- | --- | --- |
| Row count drops sharply | Source file replaced, filter drift, or missing sheet/table | Check source file timestamp and structure | Count remains out of range after source validation |
| Required keys are blank | Broken import, schema rename, or malformed source rows | Inspect key columns and compare to prior source snapshot | Key blanks affect downstream joins or classifications |
| Downstream classifications shift unexpectedly | Rule edit, override drift, or join mismatch | Sample changed rows and compare to prior successful run | Shift exceeds agreed threshold or changes signoff outputs |
| Query refreshes but downstream still looks stale | Dependency order issue or stale load target | Refresh direct dependents and confirm updated timestamp | Dependents still show old behavior after refresh |

## Suggested First Rollout

If a workbook has no query operations docs yet, start in this order:

1. Document one `qQA_*` family card for release-blocking controls.
2. Document one `qDict_*` family card for rule ownership and change handling.
3. Split family cards into exact query cards once thresholds, load settings, and owners are agreed.
