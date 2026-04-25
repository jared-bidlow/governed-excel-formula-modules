# Query Playbook Template

## Related Docs

- Use `docs/query_playbook_starter_cards.md` for first-pass `qQA_*` and `qDict_*` family cards.
- Use `docs/query_review_cadence_runbook.md` for weekly, month-close, and exception-review mechanics.

## Purpose

Use this template to document how operators should run, validate, troubleshoot, and escalate workbook queries without needing the original subject matter expert in the room.

This is an operations artifact, not a code-spec artifact. It should answer:

- what the query is for,
- what it depends on,
- how often it runs,
- how to tell whether it succeeded,
- what operators can safely do when it fails,
- and when to escalate.

## How To Use It

- Create one playbook card per query when the query is operationally important or failure-prone.
- Use one playbook card per query family when multiple queries share the same owner, cadence, and operator workflow.
- Keep each card short enough that an operator can use it during a refresh window.
- Keep examples branch-local. Do not pull workbook-specific query names from private sibling workspaces into this repo.
- Link to deeper technical notes only when the operator truly needs them.

## Query Card Template

### 1. Identity

- `Query Name`
- `Query Family`
- `Business Purpose`
- `Primary Owner`
- `Backup Owner`
- `Workbook / File`

### 2. Inputs And Dependencies

- `External Source`
- `Parameters`
- `Upstream Queries`
- `Local Tables / Named Ranges`
- `Expected Load Type`

Recommended wording:

- `Expected Load Type = Worksheet-loaded`
- `Expected Load Type = Connection-only`

### 3. Cadence And Trigger

- `Refresh Cadence`
- `Run Trigger`
- `Required Preconditions`
- `Expected Refresh Window`

Examples:

- `Refresh Cadence = Weekly`
- `Refresh Cadence = Month-close`
- `Run Trigger = New source file received`
- `Required Preconditions = Source folder populated and parameter path validated`

### 4. Success Checks

List the fastest operator checks that confirm the query is healthy.

- expected load target exists,
- row count is non-zero when expected,
- schema looks right,
- latest source date is current,
- downstream dependent query refreshes cleanly,
- output lands on the expected worksheet or remains connection-only as intended.

### 5. Common Failure Modes

Use a compact table.

| Symptom | Likely Cause | Safe Operator Action | Escalate When |
| --- | --- | --- | --- |
| Query not loaded where expected | Load setting drift | Check load destination and compare against expected load type | Query is loaded incorrectly after reapplying expected settings |
| Empty output | Missing upstream data or broken filter | Confirm source file, date window, and dependency queries | Source exists but output remains empty |
| Refresh error | Path, credential, schema, or type drift | Capture error text and validate source path / parameter | Error persists after one safe retry |
| Unexpected row spike or drop | Rule change, duplicate source, or broken join | Compare against prior run and inspect upstream counts | Variance exceeds agreed threshold |

### 6. Safe Operator Actions

List only actions that operations is allowed to perform without SME approval.

- refresh the query once,
- refresh its immediate dependencies,
- verify source path and parameter values,
- confirm load destination,
- compare row counts to the previous successful run,
- capture screenshots or error text,
- log the issue in the exception tracker.

Do not list unsafe actions here. Those belong in escalation.

### 7. Escalation Rules

Document the exact conditions that move the issue from operations to SME or owner review.

- blocking refresh error on a critical query,
- worksheet-loaded query missing from its expected sheet,
- connection-only query unexpectedly loaded to a sheet,
- output variance above threshold,
- signoff query not ready by deadline,
- operator cannot explain the issue with one safe retry.

### 8. Audit Trail

Capture enough detail that the next person can reconstruct what happened.

- `Run Date / Time`
- `Operator`
- `Source Snapshot`
- `Observed Symptom`
- `Action Taken`
- `Escalated To`
- `Resolution`

## Compact Query Card

```md
### Query: <Query Name>

- Business Purpose: <What business decision or output this supports>
- Owner: <Primary owner>
- Backup Owner: <Backup owner>
- Expected Load Type: <Worksheet-loaded / Connection-only>
- Refresh Cadence: <Weekly / Month-close / Ad hoc>
- Trigger: <When operators should run or inspect it>
- Upstream Dependencies: <Queries, parameters, tables, source files>
- Success Check: <Fastest way to confirm the query is healthy>
- Safe Operator Action: <Allowed retry/check actions>
- Escalate When: <Exact conditions for escalation>
- Downstream Consumer: <Report, sheet, gate, or process that depends on it>
```

## Query Family Starters

These are good defaults when building the first version of a workbook query playbook.

### `qDict_*`

Use for dictionary, mapping, and rule queries.

Focus on:

- who owns the rules,
- how rule changes are approved,
- expected row count range,
- and how to validate that the newest source was actually picked up.

### `qQA_*`

Use for QA and control-gate queries.

Focus on:

- whether failure blocks release,
- what threshold is tolerated,
- what counts as a known exception,
- and who signs off on overrides.

### `qPublish_*`

Use for publishable or downstream-consumed outputs.

Focus on:

- whether the query must be worksheet-loaded,
- what “ready to share” means,
- what downstream deck/report/process consumes it,
- and who approves release.

## Starter Rollout Order

If the workbook has many queries, document them in this order:

1. blocking QA / signoff queries,
2. published output queries,
3. dictionary and rule queries,
4. high-churn transformation queries,
5. low-risk helper queries.

That order gives operations the most leverage fastest.
