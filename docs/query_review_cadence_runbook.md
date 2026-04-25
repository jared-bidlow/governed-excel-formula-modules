# Query Review Cadence Runbook

## Purpose

Use this runbook to define when workbook query health is reviewed, who owns the review, what evidence is required, and when an issue escalates beyond operations.

This is an operator workflow. It is not a query-spec document and it does not replace per-query cards.

## Review Types

Maintain three review layers:

1. Weekly review for steady-state monitoring.
2. Month-close review for signoff-critical validation.
3. Ad hoc exception review for issues that cannot wait for the normal cadence.

## Weekly Review

### Trigger

- Scheduled operational review.
- First refresh after new source delivery.
- Any variance that operators notice during normal refresh work.

### Owner

- Primary: workbook operator or reporting operations owner.
- Backup: process lead or reporting lead.

### Input Set

- Query refresh status.
- Top blocking `qQA_*` outputs.
- Material row-count changes.
- New or changed `qDict_*` inputs.
- Any published `qPublish_*` outputs due this cycle.

### Review Checks

- Confirm blocking queries refreshed successfully.
- Compare high-risk query counts to the last clean run.
- Confirm expected worksheet-loaded queries still land where they should.
- Confirm connection-only queries stayed connection-only.
- Note any new rule-source or parameter changes.

### Thresholds

- Any blocking query failure is reviewed the same day.
- Any unexplained material variance is logged for follow-up.
- Any load-setting drift is corrected or escalated before signoff work begins.

### Output And Signoff

- Log a short weekly review note.
- Capture owner, date, major variances, and actions taken.
- Mark issues as `watch`, `fix locally`, or `escalate`.

## Month-Close Review

### Trigger

- Formal close window.
- Any reporting cycle where outputs support finance, executive, or governance signoff.

### Owner

- Primary: designated signoff owner.
- Backup: workbook operator plus functional owner for the output.

### Input Set

- Final `qQA_*` control outputs.
- Published or publish-ready `qPublish_*` outputs.
- Any changed `qDict_*` rule sources since the last close.
- Exception log carried from weekly reviews.

### Review Checks

- Reconfirm load destinations for all signoff-relevant queries.
- Reconfirm exception thresholds and known overrides.
- Verify source snapshot dates and parameter values.
- Confirm downstream reports or decks are using the intended outputs.

### Thresholds

- No blocking control can remain unresolved at signoff.
- Any override must have an owner and written rationale.
- Any unexplained output shift above threshold requires explicit owner review.

### Output And Signoff

- Produce a month-close review note.
- Record the queries reviewed, exceptions accepted, overrides approved, and final signoff owner.

## Ad Hoc Exception Review

### Trigger

- Blocking refresh error.
- Missing worksheet-loaded output.
- Connection-only query unexpectedly loaded to a sheet.
- Sudden unexplained count spike or drop.
- Owner challenge to an existing threshold or exception classification.

### Owner

- Operator opens the review.
- Functional owner decides whether the issue is operational, data-quality, or rule-logic.

### Input Set

- Error text or screenshot.
- Query name and expected load behavior.
- Source snapshot used.
- Prior clean comparison point.
- Safe operator actions already attempted.

### Review Checks

- Can the issue be reproduced once?
- Is the source or parameter state obviously wrong?
- Did the issue come from load drift, source drift, or rule drift?
- Is there a signoff deadline at risk?

### Output And Signoff

- Log the incident in the exception tracker.
- Assign owner and next action.
- Decide whether the issue can return to weekly review or stays escalated.

## Standard Evidence Package

Collect the same evidence every time so handoffs are fast:

- `Run Date / Time`
- `Operator`
- `Query Name`
- `Expected Load Type`
- `Source Snapshot`
- `Observed Symptom`
- `Prior Clean Comparison`
- `Safe Operator Action Taken`
- `Escalated To`
- `Resolution / Next Step`

## Escalation Rules

Escalate beyond operations when:

- one safe retry does not clear the issue,
- a blocking `qQA_*` query remains unhealthy,
- a published `qPublish_*` output cannot be trusted,
- a `qDict_*` rule change materially affects downstream behavior,
- or the operator cannot explain the variance with the available evidence.

## Lightweight Meeting Agenda

Use this sequence for weekly or month-close query review:

1. Blocking refresh failures.
2. Load-setting drift.
3. Material count variances.
4. Rule-source changes.
5. Downstream outputs waiting on signoff.
6. Escalations and owners.
