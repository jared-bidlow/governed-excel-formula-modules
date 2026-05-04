# Integration Bridge Contract

This optional bridge lets a workbook owner hand project identity to a separate review workspace, then bring approved evidence links back into Excel as advisory context.

The bridge is source-controlled as tables, sample CSV shapes, and Power Query templates. It does not ship workbook binaries, private paths, credentials, or production data.

In a local three-repo setup, keep the durable operator map and no-copy handoff command in the integration repo, for example `<LOCAL_INTEGRATION_REPO_V1>\docs\operator_cross_repo_map.md` and `<LOCAL_INTEGRATION_REPO_V1>\scripts\run_operator_handoff.ps1`. Use those files as the source of truth for the CSV handoff sequence.

## Boundary

```text
tblBudgetInput -> financial project register export -> review workspace -> approved evidence import -> advisory workbook context
```

The workbook remains the owner of financial project identity, planning status, caps, notes, and formula outputs. The review workspace owns manual approval decisions for evidence-to-project relationships.

If the integration repo produces asset scorecards or approved asset-project links, those outputs are advisory relationship context. The workbook still owns `ProjectKey` and official planning status; asset score signals or approved asset links must not create projects or update official status.

The bridge does not:

- create official financial projects,
- update official project status from documentation text,
- use raw file paths as project keys,
- treat candidate evidence mappings as approved rows,
- overwrite manual review decisions during a refresh.

## Project Register Export

Export these columns from `tblFinancialProjectRegisterExport`:

| Column | Rule |
|---|---|
| `Source ID` | Comes from `tblBudgetInput[Source ID]`. |
| `Job ID` | Comes from `tblBudgetInput[Job ID]`. |
| `ProjectKey` | Derived for the bridge as `Source ID & "-" & Job ID`. |
| `Project Description` | Comes from `tblBudgetInput[Project Description]`. |
| `Status` | Workbook planning status, exported as context only. |
| `BU` | Business-unit grouping context. |
| `Category` | Planning category context. |
| `Site` | Site context. |
| `PM` | Project manager context. |

The Power Query template `samples/power-query/integration-bridge/qBridge_FinancialProjectRegister.m` shapes this table from `tblBudgetInput`.

## Approved Evidence Import

Power Query import is the normal path for approved evidence. The Finance workbook owns the import into `Integration Bridge!tblApprovedProjectEvidence`; the integration repo only generates `data\exports\approved_project_evidence.csv`.

Configure `tblIntegrationBridgeConfig` with:

| Setting | Purpose |
|---|---|
| `IntegrationRepoRoot` | Local path to the Integration repo or extracted operator package root. |
| `ApprovedProjectEvidenceCsvRelativePath` | Relative path to the approved output. Default: `data\exports\approved_project_evidence.csv`. |
| `RefreshApprovedEvidenceOnOpen` | Optional setting. Refresh-on-open is not enabled by default; the normal operator action is manual refresh after the Integration handoff completes. |

The Power Query template `samples/power-query/integration-bridge/qBridge_ApprovedProjectEvidence.m` reads:

```text
<IntegrationRepoRoot>\data\exports\approved_project_evidence.csv
```

If the CSV is missing, empty, or header-only, the query returns an empty typed table with the approved-evidence headers so workbook refresh does not break. Manual paste into `tblApprovedProjectEvidence` is fallback only.

Import approved rows into `tblApprovedProjectEvidence` with these columns:

| Column |
|---|
| `ProjectKey` |
| `EvidenceId` |
| `EvidenceType` |
| `EvidencePath` |
| `EvidenceName` |
| `Extension` |
| `DocumentAreaID` |
| `DocumentAreaName` |
| `CategoryID` |
| `CategoryName` |
| `DateModified` |
| `ReviewStatus` |
| `ApprovedOn` |
| `ReviewerNotes` |
| `StatusSignal` |

Only rows with `ReviewStatus = Approved` should be used as trusted advisory evidence. `StatusSignal` is documentation context only and must not change the official workbook status.

The Power Query template preserves the approved-evidence shape and filters to approved rows.

## Refresh Safety

The review workspace should preserve decisions by `ReviewKey = EvidenceId & "|" & CandidateProjectKey`. The workbook side only receives approved output rows keyed by `ProjectKey` and `EvidenceId`.

Candidate mappings, review decisions, and approved exports should remain separate tables. A refresh may regenerate candidates, but it must not erase manual approval history.

## Operator Flow

1. Refresh or re-sync `tblBudgetInput`.
2. Review `Integration Bridge` / `tblFinancialProjectRegisterExport`.
3. Run the integration repo handoff command to write `tblFinancialProjectRegisterExport` as `financial_project_register.csv`.
4. Refresh the Finance workbook Power Query import so `qBridge_ApprovedProjectEvidence` pulls `approved_project_evidence.csv` into `tblApprovedProjectEvidence`.
5. Use manual paste only as a fallback when Power Query refresh is unavailable.
6. Treat approved evidence rows as context for review, not as commands to create projects or update status.
