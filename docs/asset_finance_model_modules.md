# Asset Finance Model Modules

This is the v0.4 working slice for turning asset evidence into workbook-native finance outputs on branch `codex/asset-finance-model-modules`.

`Automation Setup` remains the Office Script import handoff; this slice adds depreciation, funding requirements, totals, and chart-ready feeds from `qAssetEvidence_ModelInputs`.

## Implemented Bridge

The generated workbook path is:

```text
Governance_Starter.xltx -> Asset Evidence Setup -> qAssetEvidence_ModelInputs -> PQ Asset Evidence Model Inputs / tblAssetEvidence_ModelInputs -> AssetFinance outputs
```

The `AssetFinance` formulas live in `modules/asset_finance.formula.txt`. The generated starter installs these workbook names and output sheets:

| Sheet | Formula |
|---|---|
| `Asset Depreciation` | `=AssetFinance.DEPRECIATION_SCHEDULE` |
| `Asset Funding Requirements` | `=AssetFinance.FUNDING_REQUIREMENTS` |
| `Asset Finance Totals` | `=AssetFinance.FINANCE_TOTALS` |
| `Asset Finance Charts` | `=AssetFinance.CHART_FEEDS` |

`Asset Finance Setup` contains `tblAssetFinanceAssumptions`, sourced from `samples/asset_finance_assumptions_starter.tsv`.

## Operator Flow

In Excel:

1. Open `Governance_Starter.xltx` as a workbook copy.
2. Enter evidence in `Asset Evidence Setup` / `tblAssetEvidenceSource`.
3. Enter rules in `tblAssetEvidenceRules` or reviewed overrides in `tblAssetEvidenceOverrides`.
4. Refresh Power Query.
5. Review `PQ Asset Evidence Status`, `PQ Asset Evidence Mapping Queue`, and `PQ Asset Evidence Model Inputs`.
6. Adjust `Asset Finance Setup` / `tblAssetFinanceAssumptions`.
7. Review `Asset Depreciation`, `Asset Funding Requirements`, `Asset Finance Totals`, and `Asset Finance Charts`.
8. Use `Automation Setup` only if notes writeback is wanted; import `ApplyNotes.ts` through `Automate -> New Script`.

## Evidence Rule

The formulas read `tblAssetEvidence_ModelInputs`, not `tblAssetEvidenceSource`, `tblAssetEvidenceRules`, or `tblAssetEvidenceOverrides`.

Mapped structural hints can support review queues and context, but they do not drive final finance outputs by themselves. `AssetFinance.CLASSIFIED_MODEL_INPUTS` filters to rows where `PresentWithClassifiedEvidence = TRUE`; mapped-only rows stay visible in Power Query review outputs.

## Output Semantics

- Depreciation defaults to straight-line and uses `tblAssetFinanceAssumptions[UsefulLifeYears]`.
- Funding requirements group by `FundingSource`, `ProjectKey`, `AssetId`, and `ClassifiedCategoryName`.
- Totals summarize classified evidence amount, annual depreciation, funding requirement amount, classified evidence count, and classified asset count.
- Chart integration means chart-ready feed tables first; no polished native Excel chart objects are created in this slice.

## v0.4 Assumption Semantics

- `AssetFinance` outputs consume only rows from `tblAssetEvidence_ModelInputs` where `PresentWithClassifiedEvidence = TRUE`.
- v0.4 depreciation supports straight-line behavior only. `DepreciationMethod` is a contract/display field; unsupported or non-straight-line values keep the row visible, preserve the entered method, blank `AnnualDepreciation`, and append `DepreciationIssue`.
- v0.4 funding requirements use full grouped classified amounts. `FundingRequirementRule` is a contract/label field; unsupported values keep the grouped row visible, preserve the entered rule, blank `FundingRequirementAmount`, and append `FundingIssue`.
- `ChartGroup` affects funding chart feed grouping only. Depreciation chart feed rows group by `DepreciationClass`. Chart feeds exclude unsupported rows by reading nonblank `AnnualDepreciation` and `FundingRequirementAmount` values from the formula outputs.

## Source Boundary

Workbook binaries remain generated release artifacts. The source of truth stays in text:

- formula modules under `modules/`,
- Power Query templates under `samples/power-query/asset-evidence/`,
- starter TSVs under `samples/`,
- generated-workbook scripts under `tools/`,
- docs and audit checks in the repo.
