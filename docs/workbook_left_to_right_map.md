# Workbook Left-To-Right Map

Use this file as a generic navigation map for a governed workbook.

## Generic Flow

```text
Start Here -> Source Status -> Data Import Setup -> Planning Table -> refresh/re-sync -> tblBudgetInput -> Planning Review -> Analysis Hub -> Asset Hub -> Asset Finance Hub
```

## Workbook Areas

| Area | Primary role | Governance note |
|---|---|---|
| Start Here | Front door, workbook-flow table, visible-sheet navigation, source rule, and backend/admin explanation. | Opens first in generated workbooks. |
| Source Status | Visible source-health summary. | Reads hidden import status/issues instead of exposing every QA table. |
| Data Import Setup | Source profile, adapter selection, and 64-column contract. | Public-safe placeholders only; no credentials or private URLs. |
| Planning Table | Manual/staging/local writeback surface. | Not the formula source; refresh or re-sync after edits or ApplyNotes. |
| tblBudgetInput | Canonical formula source. | Lives on hidden `PQ Budget Input`; formulas read this table through `get`. |
| Planning Review | Main report and notes-entry surface. | Preserve report totals, subtotal flags, and cap remaining values unless intentional. |
| Analysis Hub | Scorecards, queues, burndown, working budget, and readiness output. | Includes a section index and replaces scattered analysis demo sheets. |
| Asset Hub | Asset workflow review queues. | Includes a section index; backend asset tables stay hidden/admin-scoped by default. |
| Asset Finance Hub | Depreciation, funding, totals, and chart-ready feeds. | Includes a section index and reads classified model inputs only. |

## Visibility Model

The default visible workbook surface is:

```text
Start Here
Source Status
Data Import Setup
Planning Table
Cap Setup
Planning Review
Analysis Hub
Asset Hub
Asset Finance Hub
```

The hidden backend includes `PQ Budget Input`, `PQ Budget QA`, `Validation Lists`, `Decision Staging`, `Automation Setup`, asset workflow sheets, `Asset Finance Setup`, `Workbook Manifest`, and intermediate asset-evidence Power Query sheets. They are hidden, not deleted, so the workbook remains auditable. `tblWorkbookManifest[Presence]` marks generated sheets separately from `OptionalLegacy` sheet names that are documented but not created as primary workbook surfaces.

## Change Routing

| Question | Destination |
|---|---|
| Is this a tracked formula behavior change? | Update `modules/*.formula.txt`, docs, and audit checks. |
| Is this a workbook validation, layout, or local formula change? | Capture a workbook-change packet before applying it in Excel. |
| Is this hidden Name Manager logic? | Inventory it before promoting or changing it. |
| Is this public documentation? | Remove real paths, people, workbook names, and source data. |
