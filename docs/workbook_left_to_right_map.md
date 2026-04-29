# Workbook Left-To-Right Map

Use this file as a generic navigation map for a governed workbook.

## Generic Flow

```text
Start Here -> Source Status -> Data Import Setup -> Planning Table -> refresh/re-sync -> tblBudgetInput -> Planning Review -> Analysis Hub
```

Asset workflow is optional. `AssetsLite` continues from `Analysis Hub` to `Asset Hub`. `AssetsFull` continues from `Asset Hub` to `Asset Finance Hub`. `SemanticTwin` adds an optional `Semantic Map Hub` after the asset/finance flow for REC and Brick semantic crosswalk review.

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
| Asset Hub | Optional project-to-asset workflow onboarding, mode selection, next actions, and review queues. | Start with Asset Hub only when asset tracking is in scope; backend asset tables stay hidden/admin-scoped by default. |
| Asset Finance Hub | Optional depreciation, funding, totals, and chart-ready feeds. | Asset Finance is advanced and requires classified evidence; reads classified model inputs only. |
| Semantic Map Hub | Optional REC/Brick semantic crosswalk and triple export queue. | SemanticTwin only; this is not a full ontology import or completed digital-twin integration. |

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
```

`AssetsLite` makes `Asset Hub` visible. `AssetsFull` makes `Asset Hub` and `Asset Finance Hub` visible. `SemanticTwin` makes `Semantic Map Hub` visible for optional REC and Brick mapping. Do not start with PQ asset evidence sheets. Do not start with `Asset State History`.

The hidden backend includes `PQ Budget Input`, `PQ Budget QA`, `Validation Lists`, `Decision Staging`, `Automation Setup`, asset workflow sheets, `Asset Finance Setup`, `Semantic Map Setup`, `Workbook Manifest`, and intermediate asset-evidence Power Query sheets. They are hidden, not deleted, so the workbook remains auditable. `tblWorkbookManifest[Presence]`, `tblWorkbookManifest[Edition]`, and `tblWorkbookManifest[FriendlyName]` mark generated sheets, edition visibility, user-facing labels, and `OptionalLegacy` sheet names.

SemanticTwin is optional. Use REC for buildings, spaces, rooms, real-estate context, and generic assets. Use Brick for equipment, points, sensors, meters, setpoints, commands, and building systems. This is not a full ontology dump and does not claim completed Fabric graph or Azure Digital Twins integration.

## Change Routing

| Question | Destination |
|---|---|
| Is this a tracked formula behavior change? | Update `modules/*.formula.txt`, docs, and audit checks. |
| Is this a workbook validation, layout, or local formula change? | Capture a workbook-change packet before applying it in Excel. |
| Is this hidden Name Manager logic? | Inventory it before promoting or changing it. |
| Is this public documentation? | Remove real paths, people, workbook names, and source data. |
