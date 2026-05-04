# Workbook Left-To-Right Map

Use this file as a generic navigation map for a governed workbook.

## Generic Flow

```text
Start Here -> Source Status -> Data Import Setup -> Planning Table -> refresh/re-sync -> tblBudgetInput -> Planning Review -> Analysis Hub
```

Asset workflow is optional. `AssetsLite` continues from `Analysis Hub` to `Asset Hub` and `Asset Register`. `AssetsFull` continues from `Asset Register` to `Asset Finance Hub`. A reference-only semantic crosswalk edition exists for private extension, but it is not part of the normal operator flow.

## Workbook Areas

| Area | Primary role | Governance note |
|---|---|---|
| Start Here | Front door, workbook-flow table, visible-sheet navigation, source rule, and backend/admin explanation. | Opens first in generated workbooks. |
| Source Status | Visible source-health summary. | Reads hidden import status/issues instead of exposing every QA table. |
| Data Import Setup | Source profile, adapter selection, and 64-column contract. | Public-safe placeholders only; no credentials or private URLs. |
| Planning Table | Manual/staging/local writeback surface. | Not the formula source; refresh or re-sync after edits or ApplyNotes. |
| tblBudgetInput | Canonical formula source. | Lives on hidden `PQ Budget Input`; formulas read this table through `get`. |
| Planning Review | Main report and notes-entry surface. | Preserve report totals, subtotal flags, and cap remaining values unless intentional. |
| Analysis Hub | Scorecards, queues, burndown, working budget, and readiness output. | Includes a clickable `Go to section` table and replaces scattered analysis demo sheets. |
| Asset Hub | Optional asset workflow onboarding, simple asset-entry guidance, mode selection, next actions, and review queues. | Start with Asset Hub to decide whether assets are needed; backend asset sheets stay hidden/admin-scoped by default. |
| Asset Register | Simple manual asset-entry table. | Start with Asset Register to enter a simple asset; `LinkedProjectID` is optional and advisory. |
| Asset Finance Hub | Optional depreciation, funding, totals, and chart-ready feeds. | Asset Finance is advanced and requires classified evidence; reads classified model inputs only. |
| Semantic Map Hub | Reference-only semantic crosswalk and triple-shaped review queue. | Not part of the normal operator flow; this is not a full ontology import or deployed external integration. |

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

`AssetsLite` makes `Asset Hub` and `Asset Register` visible. `AssetsFull` makes `Asset Hub`, `Asset Register`, and `Asset Finance Hub` visible. The reference crosswalk edition makes `Semantic Map Hub` visible only when a private workbook copy explicitly needs that review surface. Do not start with Asset Evidence, Asset State History, or PQ asset sheets.

The hidden backend includes `PQ Budget Input`, `PQ Budget QA`, `Validation Lists`, `Decision Staging`, `Automation Setup`, advanced asset workflow sheets, `Asset Finance Setup`, `Semantic Map Setup`, `Workbook Manifest`, and intermediate asset-evidence Power Query sheets. They are hidden, not deleted, so the workbook remains auditable. `tblWorkbookManifest[Presence]`, `tblWorkbookManifest[Edition]`, and `tblWorkbookManifest[FriendlyName]` mark generated sheets, edition visibility, user-facing labels, and `OptionalLegacy` sheet names.

tblBudgetInput remains the manual/canonical planning input table for this release because refresh is not surfaced. Simple asset entry does not auto-populate `tblAssets` from `tblBudgetInput`, `Planning Table`, or Asset Evidence.

Semantic crosswalk files are reference-only. They are not a full ontology dump and do not claim deployed graph or twin integration.

## Change Routing

| Question | Destination |
|---|---|
| Is this a tracked formula behavior change? | Update `modules/*.formula.txt`, docs, and audit checks. |
| Is this a workbook validation, layout, or local formula change? | Capture a workbook-change packet before applying it in Excel. |
| Is this hidden Name Manager logic? | Inventory it before promoting or changing it. |
| Is this public documentation? | Remove real paths, people, workbook names, and source data. |
