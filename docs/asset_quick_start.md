# Asset Quick Start

Asset workflow is optional.

Start with `Planning Review` and `Analysis Hub` unless you explicitly need project-to-asset tracking. Start with Asset Hub to decide whether assets are needed. Start with Asset Register to enter a simple asset. Do not start with Asset Evidence, Asset State History, or PQ asset sheets.

tblBudgetInput remains the manual/canonical planning input table for this release because refresh is not surfaced. Asset entry is separate from planning-source refresh.

## Do I Need The Asset Workflow?

Use assets only when you need to connect planning rows to physical or logical assets, review candidate assets, track replacements/upgrades, or prepare classified evidence for finance outputs.

If you only need capital planning, leave `AssetWorkflowMode` set to `Off` and ignore `Asset Hub`, `Asset Register`, and `Asset Finance Hub`.

## Start With Asset Hub

When assets are in scope, start with `Asset Hub`.

Choose one mode:

| Mode | Use when | First edit |
|---|---|---|
| `Off` | You only need capital planning | Nothing |
| `Map existing assets` | You already know the asset IDs | `Asset Register`, then `Project-to-Asset Map` |
| `Create candidate assets` | Projects may create new assets | Candidate approval queue and mapping staging |
| `Track replacements/upgrades` | Projects replace or improve assets | Pending project-asset changes |
| `Asset finance from evidence` | You need depreciation or funding outputs | Evidence setup and finance assumptions after mapping works |

## Enter One Simple Asset

To enter one asset, go to `Asset Register`.

Minimum fields:

- `AssetID`
- `AssetName`
- `AssetType`
- `Status`

Helpful optional fields:

- `Site`
- `Location`
- `Owner`
- `Condition`
- `Criticality`
- `ReplacementCost`
- `UsefulLifeYears`
- `LinkedProjectID`

`LinkedProjectID` is optional and advisory. It can use the current workbook project-key dropdown, but blanks and manual IDs are allowed. Entering an asset does not auto-populate `tblAssets` from `tblBudgetInput`, `Planning Table`, or Asset Evidence.

## What Not To Edit First

Do not start by editing:

- Asset Evidence,
- PQ asset evidence sheets,
- `Asset State History`,
- `PQ Asset Evidence Model Inputs`,
- raw intermediate Power Query outputs.

Those are backend or advanced finance surfaces.

## How Apply Asset Mappings Works

Asset mapping writeback is staged. Review rows in `Asset Hub`, mark intentional rows ready in the asset staging tables, then run the optional `Apply Asset Mappings` Office Script in a workbook copy.

The script updates project-to-asset links, change rows, and state-history rows when the target tables exist. It does not replace the planning report math and it does not write to an external database.

## How Asset Finance Is Different

Asset Finance is advanced and requires classified evidence.

`Asset Finance Hub` reads `tblAssetEvidence_ModelInputs` after the asset evidence Power Query path classifies evidence. It does not directly read `Asset Register`, `Project-to-Asset Map`, raw evidence rows, or mapped-only evidence.

Rows feed AssetFinance outputs only when `PresentWithClassifiedEvidence = TRUE`.

## SemanticTwin Is Separate

SemanticTwin is optional even when assets are in scope. Use it only when projects or assets need REC and Brick semantic crosswalk labels for future digital-twin-ready review.

Use REC for buildings, spaces, rooms, real-estate context, and generic assets. Use Brick for equipment, points, sensors, meters, setpoints, commands, and building systems. SemanticTwin is not a full ontology import and it does not complete graph, Fabric, or Azure Digital Twins integration.

## Asset Glossary

| Term | Plain meaning |
|---|---|
| Asset | A durable thing being planned, installed, replaced, upgraded, or tracked |
| ProjectKey | The project/job identifier from planning data |
| AssetID | Stable identifier for an accepted asset |
| Asset Register | List of known assets |
| Project-to-Asset Map | Relationship between planning rows and assets |
| Candidate Asset | Suggested asset row not yet accepted |
| Asset Change | Proposed/applied change such as new, replacement, or upgrade |
| Evidence | Supporting information used to classify or finance an asset |
