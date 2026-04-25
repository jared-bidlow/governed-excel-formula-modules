# Workbook Left-To-Right Map

Use this file as a generic navigation map for a governed workbook.

## Generic Flow

```text
Inputs / Lists -> Planning Table -> Actuals + As-Of Month -> Main Report -> Analysis Screens -> Notes / Decisions -> Controlled Writeback
```

## Workbook Areas

| Area | Primary role | Governance note |
|---|---|---|
| Planning table | Main source of project/job records, forecast values, and status fields. | Keep column contracts documented. |
| Lists and controls | Source for dropdowns, slicers, and user choices. | Capture workbook-side changes before applying them. |
| Actuals import | Converts source exports into normalized actual spend. | Keep source paths private; publish only fake examples. |
| Main report | Primary report or meeting surface. | Preserve visible totals and flags unless behavior change is intentional. |
| Analysis screens | Optional planning plugins and diagnostics. | Keep summary pivots separate from ranked detail queues. |
| Notes/writeback | Optional controlled update path. | Treat as high-risk and scenario-test before changing. |

## Change Routing

| Question | Destination |
|---|---|
| Is this a tracked formula behavior change? | Update `modules/*.formula.txt`, docs, and audit checks. |
| Is this a workbook validation, layout, or local formula change? | Capture a workbook-change packet before applying it in Excel. |
| Is this hidden Name Manager logic? | Inventory it before promoting or changing it. |
| Is this public documentation? | Remove real paths, people, workbook names, and source data. |
