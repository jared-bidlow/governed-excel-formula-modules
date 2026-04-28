# Copilot Review Playbook

Copilot support in this repo means prompt cards and clean review surfaces. It does not make Copilot the calculation engine.

## Boundary

Copilot may explain, summarize, classify text, and draft review narratives. Copilot must not be the source of governed numeric totals. Use native Excel formulas for numerical tasks requiring accuracy or reproducibility.

## Workbook Surfaces

Use Copilot against tables and output sheets that are already produced by deterministic workbook logic:

- `tblBudgetInput`
- `tblBudgetImportStatus`
- `tblBudgetImportIssues`
- `Planning Review`
- `Reforecast Queue`
- `BU Cap Scorecard`
- `Source Status`
- `Asset Finance Totals`
- `Asset Finance Charts`

Prompt cards are tracked in:

```text
samples/copilot_prompt_cards.tsv
```

## Example Review Uses

- Summarize budget import issues by severity and review status.
- Draft a meeting agenda from the reforecast queue.
- Explain which columns are missing from the canonical import contract.
- Summarize unsupported AssetFinance assumption rows.
- Draft a plain-language note from reviewed formula outputs.

Always review generated narrative before sharing it. Governed numbers remain the native Excel formula outputs.
