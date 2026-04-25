# Public Release Checklist

Use this checklist before turning the template into a public repository.

## Required Gates

- no private workbook files are tracked or copied into the release export.
- no generated workbook artifacts, workbook packages, screenshots, or sample files from real work are included.
- no employer, customer, vendor, employee, or private project names appear in tracked text.
- no local paths, email addresses, URLs, or live source locations appear in tracked text.
- no real project/job rows, business-unit names, or operational notes are included.
- all workbook-facing names use generic template labels such as `Planning Table`, `Planning Review`, and `Decision Staging`.
- the audit and formula lint pass in the source branch and again in the clean-history export.

## Clean-History Export

Do not publish this working repository directly.

Create a clean-history export from tracked plain-text files only, initialize a fresh Git repository in that export, then run the same release checks again before the first public commit.

Recommended public repo name: `governed-excel-formula-modules`.

## Local Push Helper

From the clean public export repo, run:

```powershell
.\tools\push_public.ps1 -Message "Update public formula template"
```

The helper runs the audit, formula lint, whitespace check, commit, fetch/rebase, and push sequence. It should not be run from a private drafting branch.
