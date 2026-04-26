# Git worktree workflow starter

Use `main` as the stable product branch and use linked worktrees for parallel tasks. A worktree is a Git-managed working directory connected to the same repository history, not a Codex-only concept.

The operating idea is concurrency, not branch replacement. Keep one clean folder for the public product state, one folder for active formula/add-in work, one folder for review, one folder for automated checks, and one folder for disposable Excel reference analysis.

## Working Model

| Worktree role | Branch | Excel-work use |
|---|---|---|
| `main` | `main` | Pristine public template state: release tags, README, docs, audited formula modules, add-in source, and no workbook binaries. |
| `work` | `codex/<task>` | One focused repo change: formula module edits, Office.js setup/validation behavior, starter TSV changes, docs, or audit rules. |
| `review` | `review/<branch>` | Inspect a PR, compare formula-module diffs, read workbook contract docs, or review generated structure maps without disturbing active work. |
| `fuzz` | `codex/fuzz-<topic>` | Run static audit, formula lint, JavaScript syntax checks, add-in smoke helpers, or generated stress checks away from human edits. |
| `scratch` | detached or `scratch/<topic>` | Temporary reference parsing, notes from an uploaded workbook, or design experiments that may be discarded or promoted to `work`. |

This follows the practical lesson from the Git worktree discussion: worktrees are not a replacement for branches; they are a way to manage concurrent tasks without stashing or switching one working directory in place.

In this repo, the normal day-to-day setup is:

```text
governed-excel-formula-modules/              main, pristine
governed-excel-formula-modules-ready-fix/    codex/ready-fix, formula or add-in work
governed-excel-formula-modules-review-pr-12/ review/pr-12, review only
governed-excel-formula-modules-fuzz-smoke/   codex/fuzz-smoke, automated checks
governed-excel-formula-modules-scratch-map/  scratch/workbook-map, disposable reference analysis
```

## Role-Specific Starters

From the stable repo folder, create an active work branch for normal implementation:

```powershell
.\tools\new_worktree.ps1 -Name ready-fix
```

This creates a sibling folder and branch:

```text
branch: codex/ready-fix
path:   ../governed-excel-formula-modules-ready-fix
```

Then work from the new folder:

```powershell
cd ..\governed-excel-formula-modules-ready-fix
git status
```

For review-only or automated-check roles, keep the branch prefix explicit:

```powershell
.\tools\new_worktree.ps1 -Name pr-12 -BranchPrefix review
.\tools\new_worktree.ps1 -Name fuzz-smoke
.\tools\new_worktree.ps1 -Name workbook-map -BranchPrefix scratch
```

Use `scratch` for local workbook-reference notes only. Do not copy workbook binaries into the repo, and do not treat a scratch worktree as sanitized public material.

## Finish A Worktree

Before merging back:

```powershell
python tools\audit_capex_module.py
python tools\lint_formulas.py modules\*.formula.txt
node --check addin\taskpane.js
git diff --check
```

Merge by pull request when the branch is public-facing. For a local-only fast-forward, first confirm there is no divergence:

```powershell
git fetch origin
git rev-list --left-right --count origin/main...origin/codex/ready-fix
```

`0    N` means `main` can be fast-forwarded to the feature branch. Anything else needs review.

## Remove A Finished Worktree

Use Git to remove it so Git also cleans its worktree metadata:

```powershell
git worktree remove ..\governed-excel-formula-modules-ready-fix
git worktree prune
```

Do not delete worktree folders by hand unless you are prepared to run `git worktree prune` or `git worktree repair` afterward.

## Rules For This Repo

- Keep release tags on `main`.
- Keep workbook binaries out of Git; use release assets or local workbook storage.
- Use one task per branch/worktree.
- Treat workbook copies as local operator artifacts, not repo artifacts.
- Promote scratch findings only by writing sanitized docs, TSV samples, formula modules, add-in source, or audit checks.
- Commit or discard work before removing a worktree.
- Prefer a WIP commit over an unnamed stash when context must be saved.
- Do not use linked worktrees to hide private data; public safety still requires sanitized files and audit checks.
