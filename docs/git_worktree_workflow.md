# Git worktree workflow starter

Use `main` as the stable product branch and use linked worktrees for parallel tasks. A worktree is a Git-managed working directory connected to the same repository history, not a Codex-only concept.

## Working Model

| Worktree role | Branch | Use |
|---|---|---|
| Product | `main` | Stable public state, releases, tags, and final docs. |
| Feature | `codex/<task>` | One focused implementation or documentation task. |
| Review | `review/<branch>` | Inspect someone else's branch without disturbing active work. |
| Scratch | detached or `scratch/<topic>` | Experiments that may be discarded. |

This follows the practical lesson from the Git worktree discussion: worktrees are not a replacement for branches; they are a way to manage concurrent tasks without stashing or switching one working directory in place.

## Create A Feature Worktree

From the stable repo folder:

```powershell
.\tools\new_worktree.ps1 -Name install-docs
```

This creates a sibling folder and branch:

```text
branch: codex/install-docs
path:   ../governed-excel-formula-modules-install-docs
```

Then work from the new folder:

```powershell
cd ..\governed-excel-formula-modules-install-docs
git status
```

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
git rev-list --left-right --count origin/main...origin/codex/install-docs
```

`0    N` means `main` can be fast-forwarded to the feature branch. Anything else needs review.

## Remove A Finished Worktree

Use Git to remove it so Git also cleans its worktree metadata:

```powershell
git worktree remove ..\governed-excel-formula-modules-install-docs
git worktree prune
```

Do not delete worktree folders by hand unless you are prepared to run `git worktree prune` or `git worktree repair` afterward.

## Rules For This Repo

- Keep release tags on `main`.
- Keep workbook binaries out of Git; use release assets.
- Use one task per branch/worktree.
- Commit or discard work before removing a worktree.
- Prefer a WIP commit over an unnamed stash when context must be saved.
- Do not use linked worktrees to hide private data; public safety still requires sanitized files and audit checks.
