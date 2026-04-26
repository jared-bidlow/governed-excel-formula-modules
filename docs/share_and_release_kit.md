# Share And Release Kit

This repo does not track workbook binaries. Share workbook files as release assets or external downloads, not as committed source files.

## Release Candidate

Use the current starter contract checkpoint as the first public workbook-template release.

Suggested tag:

```text
v0.1.1-starter-workbook
```

Suggested release title:

```text
Starter workbook installer and governed formula modules
```

Suggested release summary:

```text
This release demonstrates a governed Excel planning workbook pattern: formula modules live as source-controlled text, an Office.js task pane can create a starter workbook layout, and static audit tools validate the public-safe contract.

The workbook logic remains native Excel formulas after installation. The repository does not include private workbook data or tracked workbook binaries.
```

## Release Assets

Attach these files to the GitHub Release, not to the repository:

| Asset | Required | Notes |
|---|---|---|
| `governed-capital-planning-demo.xlsx` | Optional but recommended | Sanitized workbook created from a blank workbook using `Setup + Install + Validate + Outputs`. |
| `governed-capital-planning-preview.pdf` | Optional | Short LinkedIn-friendly preview or walkthrough. |

Do not attach a workbook until these checks pass:

- No private company, employee, vendor, customer, project, job, source-system, or path data.
- No hidden sheets with private data.
- No external workbook links to private paths.
- `Setup + Install + Validate + Outputs` has been run successfully.
- `Planning Table` is the public starter shape, `A:BL`.
- `Planning Table` has no source-table `JobFlag`, `Eligible`, or `Internal Ready` column.
- `Internal Jobs` shows computed `Internal Ready Final`.

## LinkedIn Post Draft

```text
I built a public Excel formula-module template for capital planning workflows.

The main idea: keep complex workbook logic in source-controlled text modules, use an Office.js task pane to create a clean starter workbook, and validate the workbook contract with static checks.

What it demonstrates:
- Excel LET/LAMBDA modules as source code
- Dynamic-array planning screens for review workflows
- A blank-workbook setup flow with controls, dropdowns, validation, and demo outputs
- A hard boundary between public formula logic and private workbook data

This is not a SaaS app and not a workbook binary repo. It is a governed Excel pattern for teams that already live in workbooks but need better reviewability and repeatability.

Repo: github.com/jared-bidlow/governed-excel-formula-modules
Release: add the GitHub release link here after publishing
```

## Sharing Workflow

1. Create the sanitized demo workbook outside the repository.
2. Export a short PDF preview or use workbook screenshots for LinkedIn.
3. Create a GitHub Release from `v0.1.1-starter-workbook`.
4. Attach the sanitized `.xlsx` and optional preview PDF as release assets.
5. Publish the LinkedIn post with the repo and release links.
