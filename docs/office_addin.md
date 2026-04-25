# Office.js Add-In Starter

This repo includes a minimal Excel Office.js task-pane add-in under `addin/`.

The add-in is an installer and validator. It does not replace the formula modules with JavaScript business logic.

## What It Does

- Creates the starter sheets: `Planning Table`, `Cap Setup`, and `Planning Review`.
- Pastes the public starter TSV data into `Planning Table!A2` and `Cap Setup!A2`.
- Reads `modules/*.formula.txt` from the hosted repo root.
- Installs workbook defined names through the Excel JavaScript API.
- Adds module-qualified names such as `kind.CapByBU` and `Analysis.REFORECAST_QUEUE`.
- Adds unqualified compatibility aliases for the first occurrence of each formula name.
- Validates required sheets, names, and starter headers.

## Local Trial Shape

Run the smoke-test helper from the repo root:

```powershell
.\tools\start_addin_smoke_test.ps1
```

With Node/npm installed, this equivalent npm script is also available:

```powershell
npm run addin:smoke
```

The helper:

- runs the static repo checks,
- creates/reuses a local trusted certificate for the server-only fallback,
- starts the local HTTPS server on `localhost` port `3000`,
- asks Excel desktop to sideload `addin/manifest.xml` when npm is available.

After Excel opens, use the task pane button:

```text
Setup + Install + Validate
```

When the test session is done, run:

```powershell
.\tools\stop_addin_smoke_test.ps1
```

or:

```powershell
npm run addin:stop
```

If npm is not installed, or the Office debugging tool is blocked on a machine, use the server-only helper and sideload the manifest manually:

```powershell
.\tools\start_addin_dev_server.ps1
```

The manifest points Excel to:

```text
addin/taskpane.html
```

The task pane reads formula modules and samples by relative path, so it needs the full repo content available from the same hosted root.

## Boundary

The add-in is not a workbook binary, not VBA, and not a calculation engine.

The calculation logic still lives in native Excel named formulas after installation. This keeps the public story aligned with governed formula modules rather than a hidden JavaScript planning engine.

## Production Notes

Before using this as a production add-in:

- replace the local development host in `addin/manifest.xml`,
- decide whether the add-in is internal-only or public Marketplace/AppSource material,
- add real icon assets if required by the deployment channel,
- test sideloading in desktop Excel and Excel for the web,
- keep formula module import validation in `tools/audit_capex_module.py`.
