# README First

Use this when you downloaded the repo ZIP and want to run the Excel add-in against your own workbook copy.

## Minimum Steps

1. Extract the ZIP folder.
2. Open the extracted folder.
3. Right-click `Start-AddIn.ps1`.
4. Click `Run with PowerShell`.
5. Confirm you are using a workbook copy.
6. When Excel opens, open your workbook copy.
7. In the add-in task pane, click `Setup + Install + Validate + Outputs`.
8. Click `Copy ApplyNotes Script`.
9. In Excel, go to `Automate -> New Script`.
10. Paste the script and save it as `ApplyNotes`.

## Safety Rule

Do not click setup or apply buttons in a production workbook. Use a workbook copy first.

## If PowerShell Says npm Is Missing

Install Node.js LTS, close PowerShell, then run `Start-AddIn.ps1` again.

## After Setup

Use `Planning Review` as the meeting surface. Type updates in `P:R`, run `ApplyNotes` once to prepare, inspect `Decision Staging`, then run `ApplyNotes` again to apply.
