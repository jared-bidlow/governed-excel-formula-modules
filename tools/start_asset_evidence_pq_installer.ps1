[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $PSScriptRoot "build_asset_evidence_pq_seed.ps1"
$installScript = Join-Path $PSScriptRoot "install_asset_evidence_pq_workbook.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-DefaultOutputPath {
    param([string]$InputPath)

    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        return ""
    }

    $directory = Split-Path -Parent $InputPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $extension = [System.IO.Path]::GetExtension($InputPath)
    return Join-Path $directory "$baseName.asset-evidence-pq$extension"
}

function Append-Status {
    param(
        [System.Windows.Forms.TextBox]$StatusBox,
        [string]$Message
    )

    $StatusBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $Message`r`n")
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Asset Evidence Power Query Installer"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(760, 430)
$form.MinimumSize = New-Object System.Drawing.Size(720, 400)

$targetLabel = New-Object System.Windows.Forms.Label
$targetLabel.Text = "Workbook copy"
$targetLabel.Location = New-Object System.Drawing.Point(16, 20)
$targetLabel.Size = New-Object System.Drawing.Size(120, 22)
$form.Controls.Add($targetLabel)

$targetText = New-Object System.Windows.Forms.TextBox
$targetText.Location = New-Object System.Drawing.Point(140, 18)
$targetText.Size = New-Object System.Drawing.Size(470, 24)
$form.Controls.Add($targetText)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Location = New-Object System.Drawing.Point(620, 16)
$browseButton.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($browseButton)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Output workbook"
$outputLabel.Location = New-Object System.Drawing.Point(16, 58)
$outputLabel.Size = New-Object System.Drawing.Size(120, 22)
$form.Controls.Add($outputLabel)

$outputText = New-Object System.Windows.Forms.TextBox
$outputText.Location = New-Object System.Drawing.Point(140, 56)
$outputText.Size = New-Object System.Drawing.Size(580, 24)
$form.Controls.Add($outputText)

$replaceCheck = New-Object System.Windows.Forms.CheckBox
$replaceCheck.Text = "Replace existing asset evidence sheets/queries in the output copy"
$replaceCheck.Location = New-Object System.Drawing.Point(140, 92)
$replaceCheck.Size = New-Object System.Drawing.Size(420, 24)
$form.Controls.Add($replaceCheck)

$forceCheck = New-Object System.Windows.Forms.CheckBox
$forceCheck.Text = "Overwrite output file if it already exists"
$forceCheck.Location = New-Object System.Drawing.Point(140, 120)
$forceCheck.Size = New-Object System.Drawing.Size(320, 24)
$form.Controls.Add($forceCheck)

$buildButton = New-Object System.Windows.Forms.Button
$buildButton.Text = "Build Seed"
$buildButton.Location = New-Object System.Drawing.Point(140, 156)
$buildButton.Size = New-Object System.Drawing.Size(130, 34)
$form.Controls.Add($buildButton)

$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "Install Asset Evidence PQ"
$installButton.Location = New-Object System.Drawing.Point(284, 156)
$installButton.Size = New-Object System.Drawing.Size(210, 34)
$form.Controls.Add($installButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Location = New-Object System.Drawing.Point(510, 156)
$closeButton.Size = New-Object System.Drawing.Size(100, 34)
$form.Controls.Add($closeButton)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Location = New-Object System.Drawing.Point(16, 210)
$statusBox.Size = New-Object System.Drawing.Size(704, 160)
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($statusBox)

$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*"
    $dialog.Title = "Select workbook copy"
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $targetText.Text = $dialog.FileName
        $outputText.Text = Get-DefaultOutputPath $dialog.FileName
    }
})

$targetText.Add_TextChanged({
    if ([string]::IsNullOrWhiteSpace($outputText.Text) -or $outputText.Text.EndsWith(".asset-evidence-pq.xlsx")) {
        $outputText.Text = Get-DefaultOutputPath $targetText.Text
    }
})

$buildButton.Add_Click({
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Append-Status $statusBox "Building seed workbook..."
        & $buildScript
        Append-Status $statusBox "Seed workbook ready."
    } catch {
        Append-Status $statusBox "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Build Seed Failed", "OK", "Error") | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$installButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($targetText.Text)) {
            throw "Select a workbook copy first."
        }
        if ([string]::IsNullOrWhiteSpace($outputText.Text)) {
            $outputText.Text = Get-DefaultOutputPath $targetText.Text
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Append-Status $statusBox "Building seed workbook..."
        & $buildScript

        $args = @("-TargetWorkbookPath", $targetText.Text, "-OutputPath", $outputText.Text)
        if ($replaceCheck.Checked) {
            $args += "-ReplaceExisting"
        }
        if ($forceCheck.Checked) {
            $args += "-Force"
        }

        Append-Status $statusBox "Installing into output workbook copy..."
        & $installScript @args
        Append-Status $statusBox "Done: $($outputText.Text)"
        [System.Windows.Forms.MessageBox]::Show($form, "Asset evidence Power Query installed into:`r`n$outputText.Text", "Install Complete", "OK", "Information") | Out-Null
    } catch {
        Append-Status $statusBox "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Install Failed", "OK", "Error") | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$closeButton.Add_Click({
    $form.Close()
})

Append-Status $statusBox "Select a workbook copy, then click Install Asset Evidence PQ."
[void]$form.ShowDialog()
