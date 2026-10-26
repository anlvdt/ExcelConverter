# Excel Converter v1.2.0 - Windows desktop Excel batch conversion.
# Source files are preserved by default. Long synced paths use local staging.
# See docs/AUDIT-2026-10-26.md for findings and verification limits.
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelConverter.Core.psm1') -Force

# Hide console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Colors
$primaryColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$successColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$errorColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
$bgColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$cardColor = [System.Drawing.Color]::White
$textColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$subtextColor = [System.Drawing.Color]::FromArgb(100, 100, 100)

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Converter"
$form.Size = New-Object System.Drawing.Size(700, 640)
$form.AllowDrop = $true
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $bgColor
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Header Panel
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size = New-Object System.Drawing.Size(700, 80)
$pnlHeader.BackColor = $primaryColor
$form.Controls.Add($pnlHeader)

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = "Excel Converter"
$lblHeader.Location = New-Object System.Drawing.Point(25, 15)
$lblHeader.Size = New-Object System.Drawing.Size(300, 30)
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = [System.Drawing.Color]::White
$lblHeader.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblHeader)

$lblSubHeader = New-Object System.Windows.Forms.Label
$lblSubHeader.Text = "Batch XLS/XLSM/XLSB to XLSX | Long synced paths supported"
$lblSubHeader.Location = New-Object System.Drawing.Point(25, 45)
$lblSubHeader.Size = New-Object System.Drawing.Size(450, 20)
$lblSubHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSubHeader.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
$lblSubHeader.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblSubHeader)

# Scan XLSTART button
$btnScanXLStart = New-Object System.Windows.Forms.Button
$btnScanXLStart.Text = "Scan XLSTART"
$btnScanXLStart.Location = New-Object System.Drawing.Point(530, 25)
$btnScanXLStart.Size = New-Object System.Drawing.Size(110, 30)
$btnScanXLStart.FlatStyle = "Flat"
$btnScanXLStart.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 158)
$btnScanXLStart.ForeColor = [System.Drawing.Color]::White
$btnScanXLStart.FlatAppearance.BorderSize = 0
$btnScanXLStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnScanXLStart.Cursor = "Hand"
$pnlHeader.Controls.Add($btnScanXLStart)

# About button
$btnAbout = New-Object System.Windows.Forms.Button
$btnAbout.Text = "?"
$btnAbout.Location = New-Object System.Drawing.Point(650, 25)
$btnAbout.Size = New-Object System.Drawing.Size(30, 30)
$btnAbout.FlatStyle = "Flat"
$btnAbout.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
$btnAbout.ForeColor = [System.Drawing.Color]::White
$btnAbout.FlatAppearance.BorderSize = 0
$btnAbout.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnAbout.Cursor = "Hand"
$pnlHeader.Controls.Add($btnAbout)

# Global cancel variable
$script:cancelProcessing = $false

# Folder Queue Card
$pnlQueue = New-Object System.Windows.Forms.Panel
$pnlQueue.Location = New-Object System.Drawing.Point(20, 95)
$pnlQueue.Size = New-Object System.Drawing.Size(645, 140)
$pnlQueue.BackColor = $cardColor
$form.Controls.Add($pnlQueue)

$lblQueueTitle = New-Object System.Windows.Forms.Label
$lblQueueTitle.Text = "Batch Queue (Drag & Drop folders/files here)"
$lblQueueTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblQueueTitle.Size = New-Object System.Drawing.Size(400, 20)
$lblQueueTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblQueueTitle.ForeColor = $textColor
$pnlQueue.Controls.Add($lblQueueTitle)

$lstQueue = New-Object System.Windows.Forms.ListBox
$lstQueue.Location = New-Object System.Drawing.Point(15, 38)
$lstQueue.Size = New-Object System.Drawing.Size(510, 85)
$lstQueue.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lstQueue.BorderStyle = "FixedSingle"
$lstQueue.AllowDrop = $true
$lstQueue.SelectionMode = "MultiExtended"
$pnlQueue.Controls.Add($lstQueue)

$btnQueueAdd = New-Object System.Windows.Forms.Button
$btnQueueAdd.Text = "Add Folder"
$btnQueueAdd.Location = New-Object System.Drawing.Point(535, 38)
$btnQueueAdd.Size = New-Object System.Drawing.Size(95, 32)
$btnQueueAdd.FlatStyle = "Flat"
$btnQueueAdd.BackColor = $primaryColor
$btnQueueAdd.ForeColor = [System.Drawing.Color]::White
$btnQueueAdd.FlatAppearance.BorderSize = 0
$btnQueueAdd.Cursor = "Hand"
$pnlQueue.Controls.Add($btnQueueAdd)

$btnQueueClear = New-Object System.Windows.Forms.Button
$btnQueueClear.Text = "Clear"
$btnQueueClear.Location = New-Object System.Drawing.Point(535, 78)
$btnQueueClear.Size = New-Object System.Drawing.Size(46, 32)
$btnQueueClear.FlatStyle = "Flat"
$btnQueueClear.BackColor = [System.Drawing.Color]::FromArgb(230,230,230)
$btnQueueClear.ForeColor = $textColor
$btnQueueClear.FlatAppearance.BorderSize = 0
$btnQueueClear.Cursor = "Hand"
$pnlQueue.Controls.Add($btnQueueClear)

$btnQueueFiles = New-Object System.Windows.Forms.Button
$btnQueueFiles.Text = "Files"
$btnQueueFiles.Location = New-Object System.Drawing.Point(584, 78)
$btnQueueFiles.Size = New-Object System.Drawing.Size(46, 32)
$pnlQueue.Controls.Add($btnQueueFiles)
$btnQueueFiles.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel workbooks|*.xls;*.xlsm;*.xlsb;*.xlsx"
    $dialog.Multiselect = $true
    try {
        if ($dialog.ShowDialog() -eq 'OK') {
            foreach ($path in $dialog.FileNames) {
                if (-not $lstQueue.Items.Contains($path)) { $lstQueue.Items.Add($path) | Out-Null }
            }
        }
    } finally { $dialog.Dispose() }
})

# Options Card
$pnlOptions = New-Object System.Windows.Forms.Panel
$pnlOptions.Location = New-Object System.Drawing.Point(20, 245)
$pnlOptions.Size = New-Object System.Drawing.Size(645, 100)
$pnlOptions.BackColor = $cardColor
$form.Controls.Add($pnlOptions)

$lblOptionsTitle = New-Object System.Windows.Forms.Label
$lblOptionsTitle.Text = "Options"
$lblOptionsTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblOptionsTitle.Size = New-Object System.Drawing.Size(200, 20)
$lblOptionsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblOptionsTitle.ForeColor = $textColor
$pnlOptions.Controls.Add($lblOptionsTitle)

$chkMoveOld = New-Object System.Windows.Forms.CheckBox
$chkMoveOld.Text = "Move original files to 'old' folder"
$chkMoveOld.Location = New-Object System.Drawing.Point(15, 40)
$chkMoveOld.Size = New-Object System.Drawing.Size(260, 22)
$chkMoveOld.Checked = $false
$chkMoveOld.ForeColor = $textColor
$chkMoveOld.Cursor = "Hand"
$pnlOptions.Controls.Add($chkMoveOld)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Include subfolders"
$chkRecursive.Location = New-Object System.Drawing.Point(15, 65)
$chkRecursive.Size = New-Object System.Drawing.Size(260, 22)
$chkRecursive.Checked = $true
$chkRecursive.ForeColor = $textColor
$chkRecursive.Cursor = "Hand"
$pnlOptions.Controls.Add($chkRecursive)

$chkShowExcel = New-Object System.Windows.Forms.CheckBox
$chkShowExcel.Text = "Show Excel window"
$chkShowExcel.Location = New-Object System.Drawing.Point(300, 40)
$chkShowExcel.Size = New-Object System.Drawing.Size(135, 22)
$chkShowExcel.Checked = $false
$chkShowExcel.ForeColor = $textColor
$chkShowExcel.Cursor = "Hand"
$pnlOptions.Controls.Add($chkShowExcel)

$chkDeleteOld = New-Object System.Windows.Forms.CheckBox
$chkDeleteOld.Text = "Delete original files (no backup)"
$chkDeleteOld.Location = New-Object System.Drawing.Point(300, 65)
$chkDeleteOld.Size = New-Object System.Drawing.Size(260, 22)
$chkDeleteOld.Checked = $false
$chkDeleteOld.ForeColor = $errorColor
$chkDeleteOld.Cursor = "Hand"
$pnlOptions.Controls.Add($chkDeleteOld)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout (s)"
$lblTimeout.Location = New-Object System.Drawing.Point(435, 43)
$lblTimeout.Size = New-Object System.Drawing.Size(80, 22)
$pnlOptions.Controls.Add($lblTimeout)
$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(520, 40)
$numTimeout.Size = New-Object System.Drawing.Size(85, 22)
$numTimeout.Minimum = 5
$numTimeout.Maximum = 86400
$numTimeout.Value = 180
$pnlOptions.Controls.Add($numTimeout)

# Log Card
$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Location = New-Object System.Drawing.Point(20, 355)
$pnlLog.Size = New-Object System.Drawing.Size(645, 175)
$pnlLog.BackColor = $cardColor
$form.Controls.Add($pnlLog)

$lblLogTitle = New-Object System.Windows.Forms.Label
$lblLogTitle.Text = "Activity Log"
$lblLogTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblLogTitle.Size = New-Object System.Drawing.Size(200, 20)
$lblLogTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblLogTitle.ForeColor = $textColor
$pnlLog.Controls.Add($lblLogTitle)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 38)
$txtLog.Size = New-Object System.Drawing.Size(615, 125)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
$txtLog.BorderStyle = "FixedSingle"
$txtLog.ForeColor = $textColor
$pnlLog.Controls.Add($txtLog)

# Status Bar
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Location = New-Object System.Drawing.Point(20, 540)
$pnlStatus.Size = New-Object System.Drawing.Size(645, 40)
$pnlStatus.BackColor = $cardColor
$form.Controls.Add($pnlStatus)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(15, 10)
$lblStatus.Size = New-Object System.Drawing.Size(150, 20)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatus.ForeColor = $subtextColor
$pnlStatus.Controls.Add($lblStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(170, 8)
$progressBar.Size = New-Object System.Drawing.Size(255, 22)
$progressBar.Style = "Continuous"
$pnlStatus.Controls.Add($progressBar)

$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Text = "CONVERT"
$btnConvert.Location = New-Object System.Drawing.Point(535, 5)
$btnConvert.Size = New-Object System.Drawing.Size(95, 30)
$btnConvert.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnConvert.BackColor = $successColor
$btnConvert.ForeColor = [System.Drawing.Color]::White
$btnConvert.FlatStyle = "Flat"
$btnConvert.FlatAppearance.BorderSize = 0
$btnConvert.Cursor = "Hand"
$pnlStatus.Controls.Add($btnConvert)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "CANCEL"
$btnCancel.Location = New-Object System.Drawing.Point(435, 5)
$btnCancel.Size = New-Object System.Drawing.Size(95, 30)
$btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnCancel.BackColor = $errorColor
$btnCancel.ForeColor = [System.Drawing.Color]::White
$btnCancel.FlatStyle = "Flat"
$btnCancel.FlatAppearance.BorderSize = 0
$btnCancel.Cursor = "Hand"
$btnCancel.Enabled = $false
$pnlStatus.Controls.Add($btnCancel)

# Footer with author info
$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text = "v1.2.0 | Le An (Vietnam IT) | Ref: gist.github.com/gabceb/954418"
$lblFooter.Location = New-Object System.Drawing.Point(20, 588)
$lblFooter.Size = New-Object System.Drawing.Size(645, 18)
$lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblFooter.ForeColor = $subtextColor
$lblFooter.TextAlign = "MiddleCenter"
$form.Controls.Add($lblFooter)

function Write-Log {
    param([string]$Message, [string]$Type = "Info")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status {
    param([string]$Text, [string]$Type = "Normal")
    $lblStatus.Text = $Text
    switch ($Type) {
        "Success" { $lblStatus.ForeColor = $successColor }
        "Error" { $lblStatus.ForeColor = $errorColor }
        "Working" { $lblStatus.ForeColor = $primaryColor }
        default { $lblStatus.ForeColor = $subtextColor }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Event Handlers
$btnScanXLStart.Text = "Reports"
$btnScanXLStart.Add_Click({
    $directory = Join-Path $env:LOCALAPPDATA 'ExcelConverter\Reports'
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    Start-Process explorer.exe -ArgumentList ('"' + $directory + '"')
})

$btnAbout.Add_Click({
        $aboutText = "EXCEL CONVERTER v1.2.0`r`n"
        $aboutText += "Author: Le An (Vietnam IT)`r`n"
        $aboutText += "`r`n"
        $aboutText += "FEATURES:`r`n"
        $aboutText += "-----------------------------------------`r`n"
        $aboutText += "[Conversion]`r`n"
        $aboutText += "  - Convert XLS (Excel 97-2003) to XLSX`r`n"
        $aboutText += "  - Convert XLSM/XLSB to XLSX`r`n"
        $aboutText += "  - Batch convert folders with subfolders`r`n"
        $aboutText += "  - Move or delete original files`r`n"
        $aboutText += "`r`n"
        $aboutText += "[Batch reliability]`r`n"
        $aboutText += "  - Long local OneDrive/SharePoint sync paths`r`n"
        $aboutText += "  - Verified XLSX output; no overwrite`r`n"
        $aboutText += "  - Per-file timeout, cancellation and CSV report`r`n"
        $aboutText += "  - VBA execution disabled; macros removed in XLSX`r`n"
        $aboutText += "  - Conversion is not an antivirus scan`r`n`r`n"
        $aboutText += "-----------------------------------------`r`n"
        $aboutText += "DONATE (Support the developer):`r`n"
        $aboutText += "-----------------------------------------`r`n"
        $aboutText += "MB Bank: 0360126996868`r`n"
        $aboutText += "Name: LE VAN AN`r`n"
        $aboutText += "`r`n"
        $aboutText += "Momo: 0976896621`r`n"
        $aboutText += "-----------------------------------------"
        [System.Windows.Forms.MessageBox]::Show($aboutText, "About Excel Converter", "OK", "Information")
    })

$chkDeleteOld.Add_CheckedChanged({
        if ($chkDeleteOld.Checked) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Warning: This will permanently delete original files!`n`nAre you sure?",
                "Confirm Delete", "YesNo", "Warning"
            )
            if ($result -ne "Yes") {
                $chkDeleteOld.Checked = $false
            }
            else {
                $chkMoveOld.Checked = $false
                $chkMoveOld.Enabled = $false
            }
        }
        else {
            $chkMoveOld.Enabled = $true
        }
    })

$lstQueue.Add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    })

    $lstQueue.Add_DragDrop({
        if ($script:processing) { return }
        $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        foreach ($f in $files) {
            if (-not $lstQueue.Items.Contains($f)) {
                $lstQueue.Items.Add($f) | Out-Null
            }
        }
    })

    $form.Add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    })

    $form.Add_DragDrop({
        if ($script:processing) { return }
        $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        foreach ($f in $files) {
            if (-not $lstQueue.Items.Contains($f)) {
                $lstQueue.Items.Add($f) | Out-Null
            }
        }
    })

    $btnQueueAdd.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select a folder to add to queue"
        $folderBrowser.ShowNewFolderButton = $true
        if ($folderBrowser.ShowDialog() -eq "OK") {
            if (-not $lstQueue.Items.Contains($folderBrowser.SelectedPath)) {
                $lstQueue.Items.Add($folderBrowser.SelectedPath) | Out-Null
            }
        }
    })

    $btnQueueClear.Add_Click({
        $lstQueue.Items.Clear()
    })

    $btnCancel.Add_Click({
        if ($btnConvert.Enabled -eq $false -and $btnCancel.Enabled -eq $true) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to cancel processing?",
                "Cancel Confirm", "YesNo", "Warning"
            )
            if ($result -eq "Yes") {
                $script:cancelProcessing = $true
                Set-Status "Canceling... please wait." "Error"
            }
        }
    })

$script:processing = $false
$script:cancelProcessing = $false
$form.Add_FormClosing({
    if ($script:processing) {
        $_.Cancel = $true
        $script:cancelProcessing = $true
        Set-Status "Canceling current file; close again when the batch stops." "Working"
    }
})

$btnConvert.Add_Click({
    if ($script:processing) { return }
    if ($lstQueue.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Add folders or files first. Extract ZIP/7z archives before adding them.", "Queue", "OK", "Information")
        return
    }
    $script:processing = $true
    $script:cancelProcessing = $false
    $txtLog.Clear()
    $paths = @($lstQueue.Items | ForEach-Object { [string]$_ })
    $recursive = $chkRecursive.Checked
    $showExcel = $chkShowExcel.Checked
    $timeout = [int]$numTimeout.Value
    $action = if ($chkDeleteOld.Checked) { 'Delete' } elseif ($chkMoveOld.Checked) { 'Move' } else { 'Keep' }
    $controls = @($btnConvert, $btnQueueAdd, $btnQueueFiles, $btnQueueClear, $btnScanXLStart, $chkRecursive, $chkShowExcel, $chkMoveOld, $chkDeleteOld, $numTimeout)
    foreach ($control in $controls) { $control.Enabled = $false }
    $btnCancel.Enabled = $true
    $results = [Collections.Generic.List[object]]::new()
    $report = ''
    $batchId = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6)
    try {
        $reportDir = Join-Path $env:LOCALAPPDATA 'ExcelConverter\Reports'
        [IO.Directory]::CreateDirectory($reportDir) | Out-Null
        $report = Join-Path $reportDir ($batchId + '.csv')
        'Source,Output,Archive,Status,Error,DurationSeconds' | Set-Content -LiteralPath $report -Encoding UTF8
        Set-Status "Scanning queue..." "Working"
        $queue = Get-ECQueue -Paths $paths -Recurse:$recursive -OnProgress {
            [System.Windows.Forms.Application]::DoEvents()
            return (-not $script:cancelProcessing)
        }
        foreach ($errorItem in $queue.Errors) {
            Write-Log "SCAN ERROR: $($errorItem.Source) - $($errorItem.Error)"
            $results.Add([pscustomobject]@{ Source = $errorItem.Source; Output = ''; Archive = ''; Status = 'ScanError'; Error = $errorItem.Error; DurationSeconds = 0 })
        }
        foreach ($file in $queue.AlreadyXlsx) {
            Write-Log "ALREADY XLSX (no conversion needed): $($file.FullName)"
            $results.Add([pscustomobject]@{ Source = $file.FullName; Output = ''; Archive = ''; Status = 'AlreadyXlsx'; Error = ''; DurationSeconds = 0 })
        }
        Write-Log "Queue: $($queue.Files.Count) to convert; $($queue.AlreadyXlsx.Count) already XLSX; $($queue.Errors.Count) scan error(s)."
        Write-Log "Original action: $action. Timeout: $timeout seconds per file."
        $progressBar.Value = 0
        $progressBar.Maximum = [Math]::Max(1, $queue.Files.Count)
        if ($results.Count -gt 0) { $results | Export-Csv -LiteralPath $report -NoTypeInformation -Encoding UTF8 }
        foreach ($file in $queue.Files) {
            if ($script:cancelProcessing) {
                $results.Add([pscustomobject]@{ Source = $file.FullName; Output = ''; Archive = ''; Status = 'NotProcessed'; Error = 'Batch canceled'; DurationSeconds = 0 })
                continue
            }
            Set-Status "Converting $($progressBar.Value + 1) / $($queue.Files.Count)..." "Working"
            Write-Log "Converting: $($file.FullName)"
            $result = Invoke-ECFile -Source $file.FullName -OriginalAction $action -TimeoutSeconds $timeout -ShowExcel $showExcel -OnProgress {
                [System.Windows.Forms.Application]::DoEvents()
                return (-not $script:cancelProcessing)
            }
            $results.Add($result)
            Write-Log "$($result.Status): $($result.Source) -> $($result.Output) $($result.Error)"
            $progressBar.Value++
            # Persist each completed file so a later failure does not lose the report.
            $result | Select-Object Source, Output, Archive, Status, Error, DurationSeconds |
                Export-Csv -LiteralPath $report -Append -NoTypeInformation -Encoding UTF8
        }
        if ($results.Count -gt 0) { $results | Select-Object Source, Output, Archive, Status, Error, DurationSeconds | Export-Csv -LiteralPath $report -NoTypeInformation -Encoding UTF8 }
        $ok = @($results | Where-Object Status -eq 'Converted').Count
        $warn = @($results | Where-Object { $_.Status -in 'ConvertedWithWarning', 'OutputSaved' }).Count
        $bad = @($results | Where-Object { $_.Status -in 'Failed', 'TimedOut', 'ScanError' }).Count
        $state = if ($script:cancelProcessing) { 'Canceled' } elseif ($bad -gt 0 -or $warn -gt 0) { 'Finished with issues' } else { 'Finished' }
        $statusType = if ($bad -gt 0) { 'Error' } elseif ($warn -gt 0 -or $script:cancelProcessing) { 'Working' } else { 'Success' }
        Set-Status "$state`: $ok OK, $warn warnings, $bad errors" $statusType
        Write-Log "Report: $report"
        [System.Windows.Forms.MessageBox]::Show("$state`n`nConverted: $ok`nWarnings: $warn`nErrors: $bad`nAlready XLSX: $($queue.AlreadyXlsx.Count)`n`nReport: $report", "Batch result", "OK", "Information")
    } catch {
        Write-Log "BATCH ERROR: $($_.Exception.Message)"
        Set-Status "Batch stopped: $($_.Exception.Message)" "Error"
    } finally {
        $script:processing = $false
        foreach ($control in $controls) { $control.Enabled = $true }
        $chkMoveOld.Enabled = -not $chkDeleteOld.Checked
        $btnCancel.Enabled = $false
    }
})

# Show form
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
