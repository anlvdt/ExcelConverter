#############################
# Excel XLS/XLSM to XLSX Converter
# Professional GUI Application with Virus Protection
#
# Author: Le An (Vietnam IT)
# Version: 1.1.0
# Date: January 2026
#
# FEATURES:
# =========
# [Conversion]
# - Convert XLS (Excel 97-2003) to XLSX
# - Convert XLSM (Macro-enabled) to XLSX
# - Batch convert entire folders with subfolders
# - Move or delete original files after conversion
# - Activity log with timestamps
#
# [Security - X97M/Laroux Virus Protection]
# - Deep scan VBA modules inside workbooks
# - Detect malicious macros by reading code content
# - Auto-remove infected VBA modules
# - Scan all XLSTART folders (Office 2007-2021, 365)
# - Detect 27+ known virus file variants
# - Detect 23+ malicious VBA module names
# - Detect 8+ virus macro functions
# - Final security scan after conversion
#
# Original conversion script reference:
# https://gist.github.com/gabceb/954418
#
# Virus detection based on:
# - Microsoft Threat Encyclopedia (X97M/Laroux)
# - TrendMicro Threat Database
#############################

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
$lblSubHeader.Text = "Convert XLS/XLSM to XLSX with X97M/Laroux virus protection"
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
$btnQueueClear.Text = "Clear List"
$btnQueueClear.Location = New-Object System.Drawing.Point(535, 78)
$btnQueueClear.Size = New-Object System.Drawing.Size(95, 32)
$btnQueueClear.FlatStyle = "Flat"
$btnQueueClear.BackColor = [System.Drawing.Color]::FromArgb(230,230,230)
$btnQueueClear.ForeColor = $textColor
$btnQueueClear.FlatAppearance.BorderSize = 0
$btnQueueClear.Cursor = "Hand"
$pnlQueue.Controls.Add($btnQueueClear)

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
$chkMoveOld.Checked = $true
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
$chkShowExcel.Size = New-Object System.Drawing.Size(260, 22)
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
$lblFooter.Text = "v1.1.0 | Le An (Vietnam IT) | Ref: gist.github.com/gabceb/954418"
$lblFooter.Location = New-Object System.Drawing.Point(20, 588)
$lblFooter.Size = New-Object System.Drawing.Size(645, 18)
$lblFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblFooter.ForeColor = $subtextColor
$lblFooter.TextAlign = "MiddleCenter"
$form.Controls.Add($lblFooter)

# X97M_LAROUX virus variants - Complete list from Microsoft Threat Encyclopedia
# Known infected file names created in XLSTART folder
$virusFileNames = @(
    # Microsoft documented variants
    "BINV.XLS", "BOOK1.XLS", "CAR.XLS", "CURE.XLS", "DIMON.XLS",
    "ECSYSTEM.XLS", "KINSLAYER.XLS", "NEGS.XLS", "NOCAL.XLS",
    "PERSONAL.XLS", "PLDT.XLS", "PRIVAT.XLS", "RESULTS.XLS",
    "SGV.XLS", "SING.XLS", "STARTUP.XLS", "VERA.XLS", "WINDOS.XLS",
    "XLSTART.XLS",
    # Additional variants from other sources
    "k4.xls", "xl5glary.xls", "mypersonnel.xls", "Xlscan.xls",
    "laroux.xls", "sheet.xls", "auto.xls", "ssheet.xls"
)

# Known malicious VBA module names from Microsoft
$virusModuleNames = @(
    "car", "cure", "foxz", "lalala", "laroux", "locas",
    "monci", "pldt", "program", "results", "sgv", "startup", "wendy",
    # Additional module names
    "vera", "binv", "dimon", "ecsystem", "kinslayer", "negs",
    "nocal", "privat", "sing", "windos", "xlstart", "k4", "xl5glary"
)

# Known malicious macro/function names from Microsoft
$virusMacroNames = @(
    "auto_open", "check_files", "ck_files", "scan_files",
    "cop", "escape", "del", "back"
)

# XLSTART paths to scan (all possible locations)
$xlstartPaths = @(
    "$env:APPDATA\Microsoft\Excel\XLSTART",
    "$env:USERPROFILE\AppData\Roaming\Microsoft\Excel\XLSTART",
    # Office 365 / Office 2019/2021
    "C:\Program Files\Microsoft Office\root\Office16\XLSTART",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\XLSTART",
    # Office 2016
    "C:\Program Files\Microsoft Office\Office16\XLSTART",
    "C:\Program Files (x86)\Microsoft Office\Office16\XLSTART",
    # Office 2013
    "C:\Program Files\Microsoft Office\Office15\XLSTART",
    "C:\Program Files (x86)\Microsoft Office\Office15\XLSTART",
    # Office 2010
    "C:\Program Files\Microsoft Office\Office14\XLSTART",
    "C:\Program Files (x86)\Microsoft Office\Office14\XLSTART",
    # Office 2007
    "C:\Program Files\Microsoft Office\Office12\XLSTART",
    "C:\Program Files (x86)\Microsoft Office\Office12\XLSTART",
    # Legacy paths
    "C:\MSOFFICE\EXCEL\XLSTART"
)

# Functions
function Write-Log {
    param([string]$Message, [string]$Type = "Info")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-X97MLaroux {
    param([string]$FilePath)
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    # Check against all known virus file names (case-insensitive)
    foreach ($virusName in $virusFileNames) {
        if ($fileName -ieq $virusName) {
            return $true
        }
    }
    return $false
}

function Remove-X97MLarouxVirus {
    $removed = $false
    $removedCount = 0
    
    # Scan all XLSTART paths
    foreach ($xlstartPath in $xlstartPaths) {
        if (Test-Path $xlstartPath) {
            # Get all xls/xlsm files in XLSTART
            $xlsFiles = Get-ChildItem -Path $xlstartPath -Include "*.xls", "*.xlsm", "*.xla", "*.xlam" -File -ErrorAction SilentlyContinue
            foreach ($file in $xlsFiles) {
                if ($script:cancelProcessing -eq $true) {
                    Write-Log "PROCESS CANCELED BY USER!"
                    break
                }
                $isVirus = $false
                
                # Check if filename matches known virus names (case-insensitive)
                foreach ($virusName in $virusFileNames) {
                    if ($file.Name -ieq $virusName) {
                        $isVirus = $true
                        break
                    }
                }
                
                if ($isVirus) {
                    try {
                        Remove-Item $file.FullName -Force
                        Write-Log "[VIRUS] Removed: $($file.FullName)"
                        $removedCount++
                        $removed = $true
                    }
                    catch {
                        Write-Log "[VIRUS] Failed to remove: $($file.FullName) - File may be in use"
                    }
                }
            }
        }
    }
    
    # Check Windows\System folder for Xlscan.386 (VCX variant)
    $xlscan386 = "C:\Windows\System\Xlscan.386"
    if (Test-Path $xlscan386) {
        try {
            Remove-Item $xlscan386 -Force
            Write-Log "[VIRUS] Removed: $xlscan386"
            $removedCount++
            $removed = $true
        }
        catch {}
    }
    
    if ($removedCount -gt 0) {
        Write-Log "[VIRUS] Total files removed: $removedCount"
    }
    
    return $removed
}

function Test-WorkbookForVirus {
    param($Workbook)
    $infected = $false
    $infectedModules = @()
    
    try {
        $vbProj = $Workbook.VBProject
        if ($vbProj) {
            foreach ($comp in $vbProj.VBComponents) {
                $compName = $comp.Name.ToLower()
                
                # Check module name against known virus module names
                foreach ($virusModule in $virusModuleNames) {
                    if ($compName -eq $virusModule.ToLower()) {
                        $infected = $true
                        $infectedModules += $comp.Name
                        break
                    }
                }
                
                # Check for virus macros inside the module
                if ($comp.Type -eq 1) {
                    # vbext_ct_StdModule
                    try {
                        $codeModule = $comp.CodeModule
                        if ($codeModule.CountOfLines -gt 0) {
                            $code = $codeModule.Lines(1, $codeModule.CountOfLines).ToLower()
                            foreach ($virusMacro in $virusMacroNames) {
                                if ($code -match $virusMacro.ToLower()) {
                                    $infected = $true
                                    if ($infectedModules -notcontains $comp.Name) {
                                        $infectedModules += $comp.Name
                                    }
                                    break
                                }
                            }
                        }
                    }
                    catch {}
                }
            }
        }
    }
    catch {
        # VBA project may be protected or inaccessible
    }
    
    return @{
        Infected = $infected
        Modules  = $infectedModules
    }
}

function Remove-VirusFromWorkbook {
    param($Workbook)
    $removed = 0
    
    try {
        $vbProj = $Workbook.VBProject
        if ($vbProj) {
            $componentsToRemove = @()
            
            foreach ($comp in $vbProj.VBComponents) {
                $compName = $comp.Name.ToLower()
                $shouldRemove = $false
                
                # Check module name
                foreach ($virusModule in $virusModuleNames) {
                    if ($compName -eq $virusModule.ToLower()) {
                        $shouldRemove = $true
                        break
                    }
                }
                
                # Check macro content
                if (-not $shouldRemove -and $comp.Type -eq 1) {
                    try {
                        $codeModule = $comp.CodeModule
                        if ($codeModule.CountOfLines -gt 0) {
                            $code = $codeModule.Lines(1, $codeModule.CountOfLines).ToLower()
                            foreach ($virusMacro in $virusMacroNames) {
                                if ($code -match $virusMacro.ToLower()) {
                                    $shouldRemove = $true
                                    break
                                }
                            }
                        }
                    }
                    catch {}
                }
                
                if ($shouldRemove) {
                    $componentsToRemove += $comp
                }
            }
            
            # Remove infected components
            foreach ($comp in $componentsToRemove) {
                try {
                    $compName = $comp.Name
                    $vbProj.VBComponents.Remove($comp)
                    Write-Log "  [VIRUS] Removed module: $compName"
                    $removed++
                }
                catch {
                    Write-Log "  [VIRUS] Could not remove module: $($comp.Name)"
                }
            }
        }
    }
    catch {
        # VBA project protected
    }
    
    return $removed
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
$btnScanXLStart.Add_Click({
        # Launch XLSStart Cleaner tool
        $cleanerPath = Join-Path $PSScriptRoot "XLSStartCleaner.ps1"
        if (Test-Path $cleanerPath) {
            try {
                Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$cleanerPath`"" -WindowStyle Hidden
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not launch XLSStart Cleaner.`n`nError: $($_.Exception.Message)",
                    "Error", "OK", "Error"
                )
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "XLSStartCleaner.ps1 not found!`n`nExpected location: $cleanerPath",
                "Error", "OK", "Error"
            )
        }
    })

$btnAbout.Add_Click({
        $aboutText = "EXCEL CONVERTER v1.1.0`r`n"
        $aboutText += "Author: Le An (Vietnam IT)`r`n"
        $aboutText += "`r`n"
        $aboutText += "FEATURES:`r`n"
        $aboutText += "-----------------------------------------`r`n"
        $aboutText += "[Conversion]`r`n"
        $aboutText += "  - Convert XLS (Excel 97-2003) to XLSX`r`n"
        $aboutText += "  - Convert XLSM (Macro-enabled) to XLSX`r`n"
        $aboutText += "  - Batch convert folders with subfolders`r`n"
        $aboutText += "  - Move or delete original files`r`n"
        $aboutText += "`r`n"
        $aboutText += "[Security - X97M/Laroux Protection]`r`n"
        $aboutText += "  - Deep scan VBA modules in workbooks`r`n"
        $aboutText += "  - Detect malicious macros by code content`r`n"
        $aboutText += "  - Auto-remove infected VBA modules`r`n"
        $aboutText += "  - Scan XLSTART folders (Office 2007-365)`r`n"
        $aboutText += "  - Detect 27+ virus file variants`r`n"
        $aboutText += "  - Detect 23+ malicious module names`r`n"
        $aboutText += "  - Detect 8+ virus macro functions`r`n"
        $aboutText += "`r`n"
        $aboutText += "Reference: gist.github.com/gabceb/954418`r`n"
        $aboutText += "Virus DB: Microsoft & TrendMicro`r`n"
        $aboutText += "`r`n"
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

$btnConvert.Add_Click({
        if ($lstQueue.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please add folders or files to the queue!", "Error", "OK", "Error")
            return
        }
    
        $script:cancelProcessing = $false
        $txtLog.Clear()
        
        $btnConvert.Enabled = $false
        $btnQueueAdd.Enabled = $false
        $btnQueueClear.Enabled = $false
        $btnCancel.Enabled = $true
        
        Set-Status "Scanning queue..." "Working"
        Write-Log "Scanning items in queue..."
    
        # Check for X97M_LAROUX virus before starting
        Write-Log "Scanning for X97M/Laroux virus variants..."
        Write-Log "Checking XLSTART folders for infected files..."
        $virusFound = Remove-X97MLarouxVirus
        if ($virusFound) {
            Write-Log "[SECURITY] X97M/Laroux virus files removed!"
            [System.Windows.Forms.MessageBox]::Show(
                "X97M/Laroux virus detected and removed!`n`nInfected files found in Excel XLSTART folder(s).`n`nKnown variants: PERSONAL.XLS, VERA.XLS, PLDT.XLS, STARTUP.XLS, BOOK1.XLS, etc.",
                "Virus Removed", "OK", "Warning"
            )
        }
        else {
            Write-Log "No virus files found in XLSTART folders"
        }
    
        $xlsFiles = @()
        try {
            foreach ($item in $lstQueue.Items) {
                if (-not (Test-Path $item)) { continue }
                $itemObj = Get-Item $item
                if ($itemObj.PSIsContainer) {
                    if ($chkRecursive.Checked) {
                        $xlsFiles += @(Get-ChildItem -Path $item -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq ".xls" -or $_.Extension -eq ".xlsm" })
                    } else {
                        $xlsFiles += @(Get-ChildItem -Path $item -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -eq ".xls" -or $_.Extension -eq ".xlsm" })
                    }
                } else {
                    if ($itemObj.Extension -eq ".xls" -or $itemObj.Extension -eq ".xlsm") {
                        $xlsFiles += $itemObj
                    }
                }
            }
            # Deduplicate
            $xlsFiles = $xlsFiles | Select-Object -Unique -Property FullName
        }
        catch {
            Write-Log "Error scanning files: $($_.Exception.Message)"
            Set-Status "Error" "Error"
            $btnConvert.Enabled = $true
            $btnQueueAdd.Enabled = $true
            $btnQueueClear.Enabled = $true
            $btnCancel.Enabled = $false
            return
        }
    
        if ($xlsFiles.Count -eq 0) {
            Write-Log "No XLS/XLSM files found in queue!"
            Set-Status "No files found" "Error"
            [System.Windows.Forms.MessageBox]::Show("No .xls or .xlsm files found!", "Info", "OK", "Information")
            $btnConvert.Enabled = $true
            $btnQueueAdd.Enabled = $true
            $btnQueueClear.Enabled = $true
            $btnCancel.Enabled = $false
            return
        }
    
        Write-Log "Found $($xlsFiles.Count) file(s) across batch"
    
        $progressBar.Value = 0
        $progressBar.Maximum = $xlsFiles.Count
    
        $objExcel = $null
    
        try {
            Set-Status "Starting Excel..." "Working"
            Write-Log "Initializing Excel..."
            $objExcel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $objExcel.Visible = $chkShowExcel.Checked
            $objExcel.DisplayAlerts = $false
            $objExcel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
            Write-Log "Excel ready"
        
            $xlOpenXMLWorkbook = 51
            $converted = 0
            $failed = 0
        
            foreach ($file in $xlsFiles) {
                if ($script:cancelProcessing -eq $true) {
                    Write-Log "PROCESS CANCELED BY USER!"
                    break
                }
                $current = $progressBar.Value + 1
                Set-Status "Converting $current of $($xlsFiles.Count)..." "Working"
            
                try {
                    Write-Log "Converting: $($file.Name)"
                
                    # Check if file name matches known virus file names
                    if (Test-X97MLaroux -FilePath $file.FullName) {
                        Write-Log "  [VIRUS] Detected X97M/Laroux infected file!"
                        try {
                            Remove-Item $file.FullName -Force
                            Write-Log "  [VIRUS] Deleted infected file"
                        }
                        catch {
                            Write-Log "  [VIRUS] Could not delete - file in use"
                        }
                        $failed++
                        $progressBar.Value++
                        [System.Windows.Forms.Application]::DoEvents()
                        continue
                    }
                
                    $doc = $objExcel.WorkBooks.Open($file.FullName)
                
                    # Deep scan workbook for virus modules and macros
                    $virusScan = Test-WorkbookForVirus -Workbook $doc
                    if ($virusScan.Infected) {
                        Write-Log "  [VIRUS] Found infected modules: $($virusScan.Modules -join ', ')"
                        $removedMacros = Remove-VirusFromWorkbook -Workbook $doc
                        if ($removedMacros -gt 0) {
                            Write-Log "  [VIRUS] Cleaned $removedMacros malicious module(s)"
                        }
                    }
                
                    $newPath = Join-Path $file.Directory "$($file.BaseName).xlsx"
                
                    if (Test-Path $newPath) {
                        $newPath = Join-Path $file.Directory "$($file.BaseName)_converted.xlsx"
                    }
                
                    $doc.SaveAs($newPath, $xlOpenXMLWorkbook)
                    $doc.Close($false)
                
                    # Handle original file
                    if ($chkDeleteOld.Checked) {
                        Remove-Item $file.FullName -Force
                        Write-Log "  Deleted original"
                    }
                    elseif ($chkMoveOld.Checked) {
                        $oldFolder = Join-Path $file.Directory "old"
                        if (-not (Test-Path $oldFolder)) {
                            New-Item $oldFolder -ItemType Directory -Force | Out-Null
                        }
                        Move-Item $file.FullName $oldFolder -Force
                        Write-Log "  Moved to old folder"
                    }
                
                    Write-Log "  OK -> $($file.BaseName).xlsx"
                    $converted++
                }
                catch {
                    Write-Log "  ERROR: $($_.Exception.Message)"
                    $failed++
                    try { $doc.Close($false) } catch {}
                }
            
                $progressBar.Value++
                [System.Windows.Forms.Application]::DoEvents()
            }
        
            Write-Log "=========================================="
            Write-Log "Completed! Success: $converted, Failed: $failed"
        
            # Final virus check after conversion
            Write-Log "Final security scan..."
            $finalVirusCheck = Remove-X97MLarouxVirus
            if ($finalVirusCheck) {
                Write-Log "[SECURITY] Additional virus files removed after conversion"
            }
        
            Set-Status "Done: $converted OK, $failed failed" "Success"
        
            [System.Windows.Forms.MessageBox]::Show(
                "Conversion completed!`n`nSuccess: $converted`nFailed: $failed", 
                "Result", "OK", "Information"
            )
        }
        catch {
            Write-Log "CRITICAL ERROR: $($_.Exception.Message)"
            Set-Status "Error occurred" "Error"
            [System.Windows.Forms.MessageBox]::Show(
                "Error: $($_.Exception.Message)`n`nMake sure Microsoft Excel is installed.", 
                "Error", "OK", "Error"
            )
        }
        finally {
            $btnConvert.Enabled = $true
            $btnQueueAdd.Enabled = $true
            $btnQueueClear.Enabled = $true
            $btnCancel.Enabled = $false
            
            if ($objExcel) {
                try {
                    $objExcel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
                }
                catch {}
                $objExcel = $null
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    })

# Show form
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
