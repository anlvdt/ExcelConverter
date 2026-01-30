#############################
# XLSStart Virus Cleaner
# Standalone Tool for X97M/Laroux Virus Detection & Removal
#
# Author: Le An (Vietnam IT)
# Version: 1.0.0
# Date: January 2026
#
# FEATURES:
# =========
# - Scan all XLSTART folders (Office 2007-2021, 365)
# - Detect 27+ known virus file variants
# - Detect 23+ malicious VBA module names
# - Deep scan VBA code in Excel files
# - Auto-remove infected files
# - Detailed scan report
# - Can run standalone or be called from other scripts
#
# Virus Database:
# - Microsoft Threat Encyclopedia (X97M/Laroux)
# - TrendMicro Threat Database
#############################

param(
    [switch]$Silent,      # Run without GUI
    [switch]$AutoClean    # Automatically clean without confirmation
)

# Hide console window when running GUI mode
if (-not $Silent) {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Colors - Professional Enterprise Theme
$primaryColor = [System.Drawing.Color]::FromArgb(0, 120, 212)      # Microsoft Blue
$successColor = [System.Drawing.Color]::FromArgb(16, 124, 16)      # Green
$warningColor = [System.Drawing.Color]::FromArgb(255, 140, 0)      # Orange
$errorColor = [System.Drawing.Color]::FromArgb(196, 43, 28)        # Red
$bgColor = [System.Drawing.Color]::FromArgb(243, 243, 243)         # Light gray
$cardColor = [System.Drawing.Color]::White
$textColor = [System.Drawing.Color]::FromArgb(32, 32, 32)          # Dark gray
$subtextColor = [System.Drawing.Color]::FromArgb(96, 96, 96)       # Medium gray

# X97M_LAROUX virus variants - Complete list
$script:virusFileNames = @(
    # Microsoft documented variants
    "BINV.XLS", "BOOK1.XLS", "CAR.XLS", "CURE.XLS", "DIMON.XLS",
    "ECSYSTEM.XLS", "KINSLAYER.XLS", "NEGS.XLS", "NOCAL.XLS",
    "PERSONAL.XLS", "PLDT.XLS", "PRIVAT.XLS", "RESULTS.XLS",
    "SGV.XLS", "SING.XLS", "STARTUP.XLS", "VERA.XLS", "WINDOS.XLS",
    "XLSTART.XLS",
    # Additional variants
    "k4.xls", "xl5glary.xls", "mypersonnel.xls", "Xlscan.xls",
    "laroux.xls", "sheet.xls", "auto.xls", "ssheet.xls"
)

# Known malicious VBA module names
$script:virusModuleNames = @(
    "car", "cure", "foxz", "lalala", "laroux", "locas",
    "monci", "pldt", "program", "results", "sgv", "startup", "wendy",
    "vera", "binv", "dimon", "ecsystem", "kinslayer", "negs",
    "nocal", "privat", "sing", "windos", "xlstart", "k4", "xl5glary"
)

# Known malicious macro/function names
$script:virusMacroNames = @(
    "auto_open", "check_files", "ck_files", "scan_files",
    "cop", "escape", "del", "back",
    "Auto_Open", "Check_Files", "Ck_Files", "Scan_Files"
)

# XLSTART paths to scan
$script:xlstartPaths = @(
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

# Scan results
$script:scanResults = @{
    TotalPaths        = 0
    ScannedPaths      = 0
    TotalFiles        = 0
    InfectedFiles     = 0
    CleanedFiles      = 0
    Errors            = 0
    InfectedFilesList = @()
    CleanedFilesList  = @()
}

# Functions
function Write-ScanLog {
    param([string]$Message, [string]$Type = "Info")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    if ($Silent) {
        Write-Host $logMessage
    }
    else {
        if ($txtLog) {
            $txtLog.AppendText("$logMessage`r`n")
            $txtLog.SelectionStart = $txtLog.Text.Length
            $txtLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
}

function Test-VirusFile {
    param([string]$FilePath)
    
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    
    foreach ($virusName in $script:virusFileNames) {
        if ($fileName -ieq $virusName) {
            return $true
        }
    }
    return $false
}

function Test-VirusInWorkbook {
    param([string]$FilePath)
    
    $infected = $false
    $infectedModules = @()
    
    try {
        $objExcel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $objExcel.Visible = $false
        $objExcel.DisplayAlerts = $false
        
        $doc = $objExcel.Workbooks.Open($FilePath, $null, $true) # Read-only
        
        try {
            $vbProj = $doc.VBProject
            if ($vbProj) {
                foreach ($comp in $vbProj.VBComponents) {
                    $compName = $comp.Name.ToLower()
                    
                    # Check module name
                    foreach ($virusModule in $script:virusModuleNames) {
                        if ($compName -eq $virusModule.ToLower()) {
                            $infected = $true
                            $infectedModules += $comp.Name
                            break
                        }
                    }
                    
                    # Check macro content
                    if ($comp.Type -eq 1) {
                        # vbext_ct_StdModule
                        try {
                            $codeModule = $comp.CodeModule
                            if ($codeModule.CountOfLines -gt 0) {
                                $code = $codeModule.Lines(1, $codeModule.CountOfLines).ToLower()
                                foreach ($virusMacro in $script:virusMacroNames) {
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
            # VBA project may be protected
        }
        
        $doc.Close($false)
        $objExcel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
        
    }
    catch {
        # Could not open file
    }
    
    return @{
        Infected = $infected
        Modules  = $infectedModules
    }
}

function Start-VirusScan {
    param([switch]$CleanMode)
    
    Write-ScanLog "=== XLSStart Virus Scanner Started ==="
    Write-ScanLog "Scan Mode: $(if($CleanMode){'CLEAN & REMOVE'}else{'SCAN ONLY'})"
    Write-ScanLog ""
    
    $script:scanResults.TotalPaths = $script:xlstartPaths.Count
    $script:scanResults.ScannedPaths = 0
    $script:scanResults.TotalFiles = 0
    $script:scanResults.InfectedFiles = 0
    $script:scanResults.CleanedFiles = 0
    $script:scanResults.Errors = 0
    $script:scanResults.InfectedFilesList = @()
    $script:scanResults.CleanedFilesList = @()
    
    # Scan all XLSTART paths
    foreach ($xlstartPath in $script:xlstartPaths) {
        if (Test-Path $xlstartPath) {
            Write-ScanLog "Scanning: $xlstartPath"
            $script:scanResults.ScannedPaths++
            
            $xlsFiles = Get-ChildItem -Path $xlstartPath -Include "*.xls", "*.xlsm", "*.xla", "*.xlam" -File -ErrorAction SilentlyContinue
            
            foreach ($file in $xlsFiles) {
                $script:scanResults.TotalFiles++
                $isVirus = $false
                $virusType = ""
                
                # Check filename
                if (Test-VirusFile -FilePath $file.FullName) {
                    $isVirus = $true
                    $virusType = "Filename Match"
                    Write-ScanLog "  [VIRUS] $($file.Name) - Known virus filename"
                }
                else {
                    # Deep scan VBA
                    Write-ScanLog "  Scanning: $($file.Name)..."
                    $vbaScan = Test-VirusInWorkbook -FilePath $file.FullName
                    if ($vbaScan.Infected) {
                        $isVirus = $true
                        $virusType = "VBA Code: $($vbaScan.Modules -join ', ')"
                        Write-ScanLog "  [VIRUS] $($file.Name) - Infected modules: $($vbaScan.Modules -join ', ')"
                    }
                    else {
                        Write-ScanLog "  [CLEAN] $($file.Name)"
                    }
                }
                
                if ($isVirus) {
                    $script:scanResults.InfectedFiles++
                    $script:scanResults.InfectedFilesList += @{
                        Path = $file.FullName
                        Name = $file.Name
                        Type = $virusType
                    }
                    
                    if ($CleanMode) {
                        try {
                            Remove-Item $file.FullName -Force
                            Write-ScanLog "  [REMOVED] $($file.FullName)"
                            $script:scanResults.CleanedFiles++
                            $script:scanResults.CleanedFilesList += $file.FullName
                        }
                        catch {
                            Write-ScanLog "  [ERROR] Could not remove: $($file.Name) - File may be in use"
                            $script:scanResults.Errors++
                        }
                    }
                }
            }
        }
    }
    
    # Check for Xlscan.386 (VCX variant)
    $xlscan386 = "C:\Windows\System\Xlscan.386"
    if (Test-Path $xlscan386) {
        Write-ScanLog "[VIRUS] Found: $xlscan386 (VCX variant)"
        $script:scanResults.InfectedFiles++
        $script:scanResults.InfectedFilesList += @{
            Path = $xlscan386
            Name = "Xlscan.386"
            Type = "VCX Variant"
        }
        
        if ($CleanMode) {
            try {
                Remove-Item $xlscan386 -Force
                Write-ScanLog "[REMOVED] $xlscan386"
                $script:scanResults.CleanedFiles++
                $script:scanResults.CleanedFilesList += $xlscan386
            }
            catch {
                Write-ScanLog "[ERROR] Could not remove: $xlscan386"
                $script:scanResults.Errors++
            }
        }
    }
    
    Write-ScanLog ""
    Write-ScanLog "=== Scan Complete ==="
    Write-ScanLog "Paths scanned: $($script:scanResults.ScannedPaths)/$($script:scanResults.TotalPaths)"
    Write-ScanLog "Files scanned: $($script:scanResults.TotalFiles)"
    Write-ScanLog "Infected files: $($script:scanResults.InfectedFiles)"
    if ($CleanMode) {
        Write-ScanLog "Files cleaned: $($script:scanResults.CleanedFiles)"
        Write-ScanLog "Errors: $($script:scanResults.Errors)"
    }
    
    return $script:scanResults
}

# GUI Mode
if (-not $Silent) {
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "XLSStart Virus Cleaner"
    $form.Size = New-Object System.Drawing.Size(700, 590)
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
    $lblHeader.Text = "XLSStart Virus Cleaner"
    $lblHeader.Location = New-Object System.Drawing.Point(25, 15)
    $lblHeader.Size = New-Object System.Drawing.Size(400, 30)
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::White
    $lblHeader.BackColor = [System.Drawing.Color]::Transparent
    $pnlHeader.Controls.Add($lblHeader)
    
    $lblSubHeader = New-Object System.Windows.Forms.Label
    $lblSubHeader.Text = "Detect and remove X97M/Laroux virus from Excel XLSTART folders"
    $lblSubHeader.Location = New-Object System.Drawing.Point(25, 50)
    $lblSubHeader.Size = New-Object System.Drawing.Size(500, 20)
    $lblSubHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSubHeader.ForeColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
    $lblSubHeader.BackColor = [System.Drawing.Color]::Transparent
    $pnlHeader.Controls.Add($lblSubHeader)
    
    # Info Panel
    $pnlInfo = New-Object System.Windows.Forms.Panel
    $pnlInfo.Location = New-Object System.Drawing.Point(20, 95)
    $pnlInfo.Size = New-Object System.Drawing.Size(645, 120)
    $pnlInfo.BackColor = $cardColor
    $form.Controls.Add($pnlInfo)
    
    $lblInfoTitle = New-Object System.Windows.Forms.Label
    $lblInfoTitle.Text = "Virus Database"
    $lblInfoTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblInfoTitle.Size = New-Object System.Drawing.Size(200, 20)
    $lblInfoTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblInfoTitle.ForeColor = $textColor
    $pnlInfo.Controls.Add($lblInfoTitle)
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "- 27+ known virus file variants`r`n- 23+ malicious VBA module names`r`n- 8+ virus macro functions"
    $lblInfo.Location = New-Object System.Drawing.Point(15, 35)
    $lblInfo.Size = New-Object System.Drawing.Size(620, 60)
    $lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblInfo.ForeColor = $subtextColor
    $lblInfo.AutoSize = $false
    $pnlInfo.Controls.Add($lblInfo)
    
    $lblPaths = New-Object System.Windows.Forms.Label
    $lblPaths.Text = "Scanning $($script:xlstartPaths.Count) XLSTART locations (Office 2007-365)"
    $lblPaths.Location = New-Object System.Drawing.Point(15, 95)
    $lblPaths.Size = New-Object System.Drawing.Size(620, 18)
    $lblPaths.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblPaths.ForeColor = $subtextColor
    $lblPaths.AutoSize = $false
    $pnlInfo.Controls.Add($lblPaths)
    
    # Log Panel
    $pnlLog = New-Object System.Windows.Forms.Panel
    $pnlLog.Location = New-Object System.Drawing.Point(20, 225)
    $pnlLog.Size = New-Object System.Drawing.Size(645, 250)
    $pnlLog.BackColor = $cardColor
    $form.Controls.Add($pnlLog)
    
    $lblLogTitle = New-Object System.Windows.Forms.Label
    $lblLogTitle.Text = "Scan Log"
    $lblLogTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblLogTitle.Size = New-Object System.Drawing.Size(200, 20)
    $lblLogTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblLogTitle.ForeColor = $textColor
    $pnlLog.Controls.Add($lblLogTitle)
    
    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(15, 38)
    $txtLog.Size = New-Object System.Drawing.Size(615, 200)
    $txtLog.Multiline = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.ReadOnly = $true
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $txtLog.BorderStyle = "FixedSingle"
    $txtLog.ForeColor = $textColor
    $pnlLog.Controls.Add($txtLog)
    
    # Button Panel
    $pnlButtons = New-Object System.Windows.Forms.Panel
    $pnlButtons.Location = New-Object System.Drawing.Point(20, 485)
    $pnlButtons.Size = New-Object System.Drawing.Size(645, 50)
    $pnlButtons.BackColor = $cardColor
    $form.Controls.Add($pnlButtons)
    
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "SCAN ONLY"
    $btnScan.Location = New-Object System.Drawing.Point(15, 10)
    $btnScan.Size = New-Object System.Drawing.Size(200, 32)
    $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnScan.BackColor = $primaryColor
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnScan.FlatStyle = "Flat"
    $btnScan.FlatAppearance.BorderSize = 0
    $btnScan.Cursor = "Hand"
    $pnlButtons.Controls.Add($btnScan)
    
    $btnClean = New-Object System.Windows.Forms.Button
    $btnClean.Text = "SCAN & CLEAN"
    $btnClean.Location = New-Object System.Drawing.Point(225, 10)
    $btnClean.Size = New-Object System.Drawing.Size(200, 32)
    $btnClean.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClean.BackColor = $warningColor
    $btnClean.ForeColor = [System.Drawing.Color]::White
    $btnClean.FlatStyle = "Flat"
    $btnClean.FlatAppearance.BorderSize = 0
    $btnClean.Cursor = "Hand"
    $pnlButtons.Controls.Add($btnClean)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "CLOSE"
    $btnClose.Location = New-Object System.Drawing.Point(435, 10)
    $btnClose.Size = New-Object System.Drawing.Size(195, 32)
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClose.BackColor = $subtextColor
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = "Flat"
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = "Hand"
    $pnlButtons.Controls.Add($btnClose)
    
    # Event Handlers
    $btnScan.Add_Click({
            $txtLog.Clear()
            $btnScan.Enabled = $false
            $btnClean.Enabled = $false
        
            $results = Start-VirusScan
        
            if ($results.InfectedFiles -gt 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Scan completed!`n`nInfected files found: $($results.InfectedFiles)`n`nUse 'SCAN & CLEAN' to remove them.",
                    "Virus Detected", "OK", "Warning"
                )
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Scan completed!`n`nNo viruses detected. Your system is clean.",
                    "Clean", "OK", "Information"
                )
            }
        
            $btnScan.Enabled = $true
            $btnClean.Enabled = $true
        })
    
    $btnClean.Add_Click({
            $result = [System.Windows.Forms.MessageBox]::Show(
                "This will PERMANENTLY DELETE all infected files!`n`nAre you sure?",
                "Confirm Clean", "YesNo", "Warning"
            )
        
            if ($result -eq "Yes") {
                $txtLog.Clear()
                $btnScan.Enabled = $false
                $btnClean.Enabled = $false
            
                $results = Start-VirusScan -CleanMode
            
                if ($results.CleanedFiles -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Cleaning completed!`n`nFiles removed: $($results.CleanedFiles)`nErrors: $($results.Errors)",
                        "Clean Complete", "OK", "Information"
                    )
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "No infected files found to clean.",
                        "Clean", "OK", "Information"
                    )
                }
            
                $btnScan.Enabled = $true
                $btnClean.Enabled = $true
            }
        })
    
    $btnClose.Add_Click({
            $form.Close()
        })
    
    # Show form
    $form.Add_Shown({ 
            $form.Activate()
            Write-ScanLog "XLSStart Virus Cleaner ready"
            Write-ScanLog "Click 'SCAN ONLY' to detect viruses or 'SCAN & CLEAN' to remove them"
        })
    [System.Windows.Forms.Application]::Run($form)
    
}
else {
    # Silent mode - Command line
    if ($AutoClean) {
        $results = Start-VirusScan -CleanMode
    }
    else {
        $results = Start-VirusScan
    }
    
    exit $results.InfectedFiles
}
