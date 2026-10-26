param([Parameter(Mandatory = $true)][string]$JobDirectory)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelConverter.Core.psm1') -Force
$request = Get-Content -LiteralPath (Join-Path $JobDirectory 'request.json') -Raw | ConvertFrom-Json
$resultPath = Join-Path $JobDirectory 'result.json'
$excel = $null; $books = $null; $book = $null
$result = $null
function Write-ECState($Path, $Value) {
    $pending = $Path + '.pending'
    [IO.File]::WriteAllText($pending, ($Value | ConvertTo-Json -Depth 5), [Text.Encoding]::UTF8)
    [IO.File]::Move($pending, $Path)
}
try {
    if ($env:OS -ne 'Windows_NT') { throw 'Microsoft Excel COM conversion requires Windows and desktop Microsoft Excel' }
    $source = [string]$request.Source
    $extension = [IO.Path]::GetExtension([ExcelConverter.LongIO]::Name($source))
    $inputPath = Join-Path $JobDirectory ('input' + $extension)
    $outputPath = Join-Path $JobDirectory 'output.xlsx'
    # Reading a synced placeholder requests hydration from the installed sync client.
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            [ExcelConverter.LongIO]::Copy($source, $inputPath)
            # Clear ReadOnly on the private copy so cleanup can remove it later.
            $attributes = [IO.File]::GetAttributes($inputPath)
            [IO.File]::SetAttributes($inputPath, [IO.FileAttributes]([int]$attributes -band (-bnot [int][IO.FileAttributes]::ReadOnly)))
            $sourceHash = [ExcelConverter.LongIO]::Hash($inputPath)
            if ([ExcelConverter.LongIO]::Hash($source) -ne $sourceHash) { throw 'Source changed while being copied' }
            break
        } catch {
            if ([IO.File]::Exists($inputPath)) { [IO.File]::Delete($inputPath) }
            if ($attempt -eq 3) { throw }
            Start-Sleep -Milliseconds (500 * $attempt)
        }
    }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelOwner {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
}
'@
    $existing = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelId = 0
    [ExcelOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelId) | Out-Null
    $owner = Get-Process -Id $excelId -ErrorAction Stop
    if ($existing -contains [int]$excelId) {
        # Never Quit or terminate an Excel process that existed before this job.
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
        $excel = $null
        throw 'Excel returned an existing user process; conversion stopped to preserve that session'
    }
    Write-ECState (Join-Path $JobDirectory 'owner.json') @{ Id = $excelId; StartTicks = $owner.StartTime.ToUniversalTime().Ticks }
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.DisplayAlerts = $false
    $excel.Visible = [bool]$request.ShowExcel
    $books = $excel.Workbooks
    # Explicit empty passwords avoid unattended password dialogs; never update links.
    $missing = [Type]::Missing
    $book = $books.Open($inputPath, 0, $true, $missing, '', '', $true, $missing, $missing, $false, $false, $missing, $false)
    # Preserve cached formula results and avoid an automatic recalc during SaveAs.
    $calculationWarning = ''
    try { $excel.Calculation = -4135; $excel.CalculateBeforeSave = $false }
    catch { $calculationWarning = 'Excel could not disable recalculation; review cached formula results.' }
    $externalLinks = @($book.LinkSources(1)) | Where-Object { $null -ne $_ -and $_ -ne '' }
    $book.CheckCompatibility = $false
    $book.SaveAs($outputPath, 51)
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
    $book = $null
    # Confirm that Excel itself can reopen the saved result before publishing it.
    $book = $books.Open($outputPath, 0, $true, $missing, '', '', $true, $missing, $missing, $false, $false, $missing, $false)
    if ($book.FileFormat -ne 51) { throw 'Excel did not produce XLSX format 51' }
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
    $book = $null
    $effectiveAction = if (@($externalLinks).Count -gt 0 -or $calculationWarning) { 'Keep' } else { $request.OriginalAction }
    $result = Publish-ECWorkbook -Source $source -StagedOutput $outputPath -SourceHash $sourceHash -OriginalAction $effectiveAction -OnPublished {
        param($checkpoint)
        Write-ECState (Join-Path $JobDirectory 'published.json') $checkpoint
    }
    if ($calculationWarning) {
        $result.Status = 'ConvertedWithWarning'
        $result.Error = ($result.Error + ' ' + $calculationWarning).Trim()
    }
    if (@($externalLinks).Count -gt 0) {
        $result.Status = 'ConvertedWithWarning'
        $result.Error = ($result.Error + ' Original retained: external workbook links were not refreshed; review relative link targets after local staging.').Trim()
    }
} catch {
    $result = [pscustomobject]@{ Source = $request.Source; Output = ''; Archive = ''; Status = 'Failed'; Error = $_.Exception.Message }
} finally {
    # Save the result before COM cleanup, which may itself hang.
    if ($result) {
        try { Write-ECState $resultPath $result }
        catch { [Console]::Error.WriteLine('Could not save worker result: ' + $_.Exception.Message) }
    }
    if ($book) {
        try { $book.Close($false) } catch { }
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null } catch { }
    }
    if ($books) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($books) | Out-Null } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
