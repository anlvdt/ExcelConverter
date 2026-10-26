# Opt-in integration test: run on Windows with desktop Microsoft Excel installed.
# Only generated fixtures are used; no private workbooks or active user workbooks are changed.
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Run this integration test on Windows with desktop Excel' }
Import-Module (Join-Path (Split-Path $PSScriptRoot) 'ExcelConverter.Core.psm1') -Force
$temp = Join-Path ([IO.Path]::GetTempPath()) ('ec-smoke-' + [guid]::NewGuid().ToString('N').Substring(0, 10))
[IO.Directory]::CreateDirectory($temp) | Out-Null
$excel = $null; $book = $null; $sheet = $null; $books = $null; $sheets = $null; $cells = $null
function Assert($Condition, $Message) { if (-not $Condition) { throw $Message } }
try {
    if (-not ('SmokeExcelOwner' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SmokeExcelOwner {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
}
'@
    }
    $preexistingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    [uint32]$fixturePid = 0
    [SmokeExcelOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$fixturePid) | Out-Null
    if ($preexistingIds -contains [int]$fixturePid) {
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
        $excel = $null
        throw 'Smoke test refused to use an existing Excel process'
    }
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3; $excel.EnableEvents = $false
    $books = $excel.Workbooks
    $book = $books.Add()
    $sheets = $book.Worksheets; $sheet = $sheets.Item(1); $cells = $sheet.Cells
    $cell = $cells.Item(1, 1); $cell.Value2 = 'Unicode HR fixture'; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) | Out-Null
    $cell = $cells.Item(2, 1); $cell.Formula = '=SUM(1,2)'; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) | Out-Null
    $source = Join-Path $temp 'fixture.xls'
    $book.SaveAs($source, 56)
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cells) | Out-Null; $cells = $null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null; $sheet = $null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheets) | Out-Null; $sheets = $null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null; $book = $null
    # Copy source using the same native IO as production to a synced-style path over 300 chars.
    $deep = $temp
    1..6 | ForEach-Object { $deep = [IO.Path]::Combine($deep, ('OneDrive - Organization [HR] ' + ('x' * 20))) }
    [ExcelConverter.LongIO]::CreateDirectory($deep)
    $longSource = [IO.Path]::Combine($deep, 'đối chiếu [2024].xls')
    [ExcelConverter.LongIO]::Copy($source, $longSource)
    $q = Get-ECQueue -Paths @($deep) -Recurse
    Assert ($q.Files.Count -eq 1) 'Long synced-style source not discovered'
    $r = Invoke-ECFile -Source $longSource -TimeoutSeconds 120
    Assert ($r.Status -eq 'Converted') "Conversion failed: $($r.Error)"
    Assert (Test-ECXlsx $r.Output) 'Published XLSX validation failed'
    Assert ([ExcelConverter.LongIO]::Exists($longSource)) 'Default conversion removed original'
    # Open a short copy to inspect real Excel values and formulas.
    $check = Join-Path $temp 'check.xlsx'
    [ExcelConverter.LongIO]::Copy($r.Output, $check)
    $book = $books.Open($check, 0, $true)
    $sheets = $book.Worksheets; $sheet = $sheets.Item(1); $cells = $sheet.Cells
    $cell = $cells.Item(1, 1); $value = $cell.Value2; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) | Out-Null
    Assert ($value -eq 'Unicode HR fixture') 'Cell content changed'
    $cell = $cells.Item(2, 1); $formula = $cell.Formula; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) | Out-Null
    Assert ($formula -eq '=SUM(1,2)') 'Formula changed'
    Assert ($excel.Hwnd -ne 0) 'Fixture Excel session was killed by converter'
    Write-Host 'PASS actual Excel conversion, long path, Unicode, values, formulas, original preservation and user-session isolation'
} finally {
    if ($book) { try { $book.Close($false) } catch { } }
    foreach ($com in @($cells, $sheet, $sheets, $book, $books)) {
        if ($com) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) | Out-Null } catch { } }
    }
    if ($excel) { try { $excel.Quit() } catch { }; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [ExcelConverter.LongIO]::DeleteTree($temp)
}
