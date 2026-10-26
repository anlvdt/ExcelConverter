$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot) 'ExcelConverter.Core.psm1') -Force
$passed = 0
function Assert($Condition, $Message) { if (-not $Condition) { throw $Message } }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('ec-worker-tests-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp) | Out-Null
try {
    $fake = Join-Path $temp 'fake.ps1'
    @'
param($JobDirectory)
$request = Get-Content -LiteralPath (Join-Path $JobDirectory 'request.json') -Raw | ConvertFrom-Json
if ($request.Source -eq 'hang') { Start-Sleep -Seconds 30; exit }
if ($request.Source -eq 'crash') { [Console]::Error.WriteLine('simulated worker failure'); exit 2 }
if ($request.Source -eq 'checkpoint') {
    @{ Source = $request.Source; Output = 'published.xlsx'; Archive = ''; Status = 'OutputSaved'; Error = 'Original cleanup not confirmed' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $JobDirectory 'published.json') -Encoding UTF8
    Start-Sleep -Seconds 30; exit
}
if ($request.Source -eq 'cleanup-hang') {
    @{ Source = $request.Source; Output = 'published.xlsx'; Archive = ''; Status = 'Converted'; Error = '' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $JobDirectory 'result.json') -Encoding UTF8
    Start-Sleep -Seconds 30; exit
}
[IO.File]::WriteAllText((Join-Path $JobDirectory 'readonly-input.xls'), 'private fixture')
[IO.File]::SetAttributes((Join-Path $JobDirectory 'readonly-input.xls'), [IO.FileAttributes]::ReadOnly)
@{ Source = $request.Source; Output = 'result.xlsx'; Archive = ''; Status = 'Converted'; Error = '' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $JobDirectory 'result.json') -Encoding UTF8
'@ | Set-Content -LiteralPath $fake -Encoding UTF8
    $previousTemp = $env:EXCELCONVERTER_TEMP
    $env:EXCELCONVERTER_TEMP = $temp
    $r = Invoke-ECFile -Source 'success' -WorkerPath $fake -TimeoutSeconds 10
    Assert ($r.Status -eq 'Converted') "Subprocess result missing: $($r.Error)"
    Assert (@(Get-ChildItem -LiteralPath $temp -Directory).Count -eq 0) 'Private staging left behind for read-only input'
    $passed++; Write-Host 'PASS worker result and read-only staging cleanup'
    $r = Invoke-ECFile -Source 'crash' -WorkerPath $fake -TimeoutSeconds 10
    Assert ($r.Status -eq 'Failed' -and $r.Error -match 'simulated worker failure') 'Crash diagnostic lost'
    $passed++; Write-Host 'PASS worker crash is isolated'
    $r = Invoke-ECFile -Source 'hang' -WorkerPath $fake -TimeoutSeconds 2
    Assert ($r.Status -eq 'TimedOut' -and $r.DurationSeconds -lt 12) 'Timeout did not bound worker duration'
    $passed++; Write-Host 'PASS hung worker times out'
    $r = Invoke-ECFile -Source 'hang' -WorkerPath $fake -TimeoutSeconds 10 -OnProgress { return $false }
    Assert ($r.Status -eq 'Canceled' -and $r.DurationSeconds -lt 5) 'Cancellation did not stop worker'
    $passed++; Write-Host 'PASS cancellation stops worker'
    $r = Invoke-ECFile -Source 'checkpoint' -WorkerPath $fake -TimeoutSeconds 3
    Assert ($r.Status -eq 'OutputSaved' -and $r.Output -eq 'published.xlsx') 'Published output checkpoint lost'
    $passed++; Write-Host 'PASS timeout preserves publication checkpoint'
    $r = Invoke-ECFile -Source 'cleanup-hang' -WorkerPath $fake -TimeoutSeconds 3
    Assert ($r.Status -eq 'ConvertedWithWarning') 'Completed conversion mislabeled as failed'
    $passed++; Write-Host 'PASS completed output survives hung COM cleanup'
    Write-Host "$passed tests passed"
} finally { $env:EXCELCONVERTER_TEMP = $previousTemp; [IO.Directory]::Delete($temp, $true) }
