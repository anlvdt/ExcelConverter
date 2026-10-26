# Headless entry point. Accepts local folders/files, including synced OneDrive/SharePoint paths.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string[]]$Path,
    [switch]$Recurse,
    [ValidateSet('Keep', 'Move', 'Delete')][string]$OriginalAction = 'Keep',
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 180,
    [string]$ReportPath = (Join-Path ([IO.Path]::GetTempPath()) ('ExcelConverter-' + [guid]::NewGuid().ToString('N') + '.csv'))
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelConverter.Core.psm1') -Force
$queue = Get-ECQueue -Paths $Path -Recurse:$Recurse
$results = [Collections.Generic.List[object]]::new()
foreach ($item in $queue.Errors) {
    $results.Add([pscustomobject]@{ Source = $item.Source; Output = ''; Archive = ''; Status = 'ScanError'; Error = $item.Error; DurationSeconds = 0 })
}
foreach ($item in $queue.AlreadyXlsx) {
    $results.Add([pscustomobject]@{ Source = $item.FullName; Output = ''; Archive = ''; Status = 'AlreadyXlsx'; Error = ''; DurationSeconds = 0 })
}
# Fail before converting anything if the report location is not writable.
'Source,Output,Archive,Status,Error,DurationSeconds' | Set-Content -LiteralPath $ReportPath -Encoding UTF8
foreach ($item in $results) { $item | Export-Csv -LiteralPath $ReportPath -Append -NoTypeInformation -Encoding UTF8 }
foreach ($file in $queue.Files) {
    Write-Host "Converting: $($file.FullName)"
    $result = Invoke-ECFile -Source $file.FullName -OriginalAction $OriginalAction -TimeoutSeconds $TimeoutSeconds
    $results.Add($result)
    $result | Select-Object Source, Output, Archive, Status, Error, DurationSeconds | Export-Csv -LiteralPath $ReportPath -Append -NoTypeInformation -Encoding UTF8
    Write-Host "$($result.Status): $($result.Error)"
}
$results.ToArray() | Write-Output
Write-Host "Report: $ReportPath"
if (@($results | Where-Object { $_.Status -notin 'Converted', 'AlreadyXlsx' }).Count -gt 0) { exit 1 }
