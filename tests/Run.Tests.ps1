$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$errorsFound = @()
Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' -and $_.FullName -notmatch '[\\/](\.git|dist|\.kiro)[\\/]' } | ForEach-Object {
    $tokens = $null; $parseErrors = $null
    [Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$tokens, [ref]$parseErrors) > $null
    $errorsFound += $parseErrors
}
if ($errorsFound) { $errorsFound | Format-List; throw 'PowerShell parse errors' }
Write-Host 'PASS all PowerShell scripts parse'
& (Join-Path $PSScriptRoot 'Core.Tests.ps1')
& (Join-Path $PSScriptRoot 'Worker.Tests.ps1')
