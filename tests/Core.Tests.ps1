$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
Import-Module (Join-Path $root 'ExcelConverter.Core.psm1') -Force
$script:passed = 0
function Assert($Condition, $Message) { if (-not $Condition) { throw $Message } }
function Test($Name, [scriptblock]$Body) { & $Body; $script:passed++; Write-Host "PASS $Name" }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('ec-tests-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp) | Out-Null
function New-Xlsx($Path, [switch]$Macro, [switch]$InvalidMain, [switch]$HiddenMacroType) {
    $zip = [IO.Compression.ZipFile]::Open($Path, 'Create')
    try {
        $entries = @{
            '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'
            '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            'xl/workbook.xml' = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></sheets></workbook>'
            'xl/_rels/workbook.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
            'xl/worksheets/sheet1.xml' = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
        }
        if ($Macro) { $entries['xl/vbaProject.bin'] = 'macro' }
        if ($InvalidMain) { $entries['xl/workbook.xml'] = '<notaworkbook/>' }
        if ($HiddenMacroType) { $entries['[Content_Types].xml'] = $entries['[Content_Types].xml'].Replace('</Types>', '<Override PartName="/xl/innocent.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'); $entries['xl/innocent.bin'] = 'macro' }
        foreach ($name in $entries.Keys) {
            $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open())
            try { $writer.Write($entries[$name]) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }
}
try {
    Test 'queue retains metadata, deduplicates overlapping roots, skips lock and old files' {
        [IO.File]::WriteAllText((Join-Path $temp 'book.xls'), 'sample')
        [IO.File]::WriteAllText((Join-Path $temp '~$lock.xls'), 'sample')
        [IO.Directory]::CreateDirectory((Join-Path $temp 'old')) | Out-Null
        [IO.File]::WriteAllText((Join-Path $temp 'old/book.xls'), 'sample')
        New-Xlsx (Join-Path $temp 'already.xlsx')
        $queue = Get-ECQueue -Paths @($temp, (Join-Path $temp 'book.xls')) -Recurse
        Assert ($queue.Files.Count -eq 1) 'Expected only one legacy workbook'
        Assert ($queue.Files[0].Name -eq 'book.xls') 'Name was lost during deduplication'
        Assert ($queue.Files[0].BaseName -eq 'book') 'BaseName was lost'
        Assert ($queue.Files[0].Directory -eq $temp) 'Directory was lost'
        Assert ($queue.AlreadyXlsx.Count -eq 1) 'Existing XLSX should be reported'
    }
    Test 'scan errors and archives are reported instead of silently discarded' {
        $q = Get-ECQueue -Paths @((Join-Path $temp 'missing'), (Join-Path $temp 'input.zip'))
        Assert ($q.Errors.Count -eq 2) 'Missing entries must be visible'
    }
    Test 'literal Unicode and brackets work at paths over 260 characters' {
        $deep = $temp
        1..7 | ForEach-Object { $deep = Join-Path $deep ('long-directory-' + ('x' * 30)) }
        [ExcelConverter.LongIO]::CreateDirectory($deep)
        $script:deep = $deep
        $file = Join-Path $deep 'HR [2024] đối chiếu.xls'
        [IO.File]::WriteAllText((Join-Path $temp 'seed'), 'sample')
        [ExcelConverter.LongIO]::Copy((Join-Path $temp 'seed'), $file)
        $q = Get-ECQueue -Paths @($deep) -Recurse
        Assert ($q.Files.Count -eq 1) 'Long path was not discovered'
        Assert ($q.Files[0].Name -eq 'HR [2024] đối chiếu.xls') 'Literal filename was changed'
        Assert ([ExcelConverter.LongIO]::Hash($file) -eq [ExcelConverter.LongIO]::Hash((Join-Path $temp 'seed'))) 'Long-path copy differs'
    }
    Test 'valid XLSX accepted; macro payload and invalid ZIP rejected' {
        Assert (Test-ECXlsx (Join-Path $temp 'already.xlsx')) 'Valid workbook rejected'
        New-Xlsx (Join-Path $temp 'macro.xlsx') -Macro
        Assert (-not (Test-ECXlsx (Join-Path $temp 'macro.xlsx'))) 'VBA payload accepted'
        Assert (-not (Test-ECXlsx (Join-Path $temp 'book.xls'))) 'Non-ZIP accepted'
    }
    Test 'wrong workbook XML and renamed VBA content types are rejected' {
        New-Xlsx (Join-Path $temp 'wrong-main.xlsx') -InvalidMain
        New-Xlsx (Join-Path $temp 'hidden-macro.xlsx') -HiddenMacroType
        Assert (-not (Test-ECXlsx (Join-Path $temp 'wrong-main.xlsx'))) 'Invalid workbook root accepted'
        Assert (-not (Test-ECXlsx (Join-Path $temp 'hidden-macro.xlsx'))) 'Renamed VBA content type accepted'
    }
    Test 'publication does not overwrite any prior output and only then archives original' {
        $source = Join-Path $temp 'book.xls'
        $hash = [ExcelConverter.LongIO]::Hash($source)
        New-Xlsx (Join-Path $temp 'book.xlsx')
        New-Xlsx (Join-Path $temp 'book_converted.xlsx')
        $r = Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash $hash -OriginalAction Move
        Assert ($r.Status -eq 'Converted') 'Publication failed'
        Assert ($r.Output -eq (Join-Path $temp 'book_converted_2.xlsx')) 'Output collision not handled'
        Assert ([ExcelConverter.LongIO]::Exists((Join-Path $temp 'old/book_converted.xls'))) 'Archive collision not handled'
        Assert (-not [ExcelConverter.LongIO]::Exists($source)) 'Original not archived'
    }
    Test 'changed source is preserved and no output is published' {
        $source = Join-Path $temp 'changing.xls'
        [IO.File]::WriteAllText($source, 'before')
        $hash = [ExcelConverter.LongIO]::Hash($source)
        [IO.File]::WriteAllText($source, 'after')
        $caught = $false
        try { Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash $hash -OriginalAction Delete } catch { $caught = $true }
        Assert $caught 'Changed source not detected'
        Assert ([ExcelConverter.LongIO]::Exists($source)) 'Changed source was deleted'
        Assert (-not [ExcelConverter.LongIO]::Exists((Join-Path $temp 'changing.xlsx'))) 'Stale output was published'
    }
    Test 'invalid converted output cannot cause source deletion' {
        $source = Join-Path $temp 'invalid.xls'
        [IO.File]::WriteAllText($source, 'source')
        $caught = $false
        try { Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'macro.xlsx') -SourceHash ([ExcelConverter.LongIO]::Hash($source)) -OriginalAction Delete } catch { $caught = $true }
        Assert $caught 'Invalid output was accepted'
        Assert ([ExcelConverter.LongIO]::Exists($source)) 'Source deleted before validation'
    }
    Test 'publishing to a long path preserves the source by default' {
        $longSource = (Get-ECQueue -Paths @($script:deep) -Recurse).Files[0].FullName
        $r = Publish-ECWorkbook -Source $longSource -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash ([ExcelConverter.LongIO]::Hash($longSource))
        Assert ($r.Status -eq 'Converted') 'Long-path publication failed'
        Assert ([ExcelConverter.LongIO]::Exists($longSource)) 'Default action changed the source'
        Assert (Test-ECXlsx $r.Output) 'Published long-path XLSX is invalid'
    }
    Test 'publication checkpoint is emitted before original cleanup' {
        $source = Join-Path $temp 'warning.xls'
        [IO.File]::WriteAllText($source, 'source')
        $script:checkpointSeen = $false
        $r = Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash ([ExcelConverter.LongIO]::Hash($source)) -OnPublished {
            param($checkpoint)
            Assert ([ExcelConverter.LongIO]::Exists($checkpoint.Output)) 'Checkpoint before output exists'
            Assert ([ExcelConverter.LongIO]::Exists($source)) 'Source changed before checkpoint'
            $script:checkpointSeen = $true
        }
        Assert $script:checkpointSeen 'No publication checkpoint'
    }
    Test 'original cleanup failure is a warning and preserves the published output' {
        $folder = Join-Path $temp 'cleanup-failure'
        [IO.Directory]::CreateDirectory($folder) | Out-Null
        $source = Join-Path $folder 'source.xls'
        [IO.File]::WriteAllText($source, 'source')
        [IO.File]::WriteAllText((Join-Path $folder 'old'), 'not a directory')
        $r = Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash ([ExcelConverter.LongIO]::Hash($source)) -OriginalAction Move
        Assert ($r.Status -eq 'ConvertedWithWarning') 'Cleanup failure incorrectly reported'
        Assert ([ExcelConverter.LongIO]::Exists($source)) 'Source removed after archive failure'
        Assert (Test-ECXlsx $r.Output) 'Successful output lost'
    }
    Test 'checkpoint failure cannot remove the source or hide the saved output' {
        $source = Join-Path $temp 'checkpoint.xls'
        [IO.File]::WriteAllText($source, 'source')
        $r = Publish-ECWorkbook -Source $source -StagedOutput (Join-Path $temp 'already.xlsx') -SourceHash ([ExcelConverter.LongIO]::Hash($source)) -OriginalAction Delete -OnPublished { throw 'disk full' }
        Assert ($r.Status -eq 'ConvertedWithWarning') 'Checkpoint failure not reported'
        Assert ([ExcelConverter.LongIO]::Exists($source)) 'Source deleted despite checkpoint failure'
        Assert (Test-ECXlsx $r.Output) 'Saved output hidden by checkpoint failure'
    }
    Write-Host "$script:passed tests passed"
} finally { [ExcelConverter.LongIO]::DeleteTree($temp) }
