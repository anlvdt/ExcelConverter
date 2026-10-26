Set-StrictMode -Version 2
if (-not ('ExcelConverter.LongIO' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'LongPathIO.cs') }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ECQueue {
    param([string[]]$Paths, [switch]$Recurse, [scriptblock]$OnProgress = {})
    $files = [Collections.Generic.List[object]]::new()
    $xlsx = [Collections.Generic.List[object]]::new()
    $errors = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $pending = [Collections.Generic.Stack[string]]::new()
    $addFile = {
        param($file)
        if ($file.Name.StartsWith('~$') -or $file.IsLink -or -not $seen.Add($file.FullName)) { return }
        switch ($file.Extension.ToLowerInvariant()) {
            '.xlsx' { $xlsx.Add($file) }
            { $_ -in '.xls', '.xlsm', '.xlsb' } { $files.Add($file) }
            default { }
        }
    }
    foreach ($path in $Paths) {
        try {
            $full = [ExcelConverter.LongIO]::FullPath($path)
            if (-not [ExcelConverter.LongIO]::Exists($full)) { throw "Path does not exist: $path" }
            if ([ExcelConverter.LongIO]::IsDirectory($full)) { $pending.Push($full) }
            else {
                $file = [ExcelConverter.FileEntry]::new($full, $false, $false)
                if ($file.Extension -notin '.xls', '.xlsm', '.xlsb', '.xlsx') { throw "Unsupported input (extract ZIP/7z first): $path" }
                & $addFile $file
            }
        } catch { $errors.Add([pscustomobject]@{ Source = $path; Status = 'ScanError'; Error = $_.Exception.Message }) }
    }
    while ($pending.Count -gt 0) {
        if ((& $OnProgress) -eq $false) { break }
        $dir = $pending.Pop()
        if (-not $visited.Add($dir)) { continue }
        try {
            foreach ($file in [ExcelConverter.LongIO]::List($dir)) {
                if ($file.IsDirectory) {
                    if ($Recurse -and -not $file.IsLink -and $file.Name -ine 'old') { $pending.Push($file.FullName) }
                } else { & $addFile $file }
            }
        } catch { $errors.Add([pscustomobject]@{ Source = $dir; Status = 'ScanError'; Error = $_.Exception.Message }) }
    }
    [pscustomobject]@{ Files = $files.ToArray(); AlreadyXlsx = $xlsx.ToArray(); Errors = $errors.ToArray() }
}

function Test-ECXlsx {
    param([string]$Path)
    $stream = $null; $zip = $null
    try {
        $stream = [ExcelConverter.LongIO]::OpenRead($Path)
        $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read)
        foreach ($required in '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels') {
            if (-not $zip.GetEntry($required)) { return $false }
        }
        $types = $null; $workbook = $null
        $partNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $zip.Entries) {
            if (-not $partNames.Add($entry.FullName)) { return $false }
            if ($entry.FullName -match '(?i)(vbaProject|macrosheets/|intlmacrosheets/)') { return $false }
            if ($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) {
                $settings = [Xml.XmlReaderSettings]::new()
                $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
                $settings.XmlResolver = $null
                $part = $entry.Open(); $reader = $null
                try {
                    $reader = [Xml.XmlReader]::Create($part, $settings)
                    $xml = [Xml.XmlDocument]::new(); $xml.XmlResolver = $null; $xml.Load($reader)
                    if ($entry.FullName -eq '[Content_Types].xml') { $types = $xml }
                    if ($entry.FullName -eq 'xl/workbook.xml') { $workbook = $xml }
                } finally { if ($reader) { $reader.Dispose() }; $part.Dispose() }
            } else {
                $part = $entry.Open()
                try { $part.CopyTo([IO.Stream]::Null) } finally { $part.Dispose() }
            }
        }
        if ($workbook.DocumentElement.LocalName -ne 'workbook' -or $workbook.DocumentElement.NamespaceURI -notin 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main') { return $false }
        foreach ($type in $types.DocumentElement.ChildNodes) {
            if ($type -is [Xml.XmlElement] -and $type.GetAttribute('ContentType') -match '(?i)(vbaProject|macroEnabled|macrosheet)') { return $false }
        }
        $main = $types.DocumentElement.ChildNodes | Where-Object { $_ -is [Xml.XmlElement] -and $_.GetAttribute('PartName') -eq '/xl/workbook.xml' }
        return (@($main).Count -eq 1 -and $main.GetAttribute('ContentType') -eq 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    } catch { return $false }
    finally { if ($zip) { $zip.Dispose() }; if ($stream) { $stream.Dispose() } }
}

function Get-ECFreePath {
    param([string]$Directory, [string]$BaseName, [string]$Extension)
    for ($n = 0; $n -lt 10000; $n++) {
        $suffix = if ($n -eq 0) { '' } elseif ($n -eq 1) { '_converted' } else { "_converted_$n" }
        # Windows and OneDrive both limit individual name components to 255 characters.
        $maxBase = 255 - $suffix.Length - $Extension.Length
        $stem = $BaseName.Substring(0, [Math]::Min($BaseName.Length, $maxBase))
        $candidate = [IO.Path]::Combine($Directory, "$stem$suffix$Extension")
        if (-not [ExcelConverter.LongIO]::Exists($candidate)) { return $candidate }
    }
    throw 'Too many filename collisions'
}

function Publish-ECWorkbook {
    param([string]$Source, [string]$StagedOutput, [string]$SourceHash,
          [ValidateSet('Keep', 'Move', 'Delete')][string]$OriginalAction = 'Keep', [scriptblock]$OnPublished = {})
    if (-not (Test-ECXlsx $StagedOutput)) { throw 'Converted output is not a valid macro-free XLSX package' }
    if ([ExcelConverter.LongIO]::Hash($Source) -ne $SourceHash) { throw 'Source changed during conversion; original preserved' }
    $directory = [ExcelConverter.LongIO]::Parent($Source)
    $base = [IO.Path]::GetFileNameWithoutExtension([ExcelConverter.LongIO]::Name($Source))
    # Copy into the destination directory first, verify, then rename without replacement.
    $partial = [IO.Path]::Combine($directory, ('.ec-' + [guid]::NewGuid().ToString('N') + '.tmp'))
    $output = $null
    try {
        [ExcelConverter.LongIO]::Copy($StagedOutput, $partial)
        if ([ExcelConverter.LongIO]::Hash($partial) -ne [ExcelConverter.LongIO]::Hash($StagedOutput)) { throw 'Destination verification failed' }
        for ($attempt = 0; $attempt -lt 10; $attempt++) {
            $output = Get-ECFreePath $directory $base '.xlsx'
            try { [ExcelConverter.LongIO]::Move($partial, $output); break }
            catch { if ($attempt -eq 9 -or -not [ExcelConverter.LongIO]::Exists($output)) { throw } }
        }
    } finally {
        if ([ExcelConverter.LongIO]::Exists($partial)) { [ExcelConverter.LongIO]::Delete($partial) }
    }
    $warning = ''; $archive = ''
    try {
        & $OnPublished ([pscustomobject]@{ Source = $Source; Output = $output; Archive = ''; Status = 'OutputSaved'; Error = 'Output published; original cleanup not confirmed' })
    } catch { $warning = 'Output saved, but checkpoint failed; original preserved: ' + $_.Exception.Message }
    try {
        if ($OriginalAction -ne 'Keep' -and -not $warning) {
            if ([ExcelConverter.LongIO]::Hash($Source) -ne $SourceHash) { throw 'Source changed after publication; original preserved' }
            if ($OriginalAction -eq 'Delete') { [ExcelConverter.LongIO]::Delete($Source) }
            else {
                $old = [IO.Path]::Combine($directory, 'old')
                [ExcelConverter.LongIO]::CreateDirectory($old)
                $archive = Get-ECFreePath $old $base ([IO.Path]::GetExtension([ExcelConverter.LongIO]::Name($Source)))
                [ExcelConverter.LongIO]::Move($Source, $archive)
            }
        }
    } catch { $warning = 'Output saved, but original cleanup failed: ' + $_.Exception.Message }
    [pscustomobject]@{ Source = $Source; Output = $output; Archive = $archive; Status = $(if ($warning) { 'ConvertedWithWarning' } else { 'Converted' }); Error = $warning }
}


function Stop-ECOwnedExcel {
    param([string]$JobDirectory)
    $path = Join-Path $JobDirectory 'owner.json'
    if (-not [IO.File]::Exists($path)) { return }
    try {
        $owner = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $process = Get-Process -Id ([int]$owner.Id) -ErrorAction Stop
        if ($process.ProcessName -ieq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq [long]$owner.StartTicks) {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
        }
    } catch { }
}

function Invoke-ECFile {
    param([Parameter(Mandatory = $true)][string]$Source,
          [ValidateSet('Keep', 'Move', 'Delete')][string]$OriginalAction = 'Keep',
          [ValidateRange(1, 86400)][int]$TimeoutSeconds = 180,
          [bool]$ShowExcel = $false, [scriptblock]$OnProgress = {},
          [string]$WorkerPath = (Join-Path $PSScriptRoot 'ConvertXLS_Worker.ps1'))
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $job = $null; $process = $null; $result = $null
    try {
        $tempRoot = if ($env:EXCELCONVERTER_TEMP) { $env:EXCELCONVERTER_TEMP } else { [IO.Path]::GetTempPath() }
        $job = [IO.Path]::Combine($tempRoot, ('EC-' + [guid]::NewGuid().ToString('N').Substring(0, 12)))
        $job = [IO.Path]::GetFullPath($job)
        if (($job.Length + 40) -gt 200 -or $job -match '[\[\]]') { throw 'Temporary path is too long or contains brackets. Set EXCELCONVERTER_TEMP to a short local folder, e.g. C:\ECtemp' }
        [IO.Directory]::CreateDirectory($job) | Out-Null
        foreach ($name in 'ExcelConverter.Core.psm1', 'LongPathIO.cs') {
            [ExcelConverter.LongIO]::Copy((Join-Path $PSScriptRoot $name), (Join-Path $job $name))
        }
        [ExcelConverter.LongIO]::Copy($WorkerPath, (Join-Path $job 'worker.ps1'))
        @{ Source = $Source; OriginalAction = $OriginalAction; ShowExcel = $ShowExcel } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $job 'request.json') -Encoding UTF8
        $engine = if ($env:OS -eq 'Windows_NT') {
            if ($PSVersionTable.PSEdition -eq 'Core') { Join-Path $PSHOME 'pwsh.exe' } else { Join-Path $PSHOME 'powershell.exe' }
        } else { Join-Path $PSHOME 'pwsh' }
        $start = [Diagnostics.ProcessStartInfo]::new()
        $start.FileName = $engine
        $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -STA -File "' + (Join-Path $job 'worker.ps1') + '" -JobDirectory "' + $job + '"'
        if ($env:OS -ne 'Windows_NT') { $start.Arguments = $start.Arguments.Replace(' -STA', '') }
        $start.UseShellExecute = $false; $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
        $process = [Diagnostics.Process]::Start($start)
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $stopStatus = ''
        while (-not $process.WaitForExit(100)) {
            if ((& $OnProgress) -eq $false) { $stopStatus = 'Canceled'; break }
            if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) { $stopStatus = 'TimedOut'; break }
        }
        if ($stopStatus) {
            $process.Kill(); $process.WaitForExit(5000) | Out-Null
            Stop-ECOwnedExcel $job
        }
        $resultPath = Join-Path $job 'result.json'
        $checkpoint = Join-Path $job 'published.json'
        if ([IO.File]::Exists($resultPath)) { $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json }
        elseif ([IO.File]::Exists($checkpoint)) { $result = Get-Content -LiteralPath $checkpoint -Raw | ConvertFrom-Json }
        else {
            $errorText = if ($stopStatus) { "$stopStatus after $([int]$watch.Elapsed.TotalSeconds)s. Cloud file may be offline, locked, or Excel waiting on a dialog." } else { $stderr.GetAwaiter().GetResult() }
            if (-not $errorText) { $errorText = "Worker exited without a result (exit $($process.ExitCode))" }
            $result = [pscustomobject]@{ Source = $Source; Output = ''; Archive = ''; Status = $(if ($stopStatus) { $stopStatus } else { 'Failed' }); Error = $errorText }
        }
        # A completed conversion remains successful even when Excel.Quit times out.
        if ($stopStatus -and $result.Status -in 'Converted', 'ConvertedWithWarning') {
            $result.Status = 'ConvertedWithWarning'
            $result.Error = ($result.Error + ' Excel cleanup interrupted: ' + $stopStatus).Trim()
        }
    } catch {
        $result = [pscustomobject]@{ Source = $Source; Output = ''; Archive = ''; Status = 'Failed'; Error = $_.Exception.Message }
    } finally {
        if ($process) {
            try { if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit(5000) | Out-Null } } catch { }
            if ($job) { Stop-ECOwnedExcel $job }
            $process.Dispose()
        }
        if ($job -and [IO.Directory]::Exists($job)) {
            try {
                foreach ($localFile in [IO.Directory]::GetFiles($job)) {
                    $attributes = [IO.File]::GetAttributes($localFile)
                    [IO.File]::SetAttributes($localFile, [IO.FileAttributes]([int]$attributes -band (-bnot [int][IO.FileAttributes]::ReadOnly)))
                }
                [IO.Directory]::Delete($job, $true)
            } catch {
                if ($result) {
                    if ($result.Status -eq 'Converted') { $result.Status = 'ConvertedWithWarning' }
                    $result.Error = ($result.Error + " Temporary files could not be removed: $job. " + $_.Exception.Message).Trim()
                }
            }
        }
    }
    $result | Add-Member -NotePropertyName DurationSeconds -NotePropertyValue ([Math]::Round($watch.Elapsed.TotalSeconds, 2)) -Force
    return $result
}
Export-ModuleMember -Function Get-ECQueue, Test-ECXlsx, Publish-ECWorkbook, Invoke-ECFile
