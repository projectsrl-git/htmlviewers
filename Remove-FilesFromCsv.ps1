<#
.SYNOPSIS
    Deletes files from a directory, taking their names from a column of a CSV.

.DESCRIPTION
    Reads one column of a CSV, treats every value as a bare file name, and removes
    the matching file from the target directory. No prompt is issued.

    Deletion is not recoverable, so the script refuses to act on anything it was not
    clearly told to act on:

      - a value containing a path separator, a drive letter, or '..' is skipped and
        reported. The column holds names, not paths: anything else is either a
        mistake or an attempt to reach outside the directory.
      - every resolved target is checked to sit directly inside -Directory. A file
        elsewhere is never touched, whatever the CSV says.
      - wildcards are not expanded. A value of '*' deletes a file literally named
        '*', or more likely nothing at all, rather than the whole directory.

    -MoveTo turns the operation into a move, which is the safer form when the list is
    not yet trusted: the files leave the directory but still exist.

    Use -WhatIf for a dry run. Plain timestamped stdout, suitable as an OpenProteo
    exec step.

.PARAMETER Directory
    Directory to delete from. Only its immediate content is considered.

.PARAMETER CsvPath
    CSV holding the names.

.PARAMETER Column
    Column name in the CSV header.

.PARAMETER ColumnIndex
    Zero-based column position, for a CSV without a usable header. Mutually exclusive
    with -Column.

.PARAMETER Delimiter
    CSV delimiter. Default ';'.

.PARAMETER Encoding
    Charset used to read the CSV. Default windows-1252.

.PARAMETER Prefix
    Prepended to every name from the CSV.

.PARAMETER Suffix
    Appended to every name from the CSV, for a column holding names without their
    extension.

.PARAMETER MoveTo
    Move the files into this directory instead of deleting them. Created if absent.

.PARAMETER IgnoreMissing
    Do not count a name with no matching file as a failure. Missing files still
    appear in the log.

.PARAMETER MaxReport
    Maximum names listed per category in the closing summary. Default 50. Counts stay
    exact.

.PARAMETER SummaryOnly
    Suppress the per-file lines. Only the periodic progress and the summary are
    printed, which is what a scheduled run wants once the list is trusted.

.PARAMETER ProgressSeconds
    Interval between progress lines. Default 5. Set 0 to disable.

.EXAMPLE
    .\Remove-FilesFromCsv.ps1 -Directory G:\ELAR\OUT\CMOD\S210969_CLIAC `
                              -CsvPath .\to_delete.csv -Column FileName -WhatIf

.EXAMPLE
    .\Remove-FilesFromCsv.ps1 -Directory G:\PROTEO\PDF -CsvPath .\list.csv `
                              -Column RecordId -Suffix .pdf -MoveTo G:\PROTEO\_removed

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 completed with skipped or missing entries, 1 error.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Directory,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $CsvPath,

    [string] $Column,

    [int] $ColumnIndex = -1,

    [string] $Delimiter = ';',

    [string] $Encoding = 'windows-1252',

    [string] $Prefix,

    [string] $Suffix,

    [string] $MoveTo,

    [switch] $IgnoreMissing,

    [int] $MaxReport = 50,

    [switch] $SummaryOnly,

    [int] $ProgressSeconds = 5
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

$exitCode = 0

try {
    if (-not $Column -and $ColumnIndex -lt 0) { throw "Give either -Column or -ColumnIndex." }
    if ($Column -and $ColumnIndex -ge 0)      { throw "-Column and -ColumnIndex are mutually exclusive." }

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { throw "Directory not found: $Directory" }
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf))        { throw "CSV not found: $CsvPath" }

    $dir = (Get-Item -LiteralPath $Directory).FullName.TrimEnd('\', '/')

    if ($MoveTo) {
        if (-not (Test-Path -LiteralPath $MoveTo)) { New-Item -ItemType Directory -Path $MoveTo -Force | Out-Null }
        $moveDir = (Get-Item -LiteralPath $MoveTo).FullName.TrimEnd('\', '/')
        if ($moveDir -eq $dir) { throw "-MoveTo is the same directory as -Directory." }
    }

    try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
    catch { throw "Unknown charset '$Encoding'." }

    # --- read the names -----------------------------------------------------
    $swCsv = [System.Diagnostics.Stopwatch]::StartNew()
    $names = New-Object 'System.Collections.Generic.List[string]'
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]'
    $blank = 0; $dupes = 0

    if ($Column) {
        # Import-Csv understands quoting, so a value containing the delimiter is read
        # correctly rather than splitting the row.
        $rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter $Delimiter -Encoding ([string]$Encoding) -ErrorAction Stop)
        if ($rows.Count -eq 0) { throw "CSV has no data rows." }
        if (-not ($rows[0].PSObject.Properties.Name -contains $Column)) {
            throw ("Column '{0}' not found. Header is: {1}" -f $Column, (($rows[0].PSObject.Properties.Name) -join ', '))
        }
        foreach ($r in $rows) {
            $v = [string]$r.$Column
            if ([string]::IsNullOrWhiteSpace($v)) { $blank++; continue }
            $v = $v.Trim()
            if ($seen.Add($v)) { $names.Add($v) } else { $dupes++ }
        }
    }
    else {
        $reader = New-Object System.IO.StreamReader($CsvPath, $enc, $false)
        try {
            $first = $true
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($first) { $first = $false; continue }   # skip header row
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $parts = $line.Split($Delimiter[0])
                if ($ColumnIndex -ge $parts.Length) { $blank++; continue }
                $v = $parts[$ColumnIndex].Trim().Trim('"')
                if ([string]::IsNullOrWhiteSpace($v)) { $blank++; continue }
                if ($seen.Add($v)) { $names.Add($v) } else { $dupes++ }
            }
        }
        finally { $reader.Dispose() }
    }

    $swCsv.Stop()

    $action = if ($MoveTo) { 'MOVE' } else { 'DELETE' }
    Log ("START action={0} directory={1}" -f $action, $dir)
    Log ("  csv={0} column={1} names={2} blank={3} duplicates={4}" -f `
         (Split-Path -Leaf $CsvPath), $(if ($Column) { $Column } else { "index $ColumnIndex" }), `
         $names.Count, $blank, $dupes)
    if ($MoveTo) { Log ("  moveTo={0}" -f $moveDir) }
    Log ("  csvRead={0:n1}s" -f $swCsv.Elapsed.TotalSeconds)

    # --- index the directory once -------------------------------------------
    # One listing instead of one Test-Path per name. On a network share each stat is
    # a round trip, so this is the difference between minutes and seconds.
    $swIdx = [System.Diagnostics.Stopwatch]::StartNew()
    $index = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
        [void]$index.Add([System.IO.Path]::GetFileName($f))
    }
    $swIdx.Stop()
    Log ("  directoryIndexed files={0:n0} in {1:n1}s" -f $index.Count, $swIdx.Elapsed.TotalSeconds)

    # --- act ----------------------------------------------------------------
    $done = 0; $missing = 0; $unsafe = 0; $failed = 0; $processed = 0
    $listMissing = New-Object 'System.Collections.Generic.List[string]'
    $listUnsafe  = New-Object 'System.Collections.Generic.List[string]'
    $swAct = [System.Diagnostics.Stopwatch]::StartNew()
    $nextTick = if ($ProgressSeconds -gt 0) { [double]$ProgressSeconds } else { [double]::MaxValue }

    foreach ($n in $names) {
        $processed++
        $target = $Prefix + $n + $Suffix

        # A name, not a path. Anything else is refused rather than interpreted.
        if ($target.IndexOfAny([char[]]@('\', '/')) -ge 0 -or
            $target.Contains('..') -or $target.Contains(':')) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            if (-not $SummaryOnly) { Log ("  SKIP  {0}  (refused: not a bare file name)" -f $target) }
            continue
        }
        if ($target.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            if (-not $SummaryOnly) { Log ("  SKIP  {0}  (refused: not a bare file name)" -f $target) }
            continue
        }

        $full = Join-Path $dir $target

        # Second line of defence: the resolved parent must be the target directory.
        $parent = [System.IO.Path]::GetDirectoryName($full)
        if ($parent.TrimEnd('\', '/') -ne $dir) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            if (-not $SummaryOnly) { Log ("  SKIP  {0}  (refused: not a bare file name)" -f $target) }
            continue
        }

        if (-not $index.Contains($target)) {
            $missing++
            if ($listMissing.Count -lt $MaxReport) { [void]$listMissing.Add($target) }
            if (-not $SummaryOnly) { Log ("  SKIP  {0}  (not found)" -f $target) }
            continue
        }

        try {
            if ($MoveTo) {
                if ($PSCmdlet.ShouldProcess($full, "Move to $moveDir")) {
                    # The .NET calls avoid the per-item cmdlet overhead, which
                    # dominates when the list is long.
                    [System.IO.File]::Move($full, (Join-Path $moveDir $target))
                    $done++
                    if (-not $SummaryOnly) { Log ("  MOVE  {0}" -f $target) }
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($full, 'Delete')) {
                    [System.IO.File]::Delete($full)
                    $done++
                    if (-not $SummaryOnly) { Log ("  DEL   {0}" -f $target) }
                }
            }
        }
        catch {
            $failed++
            Log ("  FAIL  {0}: {1}" -f $target, $_.Exception.Message)
        }

        if ($ProgressSeconds -gt 0 -and $swAct.Elapsed.TotalSeconds -ge $nextTick) {
            $nextTick = $swAct.Elapsed.TotalSeconds + $ProgressSeconds
            $rate = if ($swAct.Elapsed.TotalSeconds -gt 0) { $processed / $swAct.Elapsed.TotalSeconds } else { 0 }
            $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds(($names.Count - $processed) / $rate) } else { [TimeSpan]::Zero }
            Log ("  ... {0:n0}/{1:n0}  done={2:n0} missing={3:n0} refused={4:n0}  {5:n0}/s  ETA {6:mm\:ss}" -f `
                 $processed, $names.Count, $done, $missing, $unsafe, $rate, $eta)
        }
    }
    $swAct.Stop()

    Log "SUMMARY"
    Log ("  namesRead={0}" -f $names.Count)
    Log ("  {0}={1}" -f $(if ($MoveTo) { 'moved' } else { 'deleted' }), $done)
    Log ("  notFound={0}" -f $missing)
    Log ("  refusedUnsafeName={0}" -f $unsafe)
    Log ("  failed={0}" -f $failed)
    Log ("  elapsed csv={0:n1}s index={1:n1}s action={2:n1}s" -f `
         $swCsv.Elapsed.TotalSeconds, $swIdx.Elapsed.TotalSeconds, $swAct.Elapsed.TotalSeconds)

    if ($listUnsafe.Count -gt 0) {
        Log "  refused (a name must not contain a path separator, '..' or ':'):"
        foreach ($x in $listUnsafe) { Log ("    {0}" -f $x) }
        if ($unsafe -gt $listUnsafe.Count) { Log ("    ... +{0} more" -f ($unsafe - $listUnsafe.Count)) }
    }
    if ($listMissing.Count -gt 0) {
        Log "  not found:"
        foreach ($x in $listMissing) { Log ("    {0}" -f $x) }
        if ($missing -gt $listMissing.Count) { Log ("    ... +{0} more" -f ($missing - $listMissing.Count)) }
    }

    if ($failed -gt 0) { $exitCode = 1 }
    elseif ($unsafe -gt 0 -or ($missing -gt 0 -and -not $IgnoreMissing)) { $exitCode = 2 }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
