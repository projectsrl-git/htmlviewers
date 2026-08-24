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
    Maximum names listed per category. Default 50. Counts stay exact.

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

    [int] $MaxReport = 50
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

    $action = if ($MoveTo) { 'MOVE' } else { 'DELETE' }
    Log ("START action={0} directory={1}" -f $action, $dir)
    Log ("  csv={0} column={1} names={2} blank={3} duplicates={4}" -f `
         (Split-Path -Leaf $CsvPath), $(if ($Column) { $Column } else { "index $ColumnIndex" }), `
         $names.Count, $blank, $dupes)
    if ($MoveTo) { Log ("  moveTo={0}" -f $moveDir) }

    # --- act ----------------------------------------------------------------
    $done = 0; $missing = 0; $unsafe = 0; $failed = 0
    $listMissing = New-Object 'System.Collections.Generic.List[string]'
    $listUnsafe  = New-Object 'System.Collections.Generic.List[string]'

    foreach ($n in $names) {
        $target = $Prefix + $n + $Suffix

        # A name, not a path. Anything else is refused rather than interpreted.
        if ($target.IndexOfAny([char[]]@('\', '/')) -ge 0 -or
            $target.Contains('..') -or $target.Contains(':')) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            continue
        }
        if ($target.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            continue
        }

        $full = Join-Path $dir $target

        # Second line of defence: the resolved parent must be the target directory.
        $parent = [System.IO.Path]::GetDirectoryName($full)
        if ($parent.TrimEnd('\', '/') -ne $dir) {
            $unsafe++
            if ($listUnsafe.Count -lt $MaxReport) { [void]$listUnsafe.Add($target) }
            continue
        }

        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
            $missing++
            if ($listMissing.Count -lt $MaxReport) { [void]$listMissing.Add($target) }
            continue
        }

        try {
            if ($MoveTo) {
                if ($PSCmdlet.ShouldProcess($full, "Move to $moveDir")) {
                    Move-Item -LiteralPath $full -Destination (Join-Path $moveDir $target) -Force
                    $done++
                }
            }
            else {
                if ($PSCmdlet.ShouldProcess($full, 'Delete')) {
                    Remove-Item -LiteralPath $full -Force
                    $done++
                }
            }
        }
        catch {
            $failed++
            Log ("  FAILED {0}: {1}" -f $target, $_.Exception.Message)
        }
    }

    Log "SUMMARY"
    Log ("  namesRead={0}" -f $names.Count)
    Log ("  {0}={1}" -f $(if ($MoveTo) { 'moved' } else { 'deleted' }), $done)
    Log ("  notFound={0}" -f $missing)
    Log ("  refusedUnsafeName={0}" -f $unsafe)
    Log ("  failed={0}" -f $failed)

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
