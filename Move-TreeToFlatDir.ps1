<#
.SYNOPSIS
    Flattens a directory tree into a single directory by moving the files, deciding
    what to do when two files end up with the same name.

.DESCRIPTION
    Walks the source tree, moves every matching file into one destination directory,
    and applies a policy when the name is already taken.

    Speed comes from three things:

      - the destination is indexed once into memory, so deciding whether a name is
        taken costs nothing per file instead of one filesystem probe each. On a
        network share that is the difference between minutes and seconds.
      - moves use the .NET call rather than the cmdlet, avoiding per-item pipeline
        overhead.
      - within the same volume a move is a rename: the bytes are not copied. Across
        volumes it is a copy followed by a delete, and there is no way around that.
        The summary reports the two cases separately so a slow run is explainable.

    "Duplicate" means, by default, a name already present in the destination. That is
    a name-level test and costs nothing. -CompareContent adds a content test for the
    clashing pairs only: same size and same SHA-256 means truly identical, and the
    source copy is deleted or kept according to -OnIdentical. Different content under
    the same name is a different situation entirely, handled by -OnDifferent.

    Nothing is deleted implicitly: a file is either moved, left where it is, or - only
    with -OnIdentical Delete - removed because a byte-identical copy already exists at
    the destination.

    Use -WhatIf for a dry run. Plain timestamped stdout, no prompts.

.PARAMETER Source
    Root of the tree to flatten.

.PARAMETER Destination
    Directory receiving all the files. Created if absent. It may not sit inside the
    source tree.

.PARAMETER Filter
    Wildcard applied to file names. Default * .

.PARAMETER SkipExisting
    Skip any file whose name already exists at the destination, without looking at
    the content. This is the fastest mode and the default behaviour; the switch makes
    it explicit and overrides -OnDifferent and -CompareContent, so a command line
    that says it cannot accidentally do something else.

.PARAMETER OnDifferent
    What to do when the name is taken by a file with different content, or when the
    content is not compared. Skip (default) leaves the source file untouched; Rename
    moves it as name-1.ext, name-2.ext; Overwrite replaces the destination file.

.PARAMETER CompareContent
    For clashing names only, compare size and then SHA-256 to tell an identical copy
    from a different file. Costs a full read of both files in the clashing pairs, so
    it is off by default.

.PARAMETER OnIdentical
    With -CompareContent, what to do when the two files are byte-identical: Skip
    (default) leaves the source in place, Delete removes it.

.PARAMETER RemoveEmptyDirs
    After moving, delete the source subdirectories left empty. The source root itself
    is never removed.

.PARAMETER SummaryOnly
    Suppress the per-file lines, keeping the periodic progress and the summary.

.PARAMETER ProgressSeconds
    Interval between progress lines. Default 5. Set 0 to disable.

.PARAMETER MaxReport
    Maximum names listed per category in the summary. Default 30.

.PARAMETER LogFile
    Write every line to this file as well as to stdout. If the path is an existing
    directory, a file named move-tree-<timestamp>.log is created inside it. The file
    is written as it goes, not buffered to the end, so an interrupted run still leaves
    a usable record of what was moved.

.PARAMETER AppendLog
    Append to -LogFile instead of replacing it.

.EXAMPLE
    .\Move-TreeToFlatDir.ps1 -Source G:\PROTEO\PDF -Destination G:\PROTEO\FLAT -WhatIf

.EXAMPLE
    .\Move-TreeToFlatDir.ps1 -Source G:\PROTEO\PDF -Destination G:\PROTEO\FLAT `
        -SkipExisting -LogFile G:\PROTEO\_log

.EXAMPLE
    .\Move-TreeToFlatDir.ps1 -Source G:\PROTEO\PDF -Destination G:\PROTEO\FLAT `
        -CompareContent -OnIdentical Delete -OnDifferent Rename -RemoveEmptyDirs

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 completed with skipped files, 1 error.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Source,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $Destination,

    [string] $Filter = '*',

    [ValidateSet('Skip', 'Rename', 'Overwrite')]
    [string] $OnDifferent = 'Skip',

    [switch] $CompareContent,

    [ValidateSet('Skip', 'Delete')]
    [string] $OnIdentical = 'Skip',

    [switch] $RemoveEmptyDirs,

    [switch] $SummaryOnly,

    [int] $ProgressSeconds = 5,

    [int] $MaxReport = 30,

    [switch] $SkipExisting,

    [string] $LogFile,

    [switch] $AppendLog
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:LogWriter = $null

function Log {
    param([string] $Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    [Console]::Out.WriteLine($line)
    [Console]::Out.Flush()
    if ($script:LogWriter) {
        # Flushed per line: an interrupted run still leaves a record of what moved.
        $script:LogWriter.WriteLine($line)
        $script:LogWriter.Flush()
    }
}

function Open-LogFile {
    param([string] $Path, [bool] $Append)

    if (-not $Path) { return }

    if (Test-Path -LiteralPath $Path -PathType Container) {
        $Path = Join-Path $Path ('move-tree-{0:yyyyMMdd-HHmmss}.log' -f (Get-Date))
    }
    else {
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    try {
        # UTF-8 without BOM: readable by any viewer and by a later parse.
        $enc = New-Object System.Text.UTF8Encoding($false)
        $script:LogWriter = New-Object System.IO.StreamWriter($Path, $Append, $enc)
        $script:LogPath = (Get-Item -LiteralPath $Path).FullName
    }
    catch {
        # A missing log must not abort the move itself.
        [Console]::Out.WriteLine(('WARNING could not open log file {0}: {1}' -f $Path, $_.Exception.Message))
        $script:LogWriter = $null
    }
}

function Get-Sha256 {
    param([string] $Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        return ([BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '')
    }
    finally {
        if ($fs) { $fs.Dispose() }
        $sha.Dispose()
    }
}

$exitCode = 0

try {
    Open-LogFile -Path $LogFile -Append $AppendLog.IsPresent

    if ($SkipExisting) {
        # Explicit wins over the individual settings, so the intent on the command
        # line is what actually runs.
        $OnDifferent = 'Skip'
        if ($CompareContent) { $CompareContent = $false }
    }

    if (-not (Test-Path -LiteralPath $Source -PathType Container)) { throw "Source not found: $Source" }
    $src = (Get-Item -LiteralPath $Source).FullName.TrimEnd('\', '/')

    if (-not (Test-Path -LiteralPath $Destination)) {
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    }
    $dst = (Get-Item -LiteralPath $Destination).FullName.TrimEnd('\', '/')

    # A destination inside the tree would be re-scanned as it fills up.
    if ($dst -eq $src -or $dst.StartsWith($src + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Destination must not be inside the source tree."
    }

    $sameVolume = ([System.IO.Path]::GetPathRoot($src)) -eq ([System.IO.Path]::GetPathRoot($dst))

    Log ("START source={0}" -f $src)
    Log ("  destination={0}" -f $dst)
    Log ("  filter={0} onDifferent={1} compareContent={2} onIdentical={3}{4}" -f `
         $Filter, $OnDifferent, $CompareContent.IsPresent, $OnIdentical,
         $(if ($SkipExisting) { ' (forced by -SkipExisting)' } else { '' }))
    if ($script:LogWriter) { Log ("  log={0}" -f $script:LogPath) }
    Log ("  sameVolume={0}  ({1})" -f $sameVolume, $(if ($sameVolume) { 'moves are renames, no bytes copied' } else { 'moves copy then delete' }))

    # --- index the destination once -----------------------------------------
    $swIdx = [System.Diagnostics.Stopwatch]::StartNew()
    $index = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($f in [System.IO.Directory]::EnumerateFiles($dst)) {
        [void]$index.Add([System.IO.Path]::GetFileName($f))
    }
    $swIdx.Stop()
    Log ("  destinationIndexed files={0:n0} in {1:n1}s" -f $index.Count, $swIdx.Elapsed.TotalSeconds)

    # --- walk and move -------------------------------------------------------
    $moved = 0; $skipped = 0; $renamed = 0; $overwritten = 0
    $identicalDeleted = 0; $identicalKept = 0; $failed = 0; $seenFiles = 0

    $listSkipped = New-Object 'System.Collections.Generic.List[string]'
    $swAct = [System.Diagnostics.Stopwatch]::StartNew()
    $nextTick = if ($ProgressSeconds -gt 0) { [double]$ProgressSeconds } else { [double]::MaxValue }

    $enumOpt = [System.IO.SearchOption]::AllDirectories
    foreach ($path in [System.IO.Directory]::EnumerateFiles($src, $Filter, $enumOpt)) {
        $seenFiles++
        $name = [System.IO.Path]::GetFileName($path)
        $target = Join-Path $dst $name

        if (-not $index.Contains($name)) {
            try {
                if ($PSCmdlet.ShouldProcess($path, "Move to $dst")) {
                    [System.IO.File]::Move($path, $target)
                    [void]$index.Add($name)
                    $moved++
                    if (-not $SummaryOnly) { Log ("  MOVE  {0}" -f $name) }
                }
            }
            catch { $failed++; Log ("  FAIL  {0}: {1}" -f $name, $_.Exception.Message) }
        }
        else {
            # The name is taken. Decide what kind of clash this is.
            $identical = $false
            if ($CompareContent) {
                try {
                    $a = Get-Item -LiteralPath $path
                    $b = Get-Item -LiteralPath $target
                    if ($a.Length -eq $b.Length) {
                        # Size first: it settles most pairs without reading anything.
                        $identical = ((Get-Sha256 -Path $path) -eq (Get-Sha256 -Path $target))
                    }
                }
                catch { $failed++; Log ("  FAIL  {0}: {1}" -f $name, $_.Exception.Message); continue }
            }

            if ($identical) {
                if ($OnIdentical -eq 'Delete') {
                    try {
                        if ($PSCmdlet.ShouldProcess($path, 'Delete, identical copy already at destination')) {
                            [System.IO.File]::Delete($path)
                            $identicalDeleted++
                            if (-not $SummaryOnly) { Log ("  DUPE  {0}  (identical, source deleted)" -f $name) }
                        }
                    }
                    catch { $failed++; Log ("  FAIL  {0}: {1}" -f $name, $_.Exception.Message) }
                }
                else {
                    $identicalKept++
                    if (-not $SummaryOnly) { Log ("  DUPE  {0}  (identical, source kept)" -f $name) }
                }
            }
            else {
                switch ($OnDifferent) {

                    'Overwrite' {
                        try {
                            if ($PSCmdlet.ShouldProcess($path, "Overwrite $target")) {
                                [System.IO.File]::Delete($target)
                                [System.IO.File]::Move($path, $target)
                                $overwritten++
                                if (-not $SummaryOnly) { Log ("  OVER  {0}" -f $name) }
                            }
                        }
                        catch { $failed++; Log ("  FAIL  {0}: {1}" -f $name, $_.Exception.Message) }
                        break
                    }

                    'Rename' {
                        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
                        $ext  = [System.IO.Path]::GetExtension($name)
                        $i = 1; $cand = $null
                        while ($true) {
                            $cand = ('{0}-{1}{2}' -f $base, $i, $ext)
                            if (-not $index.Contains($cand)) { break }
                            $i++
                        }
                        try {
                            if ($PSCmdlet.ShouldProcess($path, "Move as $cand")) {
                                [System.IO.File]::Move($path, (Join-Path $dst $cand))
                                [void]$index.Add($cand)
                                $renamed++
                                if (-not $SummaryOnly) { Log ("  REN   {0} -> {1}" -f $name, $cand) }
                            }
                        }
                        catch { $failed++; Log ("  FAIL  {0}: {1}" -f $name, $_.Exception.Message) }
                        break
                    }

                    default {
                        $skipped++
                        if ($listSkipped.Count -lt $MaxReport) { [void]$listSkipped.Add($name) }
                        if (-not $SummaryOnly) { Log ("  SKIP  {0}  (name taken)" -f $name) }
                        break
                    }
                }
            }
        }

        if ($ProgressSeconds -gt 0 -and $swAct.Elapsed.TotalSeconds -ge $nextTick) {
            $nextTick = $swAct.Elapsed.TotalSeconds + $ProgressSeconds
            $rate = if ($swAct.Elapsed.TotalSeconds -gt 0) { $seenFiles / $swAct.Elapsed.TotalSeconds } else { 0 }
            Log ("  ... seen={0:n0} moved={1:n0} skipped={2:n0} renamed={3:n0} dupes={4:n0}  {5:n0}/s" -f `
                 $seenFiles, $moved, $skipped, $renamed, ($identicalDeleted + $identicalKept), $rate)
        }
    }
    $swAct.Stop()

    # --- empty directories ---------------------------------------------------
    $dirsRemoved = 0
    if ($RemoveEmptyDirs) {
        # Deepest first, so a directory emptied by its children is seen as empty.
        $dirs = @([System.IO.Directory]::EnumerateDirectories($src, '*', $enumOpt)) |
                Sort-Object -Property Length -Descending
        foreach ($d in $dirs) {
            try {
                if (-not [System.IO.Directory]::EnumerateFileSystemEntries($d).GetEnumerator().MoveNext()) {
                    if ($PSCmdlet.ShouldProcess($d, 'Remove empty directory')) {
                        [System.IO.Directory]::Delete($d)
                        $dirsRemoved++
                    }
                }
            }
            catch { Log ("  FAIL  rmdir {0}: {1}" -f $d, $_.Exception.Message) }
        }
    }

    Log "SUMMARY"
    Log ("  filesSeen={0:n0}" -f $seenFiles)
    Log ("  moved={0:n0}" -f $moved)
    Log ("  renamed={0:n0}" -f $renamed)
    Log ("  overwritten={0:n0}" -f $overwritten)
    Log ("  skippedNameTaken={0:n0}" -f $skipped)
    if ($CompareContent) {
        Log ("  identicalDeleted={0:n0}" -f $identicalDeleted)
        Log ("  identicalKept={0:n0}" -f $identicalKept)
    }
    Log ("  failed={0:n0}" -f $failed)
    if ($RemoveEmptyDirs) { Log ("  emptyDirsRemoved={0:n0}" -f $dirsRemoved) }
    Log ("  destinationNow={0:n0}" -f $index.Count)
    Log ("  elapsed index={0:n1}s move={1:n1}s" -f $swIdx.Elapsed.TotalSeconds, $swAct.Elapsed.TotalSeconds)

    if ($listSkipped.Count -gt 0) {
        Log "  skipped (name already at destination):"
        foreach ($x in $listSkipped) { Log ("    {0}" -f $x) }
        if ($skipped -gt $listSkipped.Count) { Log ("    ... +{0} more" -f ($skipped - $listSkipped.Count)) }
    }

    if ($skipped -gt 0 -and -not $CompareContent) {
        Log "  NOTE skipped files were matched by name only; -CompareContent tells an identical copy from a different file"
    }

    if ($failed -gt 0) { $exitCode = 1 }
    elseif ($skipped -gt 0) { $exitCode = 2 }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    if ($script:LogWriter) { $script:LogWriter.Flush(); $script:LogWriter.Dispose() }
    exit 1
}
finally {
    if ($script:LogWriter) { $script:LogWriter.Flush(); $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

exit $exitCode
