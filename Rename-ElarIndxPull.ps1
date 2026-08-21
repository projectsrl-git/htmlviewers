<#
.SYNOPSIS
    Renames ELAR INDX/PULL pairs by advancing the trailing time counter, and rewrites
    the PULL so that it points at the new INDX name.

.DESCRIPTION
    ELAR rejects a resubmission that reuses a file name already seen ("resend the
    files again with unique name"). The trailing segment of the name is a synthetic
    clock, so advancing it by one second yields a fresh, well-formed name:

        RZ2.ELA.FTP.CLICT@EV.D26232.INDX.C113839
                                        -> ...INDX.C113840

    The arithmetic is real time arithmetic, not a numeric increment: C113859 + 1 is
    C113900, not C113860.

    For each INDX the matching PULL is located by substituting PULL for INDX in the
    name. PullGenerator writes the index name into the PULL's Index attributes, so
    every occurrence of the old index name inside the PULL is replaced with the new
    one. The PULL file itself is renamed too, keeping the pair on the same counter.

    Nothing is renamed unless the PULL rewrite succeeded and actually replaced at
    least one occurrence: a PULL that does not mention its INDX means the pairing
    assumption is wrong, and the pair is skipped rather than half-renamed.

    Plain timestamped stdout, no prompts, no colours: intended to run as an
    OpenProteo exec step. Use -WhatIf for a dry run.

.PARAMETER Path
    Directory, file, or wildcard pattern pointing at the INDX files.

.PARAMETER Filter
    Wildcard applied when Path names a directory. Default *INDX*.

.PARAMETER Increment
    Seconds to add to the counter. Default 1.

.PARAMETER Encoding
    Charset used to read and rewrite the PULL. Defaults to windows-1252, matching
    what the legacy writer emitted.

.PARAMETER KeepPullName
    Rewrite the PULL contents but leave its file name unchanged. Off by default,
    since ELAR pairs the two by name.

.PARAMETER NoBackup
    Skip the .bak copy of the PULL before rewriting.

.PARAMETER MaxAttempts
    How many further seconds to try when the target name is already taken. Default 60.

.EXAMPLE
    .\Rename-ElarIndxPull.ps1 -Path "G:\ELAR\OUT\CMOD\S210967_CLICT" -WhatIf

.EXAMPLE
    .\Rename-ElarIndxPull.ps1 -Path "G:\ELAR\OUT\CMOD\S210967_CLICT\*INDX*"

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 one or more pairs skipped, 1 error.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [string] $Filter = '*INDX*',

    [int] $Increment = 1,

    [string] $Encoding = 'windows-1252',

    [switch] $KeepPullName,

    [switch] $NoBackup,

    [int] $MaxAttempts = 60
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

# <base>.INDX.C<hhmmss> with an optional .xml tail.
$reName = New-Object System.Text.RegularExpressions.Regex `
    '^(?<base>.+)\.(?<kind>INDX)\.(?<c>[A-Za-z]?)(?<hh>\d{2})(?<mm>\d{2})(?<ss>\d{2})(?<tail>\.[A-Za-z0-9]+)?$', `
    ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

function Get-ShiftedName {
    param([System.Text.RegularExpressions.Match] $M, [int] $Seconds)

    $h = [int]$M.Groups['hh'].Value
    $m = [int]$M.Groups['mm'].Value
    $s = [int]$M.Groups['ss'].Value

    if ($h -gt 23 -or $m -gt 59 -or $s -gt 59) {
        throw ("counter {0}{1}{2}{3} is not a valid time" -f $M.Groups['c'].Value, $M.Groups['hh'].Value, $M.Groups['mm'].Value, $M.Groups['ss'].Value)
    }

    $t = (New-TimeSpan -Hours $h -Minutes $m -Seconds $s).Add([TimeSpan]::FromSeconds($Seconds))
    $wrapped = ($t.TotalSeconds -ge 86400 -or $t.TotalSeconds -lt 0)
    $t = [TimeSpan]::FromSeconds((($t.TotalSeconds % 86400) + 86400) % 86400)

    return [pscustomobject]@{
        Counter = ('{0}{1:00}{2:00}{3:00}' -f $M.Groups['c'].Value, $t.Hours, $t.Minutes, $t.Seconds)
        Wrapped = $wrapped
    }
}

try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

$exitCode = 0

try {
    # @() matters: an unrolled single result would lose .Count under StrictMode.
    $files = @()
    foreach ($p in $Path) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            $files += @(Get-ChildItem -LiteralPath $p -File -Filter $Filter -ErrorAction SilentlyContinue)
        }
        elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $files += @(Get-Item -LiteralPath $p)
        }
        else {
            $files += @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue)
        }
    }
    $files = @($files | Where-Object { $_.Extension -notin @('.bak', '.tmp') } |
                        Sort-Object FullName -Unique)

    if ($files.Count -eq 0) {
        Log ("WARN no INDX file matched: {0}" -f ($Path -join ', '))
        Log  "END exit=0"
        exit 0
    }

    Log ("START files={0} increment={1}s charset={2} keepPullName={3}" -f `
         $files.Count, $Increment, $Encoding, $KeepPullName.IsPresent)

    $renamed = 0; $skipped = 0; $i = 0

    foreach ($f in $files) {
        $i++
        Log ("[{0}/{1}] {2}" -f $i, $files.Count, $f.Name)

        $m = $reName.Match($f.Name)
        if (-not $m.Success) {
            Log "    SKIP name does not match <base>.INDX.C<hhmmss>"
            $skipped++; $exitCode = 2; continue
        }

        $tail    = $m.Groups['tail'].Value
        $baseNm  = $m.Groups['base'].Value
        $oldIndx = $f.Name
        $oldPull = ($baseNm + '.PULL.' + $m.Groups['c'].Value + $m.Groups['hh'].Value + `
                    $m.Groups['mm'].Value + $m.Groups['ss'].Value + $tail)

        $pullPath = Join-Path $f.DirectoryName $oldPull
        if (-not (Test-Path -LiteralPath $pullPath)) {
            Log ("    SKIP matching PULL not found: {0}" -f $oldPull)
            $skipped++; $exitCode = 2; continue
        }

        # Find a free counter, stepping further if the target already exists.
        $attempt = 0; $newIndx = $null; $newPull = $null; $newCounter = $null
        while ($attempt -lt $MaxAttempts) {
            $attempt++
            $shift = Get-ShiftedName -M $m -Seconds ($Increment + $attempt - 1)
            if ($shift.Wrapped) {
                Log "    NOTE counter wrapped past midnight; the Julian day segment was NOT changed"
            }
            $newCounter = $shift.Counter
            $cIndx = $baseNm + '.INDX.' + $newCounter + $tail
            $cPull = $baseNm + '.PULL.' + $newCounter + $tail

            $freeIndx = -not (Test-Path -LiteralPath (Join-Path $f.DirectoryName $cIndx))
            $freePull = $KeepPullName -or -not (Test-Path -LiteralPath (Join-Path $f.DirectoryName $cPull))
            if ($freeIndx -and $freePull) { $newIndx = $cIndx; $newPull = $cPull; break }
        }

        if (-not $newIndx) {
            Log ("    SKIP no free counter within {0} attempt(s)" -f $MaxAttempts)
            $skipped++; $exitCode = 2; continue
        }
        if ($attempt -gt 1) {
            Log ("    NOTE target was taken; advanced {0}s instead of {1}s" -f ($Increment + $attempt - 1), $Increment)
        }

        # The PULL carries the index name WITHOUT the .xml tail, because
        # PullGenerator strips it before substituting [INDEX_NAME].
        $oldIndexRef = $baseNm + '.INDX.' + $m.Groups['c'].Value + $m.Groups['hh'].Value + `
                       $m.Groups['mm'].Value + $m.Groups['ss'].Value
        $newIndexRef = $baseNm + '.INDX.' + $newCounter

        $text = [System.IO.File]::ReadAllText($pullPath, $enc)
        $hits = ([regex]::Matches($text, [regex]::Escape($oldIndexRef))).Count

        if ($hits -eq 0) {
            Log ("    SKIP PULL does not mention {0}; pairing assumption does not hold" -f $oldIndexRef)
            $skipped++; $exitCode = 2; continue
        }

        $newText = $text.Replace($oldIndexRef, $newIndexRef)

        # If the PULL is renamed too, any self-reference must follow.
        $selfHits = 0
        if (-not $KeepPullName) {
            $oldPullRef = $baseNm + '.PULL.' + $m.Groups['c'].Value + $m.Groups['hh'].Value + `
                          $m.Groups['mm'].Value + $m.Groups['ss'].Value
            $newPullRef = $baseNm + '.PULL.' + $newCounter
            $selfHits = ([regex]::Matches($newText, [regex]::Escape($oldPullRef))).Count
            if ($selfHits -gt 0) { $newText = $newText.Replace($oldPullRef, $newPullRef) }
        }

        Log ("    plan  {0} -> {1}" -f $oldIndx, $newIndx)
        if (-not $KeepPullName) { Log ("    plan  {0} -> {1}" -f $oldPull, $newPull) }
        Log ("    pull  indexRefs={0} selfRefs={1}" -f $hits, $selfHits)

        if (-not $PSCmdlet.ShouldProcess($f.FullName, "Rename pair to $newCounter and rewrite PULL")) {
            continue
        }

        # Order matters: rewrite the PULL to a temporary file first, so a failure
        # never leaves a renamed INDX beside a PULL still pointing at the old name.
        $tmp = Join-Path $f.DirectoryName ($oldPull + '.rename.tmp')
        [System.IO.File]::WriteAllText($tmp, $newText, $enc)

        if (-not $NoBackup) {
            Copy-Item -LiteralPath $pullPath -Destination ($pullPath + '.bak') -Force
            Log ("    backup {0}.bak" -f $oldPull)
        }

        Rename-Item -LiteralPath $f.FullName -NewName $newIndx
        $targetPull = if ($KeepPullName) { $pullPath } else { Join-Path $f.DirectoryName $newPull }
        Move-Item -LiteralPath $tmp -Destination $targetPull -Force
        if (-not $KeepPullName -and $targetPull -ne $pullPath) {
            Remove-Item -LiteralPath $pullPath -Force
        }

        Log ("    done  counter={0}" -f $newCounter)
        $renamed++
    }

    Log "SUMMARY"
    Log ("  pairsRenamed={0}" -f $renamed)
    Log ("  pairsSkipped={0}" -f $skipped)
    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
