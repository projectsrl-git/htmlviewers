<#
.SYNOPSIS
    Reports - and optionally fills - empty mandatory elements in ELAR INDX files,
    and reports elements whose occurrence count per document is not exactly one.

.DESCRIPTION
    ELAR rejects an INDX with

        "Definition of the Tag is Mandatory with Single occurrence where as it
         either occurs multiple times in the Document or is not present at all"

    naming an element and a record number. An element that carries no value is
    serialized by the legacy builder as an empty element - <ELAR:AccountID/> - which
    the receiving schema treats as absent. Filling it with a placeholder makes the
    document valid.

    Two distinct conditions are detected, and only one of them is repaired:

      - EMPTY      the element is present but carries no text. Repaired by writing
                   -Value into it.
      - MISSING    the element does not appear in the document at all. Not repaired:
                   inserting it would require knowing where the schema expects it in
                   the sequence, which this script does not know.
      - DUPLICATE  the element appears more than once. Not repaired: deciding which
                   occurrence to keep is a data question, not a formatting one.

    Both unrepaired conditions are reported per record number, matching the numbering
    ELAR uses, so they can be traced back to the source row.

    Scanning is line-based, which assumes each element sits on a single line. Run the
    line-break repair first if a file may still carry breaks inside markup.

    Plain timestamped stdout, no prompts, no colours: intended to run as an OpenProteo
    exec step. Use -WhatIf for a dry run.

.PARAMETER Path
    Files, directories, or wildcard patterns pointing at INDX files.

.PARAMETER Filter
    Wildcard applied when Path names a directory. Default *INDX*.

.PARAMETER Tag
    Elements to check, e.g. ELAR:AccountID. With a prefix the match is exact; without
    one ("AccountID") any prefix matches.

    Accepts a PowerShell array, a single comma- or semicolon-separated string, or a
    mixture: -Tag ELAR:AccountID,ELAR:CaseId and -Tag "ELAR:AccountID,ELAR:CaseId"
    are equivalent. The string form is what a workflow step passes as one parameter.

.PARAMETER Value
    Placeholder written into empty elements. Required with -Fix.

    Either one value applied to every tag, or a comma-separated list matching the
    tags one to one: -Tag "AccountID,CaseId" -Value "-,N/A" fills AccountID with "-"
    and CaseId with "N/A". A count mismatch is an error rather than a silent reuse.

    Remember this becomes archived data. "-" or "N/A" stay recognisable as a missing
    value; a plausible-looking number is indistinguishable from a real one.

.PARAMETER Fix
    Write the placeholder into empty elements, in place.

.PARAMETER NoBackup
    Skip the .bak copy when repairing.

.PARAMETER DocElement
    Local name of the element delimiting one document. Default Doc.

.PARAMETER Encoding
    Charset used to read and rewrite. Defaults to windows-1252.

.PARAMETER ProgressSeconds
    Interval between in-file progress lines. Default 10. Set 0 to disable.

.PARAMETER MaxReport
    Maximum number of affected records listed per file. Default 50.

.EXAMPLE
    .\Set-ElarEmptyTag.ps1 -Path "G:\ELAR\OUT\CMOD\S210969_CLIAC\*INDX*" -Tag ELAR:AccountID

.EXAMPLE
    .\Set-ElarEmptyTag.ps1 -Path "...INDX.C113912" -Tag ELAR:AccountID -Value "-" -Fix

.EXAMPLE
    .\Set-ElarEmptyTag.ps1 -Path "...INDX.C113912" `
        -Tag "ELAR:AccountID,ELAR:CaseId,ELAR:AltClientId" -Value "-" -Fix

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 unrepaired conditions remain, 1 error.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [Parameter(Mandatory = $true, Position = 1)]
    [string[]] $Tag,

    [string] $Value,

    [switch] $Fix,

    [switch] $NoBackup,

    [string] $Filter = '*INDX*',

    [string] $DocElement = 'Doc',

    [string] $Encoding = 'windows-1252',

    [int] $ProgressSeconds = 10,

    [int] $MaxReport = 50
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

# A workflow step passes one string, so accept "a,b;c" as well as an array.
$Tag = @($Tag | ForEach-Object { $_ -split '[,;]' } |
                ForEach-Object { $_.Trim() } |
                Where-Object  { $_ } )
if ($Tag.Count -eq 0) { throw "-Tag is empty." }
$Tag = @($Tag | Select-Object -Unique)

if ($Fix -and [string]::IsNullOrEmpty($Value)) { throw "-Fix requires -Value." }

# One value for every tag, or one value per tag in the same order.
$values = @()
if (-not [string]::IsNullOrEmpty($Value)) {
    $values = @($Value -split ',' | ForEach-Object { $_.Trim() })
    if ($values.Count -eq 1) {
        $values = @(1..$Tag.Count | ForEach-Object { $values[0] })
    }
    elseif ($values.Count -ne $Tag.Count) {
        throw ("-Value has {0} entries but -Tag has {1}. Give one value for all tags, or one per tag." -f `
               $values.Count, $Tag.Count)
    }
}

try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

function New-TagPattern {
    param([string] $Name, [string] $Placeholder)

    if ($Name.Contains(':')) { $n = [regex]::Escape($Name) }
    else { $n = '(?:[A-Za-z0-9_.\-]+:)?' + [regex]::Escape($Name) }

    $opts = [System.Text.RegularExpressions.RegexOptions]::None

    # XML-escape the placeholder: it becomes element content.
    $esc = $null
    if ($null -ne $Placeholder) {
        $esc = $Placeholder.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    }

    return [pscustomobject]@{
        Name    = $Name
        Escaped = $esc
        # Any occurrence, open or self-closed.
        Any     = New-Object System.Text.RegularExpressions.Regex ('<' + $n + '(?=[\s/>])'), $opts
        # Self-closed:  <X/>  or  <X attr="v"/>
        Self    = New-Object System.Text.RegularExpressions.Regex ('<(' + $n + ')((?:\s[^>]*?)?)\s*/>'), $opts
        # Open+close with nothing, or only whitespace, in between.
        Pair    = New-Object System.Text.RegularExpressions.Regex ('<(' + $n + ')((?:\s[^>]*?)?)>\s*</' + $n + '\s*>'), $opts
    }
}

$patterns = @()
for ($k = 0; $k -lt $Tag.Count; $k++) {
    $ph = if ($values.Count -gt 0) { $values[$k] } else { $null }
    $patterns += New-TagPattern -Name $Tag[$k] -Placeholder $ph
}

$reDoc = New-Object System.Text.RegularExpressions.Regex `
    ('<(?:[A-Za-z0-9_.\-]+:)?' + [regex]::Escape($DocElement) + '(?=[\s>])')

$exitCode = 0

try {
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
    $files = @($files | Where-Object { $_.Extension -notin @('.bak', '.tmp', '.ps1') } |
                        Sort-Object FullName -Unique)

    if ($files.Count -eq 0) {
        Log ("WARN no file matched: {0}" -f ($Path -join ', '))
        Log  "END exit=0"
        exit 0
    }

    $mode = if ($Fix) { 'FIX' } else { 'REPORT' }
    $pairsShown = if ($values.Count -gt 0) {
        (0..($Tag.Count - 1) | ForEach-Object { '{0}=''{1}''' -f $Tag[$_], $values[$_] }) -join ' '
    } else { ($Tag -join ',') }
    Log ("START mode={0} files={1} tags={2} doc={3} charset={4}" -f `
         $mode, $files.Count, $pairsShown, $DocElement, $Encoding)

    $i = 0; $filesChanged = 0; $totalFilled = 0; $totalMissing = 0; $totalDup = 0

    foreach ($f in $files) {
        $i++
        Log ("[{0}/{1}] {2}  {3:n1} MB" -f $i, $files.Count, $f.Name, ($f.Length / 1MB))

        $tmp = if ($Fix) { Join-Path $f.DirectoryName ($f.Name + '.tagfix.tmp') } else { $null }
        $reader = $null; $writer = $null; $fs = $null

        $record   = 0
        $lineNo   = 0
        $filled   = 0
        $emptyRecs   = New-Object 'System.Collections.Generic.List[int]'
        $missingRecs = New-Object 'System.Collections.Generic.List[int]'
        $dupRecs     = New-Object 'System.Collections.Generic.List[int]'
        # Per-document counters, one slot per requested tag.
        $seen  = New-Object int[] $patterns.Count
        $blank = New-Object int[] $patterns.Count
        $fillPerTag = New-Object int[] $patterns.Count
        # Per-tag record lists, so a run over several tags stays diagnosable.
        $emptyByTag   = @(); $missingByTag = @(); $dupByTag = @()
        for ($k = 0; $k -lt $patterns.Count; $k++) {
            $emptyByTag   += ,(New-Object 'System.Collections.Generic.List[int]')
            $missingByTag += ,(New-Object 'System.Collections.Generic.List[int]')
            $dupByTag     += ,(New-Object 'System.Collections.Generic.List[int]')
        }

        $swOne = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $fs = New-Object System.IO.FileStream($f.FullName,
                      [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
                      [System.IO.FileShare]::ReadWrite, 1048576)
            $reader = New-Object System.IO.StreamReader($fs, $enc, $false, 1048576)
            if ($Fix) { $writer = New-Object System.IO.StreamWriter($tmp, $false, $enc, 1048576) }

            $total    = [double]$f.Length
            $nextTick = if ($ProgressSeconds -gt 0) { [double]$ProgressSeconds } else { [double]::MaxValue }
            $checkAt  = 2000

            while ($null -ne ($line = $reader.ReadLine())) {
                $lineNo++

                if ($lineNo -ge $checkAt) {
                    $checkAt = $lineNo + 2000
                    $el = $swOne.Elapsed.TotalSeconds
                    if ($el -ge $nextTick) {
                        $nextTick = $el + $ProgressSeconds
                        $pos = [double]$fs.Position
                        $pct = if ($total -gt 0) { [Math]::Min(100, ($pos / $total) * 100) } else { 0 }
                        Log ("    ... {0:n0}/{1:n0} MB ({2:n0}%)  {3:n0} records" -f `
                             ($pos / 1MB), ($total / 1MB), $pct, $record)
                    }
                }

                # A document boundary closes the previous record's tally.
                if ($line.IndexOf('<') -ge 0) {
                    $docHits = $reDoc.Matches($line).Count
                    for ($d = 0; $d -lt $docHits; $d++) {
                        if ($record -gt 0) {
                            for ($k = 0; $k -lt $patterns.Count; $k++) {
                                if     ($seen[$k] -eq 0) { [void]$missingRecs.Add($record); [void]$missingByTag[$k].Add($record) }
                                elseif ($seen[$k] -gt 1) { [void]$dupRecs.Add($record);     [void]$dupByTag[$k].Add($record) }
                                if ($blank[$k] -gt 0)    { [void]$emptyRecs.Add($record);   [void]$emptyByTag[$k].Add($record) }
                            }
                        }
                        $record++
                        for ($k = 0; $k -lt $patterns.Count; $k++) { $seen[$k] = 0; $blank[$k] = 0 }
                    }

                    for ($k = 0; $k -lt $patterns.Count; $k++) {
                        $pt = $patterns[$k]
                        $n = $pt.Any.Matches($line).Count
                        if ($n -gt 0) {
                            $seen[$k] += $n
                            $b = $pt.Self.Matches($line).Count + $pt.Pair.Matches($line).Count
                            if ($b -gt 0) {
                                $blank[$k] += $b
                                if ($Fix) {
                                    # Preserve the original prefix and any attributes.
                                    $line = $pt.Self.Replace($line, ('<$1$2>' + $pt.Escaped + '</$1>'))
                                    $line = $pt.Pair.Replace($line, ('<$1$2>' + $pt.Escaped + '</$1>'))
                                    $filled += $b
                                    $fillPerTag[$k] += $b
                                }
                            }
                        }
                    }
                }

                if ($Fix) { $writer.Write($line); $writer.Write("`n") }
            }

            # Close the last record.
            if ($record -gt 0) {
                for ($k = 0; $k -lt $patterns.Count; $k++) {
                    if     ($seen[$k] -eq 0) { [void]$missingRecs.Add($record); [void]$missingByTag[$k].Add($record) }
                    elseif ($seen[$k] -gt 1) { [void]$dupRecs.Add($record);     [void]$dupByTag[$k].Add($record) }
                    if ($blank[$k] -gt 0)    { [void]$emptyRecs.Add($record);   [void]$emptyByTag[$k].Add($record) }
                }
            }
        }
        finally {
            if ($writer) { $writer.Flush(); $writer.Dispose() }
            if ($reader) { $reader.Dispose() }
            elseif ($fs) { $fs.Dispose() }
        }
        $swOne.Stop()

        Log ("    scanned records={0:n0} lines={1:n0} elapsed={2:n1}s" -f $record, $lineNo, $swOne.Elapsed.TotalSeconds)
        Log ("    empty={0} missing={1} duplicate={2}" -f $emptyRecs.Count, $missingRecs.Count, $dupRecs.Count)

        # Report per tag: with several tags a combined list says nothing useful.
        for ($k = 0; $k -lt $patterns.Count; $k++) {
            $tn = $patterns[$k].Name
            foreach ($pair in @(
                @{ N = 'empty';     L = $emptyByTag[$k] },
                @{ N = 'missing';   L = $missingByTag[$k] },
                @{ N = 'duplicate'; L = $dupByTag[$k] })) {

                if ($pair.L.Count -gt 0) {
                    $shown = @($pair.L | Select-Object -First $MaxReport) -join ','
                    $more  = if ($pair.L.Count -gt $MaxReport) { (' ... +{0} more' -f ($pair.L.Count - $MaxReport)) } else { '' }
                    Log ("    {0} {1}={2} records: {3}{4}" -f $tn, $pair.N, $pair.L.Count, $shown, $more)
                }
            }
            if ($Fix -and $fillPerTag[$k] -gt 0) {
                Log ("    {0} filled={1}" -f $tn, $fillPerTag[$k])
            }
        }

        $totalMissing += $missingRecs.Count
        $totalDup     += $dupRecs.Count

        if (-not $Fix) {
            if ($emptyRecs.Count -gt 0 -or $missingRecs.Count -gt 0 -or $dupRecs.Count -gt 0) { $exitCode = 2 }
            continue
        }

        if ($filled -eq 0) {
            Log "    no empty element to fill"
            if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force }
        }
        elseif ($PSCmdlet.ShouldProcess($f.FullName, "Fill $filled empty element(s) with '$Value'")) {
            if (-not $NoBackup) {
                Copy-Item -LiteralPath $f.FullName -Destination ($f.FullName + '.bak') -Force
                Log ("    backup={0}.bak" -f $f.Name)
            }
            Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
            $filesChanged++
            $totalFilled += $filled
            Log ("    filled={0}" -f $filled)
        }
        elseif ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force
        }

        # Filling does not address these two.
        if ($missingRecs.Count -gt 0 -or $dupRecs.Count -gt 0) { $exitCode = 2 }
    }

    Log "SUMMARY"
    Log ("  filesScanned={0}" -f $files.Count)
    if ($Fix) {
        Log ("  filesChanged={0}" -f $filesChanged)
        Log ("  elementsFilled={0}" -f $totalFilled)
    }
    Log ("  recordsMissingTag={0}" -f $totalMissing)
    Log ("  recordsDuplicateTag={0}" -f $totalDup)

    if ($totalMissing -gt 0) {
        Log "  NOTE a missing element is not inserted: its position in the schema sequence is unknown here"
    }
    if ($totalDup -gt 0) {
        Log "  NOTE duplicates are not removed: choosing which occurrence survives is a data decision"
    }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
