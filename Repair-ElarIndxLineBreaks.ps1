<#
.SYNOPSIS
    Finds - and optionally repairs - metadata values corrupted by the legacy
    fixed-offset line wrapping in delivered ELAR INDX files.

.DESCRIPTION
    IndxBuilder.createIndx writes the serialized XML in fixed-length slices counted
    from the start of the file, appending a newline after each slice regardless of
    what sits at that offset. Three outcomes:

      - break inside the Base64 payload of ELAR:Content
            harmless, Base64 decoders ignore whitespace
      - break inside element character data
            the metadata value silently gains a line feed ("PDF" -> "P<LF>DF")
            NOTE: this is still well-formed XML. A well-formedness check does not
            detect it. The document parses; the value is simply wrong.
      - break inside a tag or an attribute name
            the file is not well-formed XML at all

    Corruption can only occur AT a line ending, and an INDX wrapped at 20000-25000
    characters has a few thousand lines even at 150 MB, so the scanner reads lines
    rather than characters. A line ending in '>' ends between elements and is
    classified with no further analysis; that fast path covers the large majority.

    Repair joins only the destructive breaks. Indentation and Base64 wrapping are
    written back exactly as delivered.

    Designed to run unattended as an OpenProteo exec step: no prompts, no colours,
    no progress bar. Every line goes to stdout with a timestamp and is flushed
    immediately, so the workflow log shows progress while the step is still running.

.PARAMETER Path
    Files, directories, or wildcard patterns. Wildcards are supported in both the
    directory and the file part, e.g. "G:\ELAR\OUT\*\*.INDX.*". Directories are
    expanded with -Filter.

.PARAMETER Filter
    Wildcard applied when Path names a directory. Default *INDX*.

.PARAMETER Recurse
    Recurse into subdirectories when Path names a directory.

.PARAMETER Fix
    Repair the files in place. No confirmation is asked. Use -WhatIf for a dry run.

.PARAMETER NoBackup
    Skip the .bak copy when repairing. Not recommended.

.PARAMETER Encoding
    Charset used to read and rewrite. Defaults to windows-1252, which is what the
    legacy writer actually emitted; the XML declaration claims ISO-8859-1 and the two
    differ only over 0x80-0x9F. The declaration is left untouched.

.PARAMETER ContentElement
    Local name of the element holding the Base64 payload. Default Content.

.PARAMETER ProgressSeconds
    Interval between in-file progress lines. Default 10. Set 0 to disable.

.PARAMETER ShowValues
    Include offending values in the output. Values may contain personal data, so
    they are omitted by default.

.PARAMETER CsvOut
    Optional CSV report path.

.PARAMETER FailOnFindings
    Exit with code 2 when findings exist, so a workflow can branch on the result.

.EXAMPLE
    .\Repair-ElarIndxLineBreaks.ps1 -Path "G:\ELAR\OUT\*\*.INDX.*"

.EXAMPLE
    .\Repair-ElarIndxLineBreaks.ps1 -Path "G:\ELAR\OUT\CMOD\S210969_CLIAC" -Fix

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 findings with -FailOnFindings, 1 error.
    With -Fix the repaired file is built alongside the original and swapped in only
    after a clean close, so an interrupted run leaves the original intact.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [string] $Filter = '*INDX*',

    [switch] $Recurse,

    [switch] $Fix,

    [switch] $NoBackup,

    [string] $Encoding = 'windows-1252',

    [string] $ContentElement = 'Content',

    [int] $ProgressSeconds = 10,

    [switch] $ShowValues,

    [string] $CsvOut,

    [switch] $FailOnFindings
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

function Resolve-IndxFiles {
    param([string[]] $Spec, [string] $Filter, [bool] $Recurse)

    $seen  = New-Object 'System.Collections.Generic.HashSet[string]'
    $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'

    foreach ($s in $Spec) {
        $cand = @()

        if (Test-Path -LiteralPath $s -PathType Container) {
            $g = @{ LiteralPath = $s; File = $true; Filter = $Filter }
            if ($Recurse) { $g['Recurse'] = $true }
            $cand = @(Get-ChildItem @g -ErrorAction SilentlyContinue)
        }
        elseif (Test-Path -LiteralPath $s -PathType Leaf) {
            $cand = @(Get-Item -LiteralPath $s)
        }
        else {
            $g = @{ Path = $s; File = $true }
            if ($Recurse) { $g['Recurse'] = $true }
            $cand = @(Get-ChildItem @g -ErrorAction SilentlyContinue)

            foreach ($d in @(Get-Item -Path $s -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer })) {
                $g2 = @{ LiteralPath = $d.FullName; File = $true; Filter = $Filter }
                if ($Recurse) { $g2['Recurse'] = $true }
                $cand += Get-ChildItem @g2 -ErrorAction SilentlyContinue
            }
        }

        foreach ($c in $cand) {
            if (-not $c -or $c.PSIsContainer) { continue }
            # This script's own name contains "Indx" and would match *INDX*.
            if ($c.Extension -in @('.ps1', '.psm1', '.bak', '.tmp')) { continue }
            if ($PSCommandPath -and $c.FullName -eq $PSCommandPath) { continue }
            if ($seen.Add($c.FullName)) { $files.Add($c) }
        }
    }
    return $files
}

function Test-NameChar {
    param([char] $c)
    return ([char]::IsLetterOrDigit($c) -or $c -eq '_' -or $c -eq '-' -or $c -eq '.' -or $c -eq ':')
}

function Invoke-IndxScan {
    param(
        [System.IO.FileInfo]   $File,
        [System.Text.Encoding] $Enc,
        [string]               $ContentLocalName,
        [bool]                 $DoFix,
        [string]               $OutPath,
        [int]                  $TickSeconds
    )

    $findings = New-Object 'System.Collections.Generic.List[psobject]'
    $reader = $null; $writer = $null; $fs = $null
    $markupBreaks = 0; $valueBreaks = 0; $docCount = 0; $lineNo = 0; $fastPath = 0
    $badAngles = 0
    $reBadAngle = New-Object System.Text.RegularExpressions.Regex '<[ \t]+(?=[A-Za-z_])'

    try {
        $fs = New-Object System.IO.FileStream($File.FullName,
                  [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite, 1048576)
        $reader = New-Object System.IO.StreamReader($fs, $Enc, $false, 1048576)
        if ($DoFix) { $writer = New-Object System.IO.StreamWriter($OutPath, $false, $Enc, 1048576) }

        $total    = [double]$File.Length
        $swFile   = [System.Diagnostics.Stopwatch]::StartNew()
        $nextTick = if ($TickSeconds -gt 0) { $TickSeconds } else { [double]::MaxValue }
        $checkAt  = 2000

        $prevLine   = $null
        $prevSafe   = $true
        $prevInTag  = $false
        $prevTail   = ''
        $curElement = ''

        while ($null -ne ($line = $reader.ReadLine())) {
            $lineNo++

            if ($lineNo -ge $checkAt) {
                $checkAt = $lineNo + 2000
                $el = $swFile.Elapsed.TotalSeconds
                if ($el -ge $nextTick) {
                    $nextTick = $el + $TickSeconds
                    $pos  = [double]$fs.Position
                    $pct  = if ($total -gt 0) { [Math]::Min(100, ($pos / $total) * 100) } else { 0 }
                    $mbps = if ($el -gt 0) { ($pos / 1MB) / $el } else { 0 }
                    $eta  = if ($mbps -gt 0) { [TimeSpan]::FromSeconds((($total - $pos) / 1MB) / $mbps) } else { [TimeSpan]::Zero }
                    Log ("    ... {0:n0}/{1:n0} MB ({2:n0}%)  {3:n0} lines  {4:n1} MB/s  ETA {5:mm\:ss}" -f `
                         ($pos / 1MB), ($total / 1MB), $pct, $lineNo, $mbps, $eta)
                }
            }

            if ($null -ne $prevLine) {
                if ($prevSafe) {
                    if ($DoFix) { $writer.Write($prevLine); $writer.Write("`n") }
                }
                else {
                    if ($prevInTag) { $markupBreaks++; $kind = 'MarkupLineBreak' }
                    else            { $valueBreaks++;  $kind = 'TextLineBreak' }

                    $head = if ($prevInTag) { $prevLine.Substring([Math]::Max(0, $prevLine.Length - 40)) } else { $prevTail }
                    $findings.Add([pscustomobject]@{
                        File    = $File.Name
                        Kind    = $kind
                        Element = $curElement
                        Line    = ($lineNo - 1)
                        Value   = $head + '<LF>' + $line.Substring(0, [Math]::Min(40, $line.Length))
                    })

                    if ($DoFix) {
                        $writer.Write($prevLine)
                        if ($prevInTag) {
                            $a = if ($prevLine.Length -gt 0) { $prevLine[$prevLine.Length - 1] } else { [char]0 }
                            $b = if ($line.Length -gt 0) { $line[0] } else { [char]0 }
                            # A space belongs here ONLY where the break separated two
                            # attributes, i.e. right after a closing quote. Anywhere
                            # else - and in particular straight after '<' - the break
                            # split a token and the halves must be rejoined.
                            # Inserting a space after '<' produces "< ELAR:TaxCode>",
                            # which is not a valid element start.
                            if (($a -eq '"' -or $a -eq "'") -and (Test-NameChar $b)) { $writer.Write(' ') }
                        }
                    }
                }
            }

            # Independent of line breaks: '<' followed by whitespace is an invalid
            # element start. Earlier versions of this script could introduce it while
            # repairing a break that fell immediately after '<'.
            if ($line.IndexOf('<') -ge 0) {
                $mm = $reBadAngle.Matches($line)
                if ($mm.Count -gt 0) {
                    $badAngles += $mm.Count
                    foreach ($mx in $mm) {
                        $findings.Add([pscustomobject]@{
                            File    = $File.Name
                            Kind    = 'InvalidSpaceAfterAngle'
                            Element = ''
                            Line    = $lineNo
                            Value   = $line.Substring($mx.Index, [Math]::Min(40, $line.Length - $mx.Index))
                        })
                    }
                    if ($DoFix) { $line = $reBadAngle.Replace($line, '<') }
                }
            }

            # Fast path: a line ending in '>' ends between elements.
            if ($line.EndsWith('>')) {
                $fastPath++
                $prevSafe = $true; $prevInTag = $false; $prevTail = ''

                $lt = $line.LastIndexOf('<')
                if ($lt -ge 0 -and $lt + 1 -lt $line.Length) {
                    $c = $line[$lt + 1]
                    if ($c -eq '/' -or $line[$line.Length - 2] -eq '/') { $curElement = '' }
                    elseif ($c -ne '?' -and $c -ne '!') {
                        $j = $lt + 1
                        while ($j -lt $line.Length -and (Test-NameChar $line[$j])) { $j++ }
                        $curElement = $line.Substring($lt + 1, $j - $lt - 1)
                        if ($line.IndexOf('<ELAR:Doc') -ge 0) { $docCount++ }
                    }
                }
            }
            else {
                $lastGt = $line.LastIndexOf('>')
                $lastLt = $line.LastIndexOf('<')

                if ($lastLt -gt $lastGt) {
                    $prevInTag = $true; $prevSafe = $false; $prevTail = ''
                }
                else {
                    $prevInTag = $false
                    $tail = if ($lastGt -ge 0) { $line.Substring($lastGt + 1) } else { $line }
                    $prevTail = $tail

                    if ([string]::IsNullOrWhiteSpace($tail)) {
                        $prevSafe = $true
                    }
                    else {
                        if ($lastLt -ge 0 -and $lastLt + 1 -lt $line.Length) {
                            $c = $line[$lastLt + 1]
                            if ($c -eq '/') { $curElement = '' }
                            elseif ($c -ne '?' -and $c -ne '!') {
                                $j = $lastLt + 1
                                while ($j -lt $line.Length -and (Test-NameChar $line[$j])) { $j++ }
                                $curElement = $line.Substring($lastLt + 1, $j - $lastLt - 1)
                            }
                        }
                        $local = if ($curElement) { $curElement.Substring($curElement.IndexOf(':') + 1) } else { '' }
                        $prevSafe = ($local -eq $ContentLocalName)
                    }
                }
            }

            $prevLine = $line
        }

        if ($DoFix -and $null -ne $prevLine) { $writer.Write($prevLine); $writer.Write("`n") }
    }
    finally {
        if ($writer) { $writer.Flush(); $writer.Dispose() }
        if ($reader) { $reader.Dispose() }
        elseif ($fs) { $fs.Dispose() }
    }

    return [pscustomobject]@{
        File         = $File
        Findings     = $findings
        MarkupBreaks = $markupBreaks
        ValueBreaks  = $valueBreaks
        BadAngles    = $badAngles
        DocCount     = $docCount
        Lines        = $lineNo
        FastPath     = $fastPath
    }
}

# ---------------------------------------------------------------------------

$exitCode = 0

try {
    try   { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
    catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

    # @() is required: returning a List from a function unrolls it, so a single
    # match would come back as a bare object and an empty one as $null - both of
    # which lack .Count under Set-StrictMode.
    $files = @(Resolve-IndxFiles -Spec $Path -Filter $Filter -Recurse:$Recurse.IsPresent)
    if ($files.Count -eq 0) {
        Log ("WARN  no files matched: {0}" -f ($Path -join ', '))
        Log  "END exit=0"
        exit 0
    }

    $mode = if ($Fix) { 'REPAIR' } else { 'REPORT' }
    Log ("START mode={0} files={1} charset={2} payload={3}" -f $mode, $files.Count, $Encoding, $ContentElement)

    $all = New-Object 'System.Collections.Generic.List[psobject]'
    $totalDocs = 0; $filesChanged = 0; $i = 0
    $swAll = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($f in $files) {
        $i++
        Log ("[{0}/{1}] {2}  {3:n1} MB" -f $i, $files.Count, $f.Name, ($f.Length / 1MB))

        $tmp = if ($Fix) { Join-Path $f.DirectoryName ($f.Name + '.repair.tmp') } else { $null }
        $swOne = [System.Diagnostics.Stopwatch]::StartNew()

        $r = Invoke-IndxScan -File $f -Enc $enc -ContentLocalName $ContentElement `
                             -DoFix:$Fix.IsPresent -OutPath $tmp -TickSeconds $ProgressSeconds
        $swOne.Stop()

        $totalDocs += $r.DocCount
        foreach ($x in @($r.Findings)) { $all.Add($x) }
        $bad = $r.MarkupBreaks + $r.ValueBreaks + $r.BadAngles
        $fastPct = if ($r.Lines -gt 0) { ($r.FastPath / $r.Lines) * 100 } else { 0 }

        Log ("    scanned lines={0:n0} docs={1:n0} elapsed={2:n1}s fastpath={3:n0}%" -f `
             $r.Lines, $r.DocCount, $swOne.Elapsed.TotalSeconds, $fastPct)

        if ($bad -eq 0) {
            Log "    result=OK"
            if ($Fix -and $tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force }
            continue
        }

        $label = if ($r.MarkupBreaks -gt 0 -or $r.BadAngles -gt 0) { 'MALFORMED' } else { 'CORRUPTED' }
        Log ("    result={0} values={1} markup={2} badAngle={3}" -f `
             $label, $r.ValueBreaks, $r.MarkupBreaks, $r.BadAngles)

        if (-not $Fix) { continue }

        if ($PSCmdlet.ShouldProcess($f.FullName, "Repair $bad line break(s) in place")) {
            if (-not $NoBackup) {
                Copy-Item -LiteralPath $f.FullName -Destination ($f.FullName + '.bak') -Force
                Log ("    backup={0}.bak" -f $f.Name)
            }
            Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
            $filesChanged++
            Log ("    repaired={0}" -f $f.Name)
        }
        elseif ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force
        }
    }
    $swAll.Stop()

    $textN   = @($all | Where-Object { $_.Kind -eq 'TextLineBreak' }).Count
    $tagN    = @($all | Where-Object { $_.Kind -eq 'MarkupLineBreak' }).Count
    $angleN  = @($all | Where-Object { $_.Kind -eq 'InvalidSpaceAfterAngle' }).Count

    Log "SUMMARY"
    Log ("  filesScanned={0}" -f $files.Count)
    Log ("  documents={0}" -f $totalDocs)
    Log ("  corruptedValues={0}" -f $textN)
    Log ("  markupBreaks={0}" -f $tagN)
    Log ("  invalidSpaceAfterAngle={0}" -f $angleN)
    if ($Fix) { Log ("  filesRepaired={0}" -f $filesChanged) }
    Log ("  elapsed={0:n1}s" -f $swAll.Elapsed.TotalSeconds)

    if ($all.Count -gt 0) {
        foreach ($g in ($all | Group-Object Element | Sort-Object Count -Descending)) {
            Log ("  element {0}={1}" -f $g.Name, $g.Count)
        }

        if ($ShowValues) {
            Log "  values follow (may contain personal data)"
            foreach ($x in $all) {
                Log ("    {0}:{1} {2} = {3}" -f $x.File, $x.Line, $x.Element, $x.Value)
            }
        }

        if ($CsvOut) {
            $export = if ($ShowValues) { $all } else { $all | Select-Object File, Kind, Element, Line }
            $export | Export-Csv -LiteralPath $CsvOut -NoTypeInformation -Encoding UTF8
            Log ("  report={0}" -f $CsvOut)
        }

        if ($tagN -gt 0) {
            Log "  NOTE markup breaks mean those files were not well-formed XML"
        }
        if ($angleN -gt 0) {
            Log "  NOTE '< name' is an invalid element start; ELAR reports 'invalid start of an element'"
        }
        if ($textN -gt 0) {
            Log "  NOTE corrupted values still parse as valid XML; the values themselves are wrong"
        }
        if ($FailOnFindings) { $exitCode = 2 }
    }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
