<#
.SYNOPSIS
    Reports - and optionally fixes - INDX lines longer than the receiving system
    tolerates, by re-wrapping the Base64 payload only.

.DESCRIPTION
    ELAR truncates any line beyond roughly 30000 characters, horizontally. A document
    whose Base64 content sits on one long line therefore loses its closing
    </ELAR:Content> tag, and the file is rejected with the content unterminated.

    The legacy builder wrapped the serialized XML every max.line.length characters
    precisely to stay under that limit. When the property is absent or set too high,
    the payload lands on a single line and the file breaks.

    This script measures every line and, with -Fix, re-wraps only the character data
    of the payload element:

      - inside the payload   the text is Base64, where whitespace is ignored by any
                             decoder, so breaking it is safe. Chunks are a multiple
                             of 4 characters, keeping whole Base64 quads.
      - anywhere else        nothing is touched. Breaking markup would invalidate the
                             document and breaking a metadata value would corrupt it,
                             which is exactly the defect the legacy wrapper caused.

    Over-length lines outside the payload are reported and left alone: they need a
    data or configuration fix, not a formatting one.

    Plain timestamped stdout, no prompts, no colours. Use -WhatIf for a dry run.

.PARAMETER Path
    Files, directories, or wildcard patterns.

.PARAMETER Filter
    Wildcard applied when Path names a directory. Default *INDX*.

.PARAMETER MaxLength
    Maximum line length. Default 25000, which is the agreed target and leaves margin
    under the 30000 the receiver enforces.

.PARAMETER Fix
    Re-wrap the payload in place.

.PARAMETER Truncate
    In addition to wrapping, shorten over-length lines that contain no payload by
    truncating their character data until the line fits.

    Only the text inside elements is cut; the markup around it is preserved, so the
    document stays well-formed. Truncating the raw line instead would drop the
    closing tag and break the file, which is the failure this whole exercise started
    from.

    This is data loss and it is deliberate: the shortened value is what gets
    archived. Prefer fixing the source data where that is possible.

.PARAMETER TruncatePayload
    Allow truncation of payload lines as well. Off by default and rarely correct: a
    cut Base64 stream decodes to an incomplete document. Wrapping achieves the same
    line length with no loss, so use this only when a payload genuinely must be
    discarded.

.PARAMETER TruncateMarker
    Text appended where a value was cut, e.g. "..." . Empty by default. It counts
    against the line budget and, being archived content, makes the truncation
    visible to whoever reads the record later.

.PARAMETER NoBackup
    Skip the .bak copy when fixing.

.PARAMETER ContentElement
    Local name of the payload element. Default Content.

.PARAMETER Encoding
    Charset used to read and rewrite. Defaults to windows-1252.

.PARAMETER ProgressSeconds
    Interval between in-file progress lines. Default 10. Set 0 to disable.

.PARAMETER MaxReport
    Maximum number of over-length lines listed per file. Default 20.

.EXAMPLE
    .\Test-ElarLineLength.ps1 -Path "G:\ELAR\OUT\CMOD\S210969_CLIAC\*INDX*"

.EXAMPLE
    .\Test-ElarLineLength.ps1 -Path "...INDX.C113912" -Fix

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 all lines within the limit, 2 over-length lines remain, 1 error.

    Memory note: a line is read whole. A file whose entire payload is one 200 MB line
    costs about twice that in memory while it is held. The script reports the largest
    line it saw so the situation is visible.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [string] $Filter = '*INDX*',

    [int] $MaxLength = 25000,

    [switch] $Fix,

    [switch] $Truncate,

    [switch] $TruncatePayload,

    [string] $TruncateMarker = '',

    [switch] $NoBackup,

    [string] $ContentElement = 'Content',

    [string] $Encoding = 'windows-1252',

    [int] $ProgressSeconds = 10,

    [int] $MaxReport = 20
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

if ($MaxLength -lt 100) { throw "-MaxLength must be at least 100." }

try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

# Base64 decodes in groups of four, so chunks must be a multiple of 4 to keep
# every line a whole number of quads.
$chunk = [int]([Math]::Floor($MaxLength / 4) * 4)

$esc = [regex]::Escape($ContentElement)
$reOpen  = New-Object System.Text.RegularExpressions.Regex ('<(?:[A-Za-z0-9_.\-]+:)?' + $esc + '(?:\s[^>]*)?>')
$reClose = New-Object System.Text.RegularExpressions.Regex ('</(?:[A-Za-z0-9_.\-]+:)?' + $esc + '\s*>')
$reSelf  = New-Object System.Text.RegularExpressions.Regex ('<(?:[A-Za-z0-9_.\-]+:)?' + $esc + '(?:\s[^>]*?)?/>')

# Splits a run of XML into alternating markup and character-data pieces, so that
# only the latter can be shortened.
function Split-Markup {
    param([string] $Text)

    $out = New-Object 'System.Collections.Generic.List[psobject]'
    $i = 0
    while ($i -lt $Text.Length) {
        $lt = $Text.IndexOf('<', $i)
        if ($lt -lt 0) {
            $out.Add([pscustomobject]@{ Text = $Text.Substring($i); IsText = $true })
            break
        }
        if ($lt -gt $i) {
            $out.Add([pscustomobject]@{ Text = $Text.Substring($i, $lt - $i); IsText = $true })
        }
        $gt = $Text.IndexOf('>', $lt)
        if ($gt -lt 0) {
            # Unterminated tag: leave it alone, it is not ours to shorten.
            $out.Add([pscustomobject]@{ Text = $Text.Substring($lt); IsText = $false })
            break
        }
        $out.Add([pscustomobject]@{ Text = $Text.Substring($lt, $gt - $lt + 1); IsText = $false })
        $i = $gt + 1
    }
    return $out
}

# Shortens character data until the whole line fits, longest value first, leaving
# every tag intact. Returns the rebuilt line and how many characters were dropped.
function Compress-Line {
    param([string] $Line, [int] $Limit, [string] $Marker)

    $pieces = Split-Markup -Text $Line
    $markupLen = 0
    foreach ($p in $pieces) { if (-not $p.IsText) { $markupLen += $p.Text.Length } }

    $budget = $Limit - $markupLen
    if ($budget -le 0) {
        # The tags alone exceed the limit: nothing can be cut safely.
        return [pscustomobject]@{ Line = $Line; Dropped = 0; Possible = $false }
    }

    $textLen = 0
    foreach ($p in $pieces) { if ($p.IsText) { $textLen += $p.Text.Length } }
    if ($textLen -le $budget) {
        return [pscustomobject]@{ Line = $Line; Dropped = 0; Possible = $true }
    }

    $excess = $textLen - $budget
    $dropped = 0

    # Greedy: take from the longest value first, so short fields such as dates and
    # codes survive untouched and only the genuinely oversized one is cut.
    while ($excess -gt 0) {
        $target = $null
        foreach ($p in $pieces) {
            if ($p.IsText -and ($null -eq $target -or $p.Text.Length -gt $target.Text.Length)) { $target = $p }
        }
        if ($null -eq $target -or $target.Text.Length -eq 0) { break }

        $cut = [Math]::Min($excess, $target.Text.Length)
        $keep = $target.Text.Length - $cut
        $newText = $target.Text.Substring(0, $keep)
        if ($Marker -and $keep -gt $Marker.Length) {
            $newText = $newText.Substring(0, $keep - $Marker.Length) + $Marker
        }
        $target.Text = $newText
        $dropped += $cut
        $excess  -= $cut
    }

    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $pieces) { [void]$sb.Append($p.Text) }
    return [pscustomobject]@{ Line = $sb.ToString(); Dropped = $dropped; Possible = $true }
}

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
    Log ("START mode={0} files={1} maxLength={2} chunk={3} payload={4} charset={5}" -f `
         $mode, $files.Count, $MaxLength, $chunk, $ContentElement, $Encoding)

    $i = 0; $filesChanged = 0; $totalOverOutside = 0
    $totalTruncated = 0; $totalDropped = 0; $totalUnfixable = 0

    foreach ($f in $files) {
        $i++
        Log ("[{0}/{1}] {2}  {3:n1} MB" -f $i, $files.Count, $f.Name, ($f.Length / 1MB))

        $tmp = if ($Fix) { Join-Path $f.DirectoryName ($f.Name + '.wrap.tmp') } else { $null }
        $reader = $null; $writer = $null; $fs = $null

        $lineNo = 0; $maxSeen = 0; $maxSeenLine = 0
        $overPayload = 0; $overOutside = 0; $wrapped = 0
        # Measured on what is written, so the fix proves itself rather than
        # requiring a second pass to confirm the limit was respected.
        $script:outMax = 0; $script:outOver = 0
        $truncatedLines = 0; $droppedChars = 0; $unfixable = 0
        $overLines = New-Object 'System.Collections.Generic.List[string]'
        $inContent = $false

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
                        Log ("    ... {0:n0}/{1:n0} MB ({2:n0}%)  {3:n0} lines  longest={4:n0}" -f `
                             ($pos / 1MB), ($total / 1MB), $pct, $lineNo, $maxSeen)
                    }
                }

                if ($line.Length -gt $maxSeen) { $maxSeen = $line.Length; $maxSeenLine = $lineNo }

                # Work out which parts of this line are payload character data.
                # Segments are emitted in order; only payload ones may be wrapped.
                $segs = @()   # each: @{ Text = ...; Payload = $true/$false }
                $pos0 = 0
                $stateIn = $inContent

                while ($pos0 -lt $line.Length) {
                    if ($stateIn) {
                        $mc = $reClose.Match($line, $pos0)
                        if ($mc.Success) {
                            if ($mc.Index -gt $pos0) { $segs += @{ Text = $line.Substring($pos0, $mc.Index - $pos0); Payload = $true } }
                            $segs += @{ Text = $mc.Value; Payload = $false }
                            $pos0 = $mc.Index + $mc.Length
                            $stateIn = $false
                        }
                        else {
                            $segs += @{ Text = $line.Substring($pos0); Payload = $true }
                            $pos0 = $line.Length
                        }
                    }
                    else {
                        $mo = $reOpen.Match($line, $pos0)
                        $ms = $reSelf.Match($line, $pos0)
                        # A self-closing payload element opens nothing.
                        if ($ms.Success -and (-not $mo.Success -or $ms.Index -le $mo.Index)) {
                            $segs += @{ Text = $line.Substring($pos0, $ms.Index + $ms.Length - $pos0); Payload = $false }
                            $pos0 = $ms.Index + $ms.Length
                        }
                        elseif ($mo.Success) {
                            $segs += @{ Text = $line.Substring($pos0, $mo.Index + $mo.Length - $pos0); Payload = $false }
                            $pos0 = $mo.Index + $mo.Length
                            $stateIn = $true
                        }
                        else {
                            $segs += @{ Text = $line.Substring($pos0); Payload = $false }
                            $pos0 = $line.Length
                        }
                    }
                }
                $inContent = $stateIn

                if ($line.Length -gt $MaxLength) {
                    $hasPayload = @($segs | Where-Object { $_.Payload }).Count -gt 0
                    if ($hasPayload) { $overPayload++ }
                    else {
                        $overOutside++
                        if ($overLines.Count -lt $MaxReport) {
                            [void]$overLines.Add(('line {0} len {1}' -f $lineNo, $line.Length))
                        }
                    }
                }

                if (-not $Fix) { continue }

                $hasPayloadSeg = @($segs | Where-Object { $_.Payload }).Count -gt 0

                # No payload on this line: wrapping does not apply, so truncation is
                # the only lever - and only when asked for.
                if (-not $hasPayloadSeg) {
                    if ($Truncate -and $line.Length -gt $MaxLength) {
                        $res = Compress-Line -Line $line -Limit $MaxLength -Marker $TruncateMarker
                        if ($res.Possible -and $res.Dropped -gt 0) {
                            $truncatedLines++
                            $droppedChars += $res.Dropped
                            $line = $res.Line
                        }
                        elseif (-not $res.Possible) {
                            $unfixable++
                        }
                    }
                    $writer.Write($line); $writer.Write("`n")
                    continue
                }

                if ($TruncatePayload -and $line.Length -gt $MaxLength) {
                    # Explicitly requested: cut the payload rather than wrap it. The
                    # decoded document will be incomplete.
                    $truncatedLines++
                    $droppedChars += ($line.Length - $MaxLength)
                    $writer.Write($line.Substring(0, $MaxLength)); $writer.Write("`n")
                    continue
                }

                # Emit, wrapping payload segments and leaving everything else intact.
                $col = 0
                $closeOut = {
                    if ($col -gt $script:outMax) { $script:outMax = $col }
                    if ($col -gt $MaxLength) { $script:outOver++ }
                }
                foreach ($sg in $segs) {
                    if (-not $sg.Payload) {
                        $writer.Write($sg.Text)
                        $col += $sg.Text.Length
                        continue
                    }
                    $t = $sg.Text
                    $off = 0
                    while ($off -lt $t.Length) {
                        $room = $chunk - $col
                        if ($room -le 0) { & $closeOut; $writer.Write("`n"); $wrapped++; $col = 0; $room = $chunk }
                        # Keep the break on a quad boundary relative to this segment.
                        $take = [Math]::Min($room, $t.Length - $off)
                        if (($off + $take) -lt $t.Length) { $take = [int]([Math]::Floor($take / 4) * 4) }
                        if ($take -le 0) { & $closeOut; $writer.Write("`n"); $wrapped++; $col = 0; continue }
                        $writer.Write($t.Substring($off, $take))
                        $off += $take
                        $col += $take
                    }
                }
                & $closeOut
                $writer.Write("`n")
            }
        }
        finally {
            if ($writer) { $writer.Flush(); $writer.Dispose() }
            if ($reader) { $reader.Dispose() }
            elseif ($fs) { $fs.Dispose() }
        }
        $swOne.Stop()

        Log ("    scanned lines={0:n0} elapsed={1:n1}s longestLine={2:n0} at line {3:n0}" -f `
             $lineNo, $swOne.Elapsed.TotalSeconds, $maxSeen, $maxSeenLine)
        Log ("    overLimit payload={0} outsidePayload={1}" -f $overPayload, $overOutside)
        if ($Fix) {
            Log ("    output longestLine={0:n0} overLimit={1}" -f $script:outMax, $script:outOver)
            if ($script:outOver -gt 0) {
                Log ("    WARNING {0} written line(s) still exceed {1}; these are markup or metadata, not payload" -f $script:outOver, $MaxLength)
            }
        }
        if ($Fix -and ($truncatedLines -gt 0 -or $unfixable -gt 0)) {
            Log ("    truncatedLines={0} charsDropped={1:n0} unfixable={2}" -f `
                 $truncatedLines, $droppedChars, $unfixable)
        }

        if ($overLines.Count -gt 0) {
            Log ("    outside-payload offenders: {0}" -f ($overLines -join '; '))
        }

        $totalOverOutside += $overOutside
        $totalTruncated += $truncatedLines
        $totalDropped   += $droppedChars
        $totalUnfixable += $unfixable

        if (-not $Fix) {
            if ($maxSeen -gt $MaxLength) { $exitCode = 2 }
            if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force }
            continue
        }

        if ($wrapped -eq 0 -and $truncatedLines -eq 0) {
            Log "    nothing to change"
            if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force }
        }
        elseif ($PSCmdlet.ShouldProcess($f.FullName, "Re-wrap payload at $chunk and truncate $truncatedLines line(s)")) {
            if (-not $NoBackup) {
                Copy-Item -LiteralPath $f.FullName -Destination ($f.FullName + '.bak') -Force
                Log ("    backup={0}.bak" -f $f.Name)
            }
            Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
            $filesChanged++
            Log ("    rewrapped breaksAdded={0} truncated={1} charsDropped={2:n0}" -f `
                 $wrapped, $truncatedLines, $droppedChars)
        }
        elseif ($tmp -and (Test-Path -LiteralPath $tmp)) {
            Remove-Item -LiteralPath $tmp -Force
        }

        if ($overOutside -gt 0 -or $script:outOver -gt 0) { $exitCode = 2 }
    }

    Log "SUMMARY"
    Log ("  filesScanned={0}" -f $files.Count)
    if ($Fix) { Log ("  filesChanged={0}" -f $filesChanged) }
    Log ("  overLimitOutsidePayload={0}" -f $totalOverOutside)
    if ($Fix -and $Truncate) {
        Log ("  truncatedLines={0}" -f $totalTruncated)
        Log ("  charactersDropped={0:n0}" -f $totalDropped)
        Log ("  unfixableLines={0}" -f $totalUnfixable)
    }

    if ($totalOverOutside -gt 0 -and -not $Truncate) {
        Log "  NOTE these lines are markup or metadata, not payload, and were left untouched"
        Log "  NOTE add -Truncate to shorten their character data, keeping the tags intact"
    }
    if ($totalUnfixable -gt 0) {
        Log "  NOTE some lines exceed the limit on markup alone; no truncation can fix those"
    }
    if ($totalDropped -gt 0) {
        Log "  NOTE truncated values are archived in their shortened form"
    }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
