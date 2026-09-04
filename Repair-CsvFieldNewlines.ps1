<#
.SYNOPSIS
    Removes CR/LF that fall inside a field of a CSV, rejoining each record onto a
    single line. The header's field count is the reference.

.DESCRIPTION
    A line break inside a value splits one logical record across several physical
    lines. Any reader that works line by line - the legacy ELAR tooling included -
    then sees two malformed records instead of one, and usually drops both without a
    word.

    This script reads physical lines and accumulates them until the record is
    complete, then writes it out as one line. A record is complete when both hold:

      - the quote state is balanced, so we are not inside a quoted field, and
      - the field count equals the header's

    Those two tests cover the two ways a break gets in. In a quoted file the break
    sits inside quotes and the quote state finds it. In an unquoted file - which is
    what the legacy exports are - nothing marks it, and the only evidence is that the
    line has fewer fields than the header declares. Hence the header as reference.

    Quoting is interpreted the lenient way: a quote opens a quoted field only at the
    start of a field. Elsewhere it is a literal character, which is what stray
    quotation marks in unquoted legacy data actually are.

    The rejoining is textual: the physical lines are concatenated and nothing else is
    touched, so quoting, spacing and the delimiter survive exactly as they were. Only
    the newline characters that were inside a record disappear.

    A record whose field count EXCEEDS the header is a different defect - an
    unescaped delimiter inside a value - and cannot be repaired by joining. It is
    written through untouched and reported, never silently altered.

    Use -WhatIf for a dry run. Plain timestamped stdout, no prompts.

.PARAMETER Path
    CSV to repair.

.PARAMETER OutFile
    Write to this file instead of replacing the input.

.PARAMETER JoinWith
    String inserted where the line break was. Empty by default, which is removal: the
    legacy corruption INSERTED a break into a value, so removing it restores the
    original. Use ' ' when the break replaced a space rather than being added.

.PARAMETER Delimiter
    CSV delimiter. Default ';'.

.PARAMETER Encoding
    Charset used to read and write. Default windows-1252.

.PARAMETER Eol
    Line ending of the output: Auto (default, follows the input), CRLF, or LF.

.PARAMETER MaxLinesPerRecord
    Safety limit on how many physical lines may be joined into one record. Default 50.
    Beyond it the accumulation is abandoned and the lines are written through
    unchanged, so a malformed file cannot make the script swallow the rest of itself.

.PARAMETER NoBackup
    Skip the .bak copy when replacing the input.

.PARAMETER MaxReport
    Maximum line numbers listed per category. Default 30.

.EXAMPLE
    .\Repair-CsvFieldNewlines.ps1 -Path .\feed.csv -WhatIf

.EXAMPLE
    .\Repair-CsvFieldNewlines.ps1 -Path .\feed.csv -OutFile .\feed-fixed.csv -JoinWith ' '

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 nothing to repair, 2 records were rejoined or anomalies found,
    1 error.

    Streaming: one record is held at a time, so file size is not a constraint.
    The output is built alongside the input and swapped in only after a clean close.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [string] $OutFile,

    [AllowEmptyString()]
    [string] $JoinWith = '',

    [string] $Delimiter = ';',

    [string] $Encoding = 'windows-1252',

    [ValidateSet('Auto', 'CRLF', 'LF')]
    [string] $Eol = 'Auto',

    [int] $MaxLinesPerRecord = 50,

    [switch] $NoBackup,

    [int] $MaxReport = 30
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

<#
    Counts the fields of a candidate record and reports whether a quoted field is
    still open at the end. A quote opens a quoted field only at the start of a field;
    anywhere else it is a literal character.
#>
function Measure-CsvFields {
    param([string] $Text, [char] $Sep)

    $fields = 1
    $inQ = $false
    $atFieldStart = $true

    for ($i = 0; $i -lt $Text.Length; $i++) {
        $c = $Text[$i]

        if ($inQ) {
            if ($c -eq '"') {
                if (($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq '"') { $i++ }   # escaped quote
                else { $inQ = $false }
            }
            continue
        }

        if ($atFieldStart -and $c -eq '"') { $inQ = $true; $atFieldStart = $false; continue }

        if ($c -eq $Sep) { $fields++; $atFieldStart = $true; continue }

        $atFieldStart = $false
    }

    return [pscustomobject]@{ Fields = $fields; InQuote = $inQ }
}

$exitCode = 0

try {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }
    if ($MaxLinesPerRecord -lt 2) { throw "-MaxLinesPerRecord must be at least 2." }

    try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
    catch { throw "Unknown charset '$Encoding'." }

    $src = Get-Item -LiteralPath $Path
    $sep = $Delimiter[0]
    $inPlace = -not $OutFile
    $dest = if ($inPlace) { Join-Path $src.DirectoryName ($src.Name + '.eolfix.tmp') } else { $OutFile }

    $outDir = Split-Path -Parent $dest
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    # --- detect the dominant line ending, so the output keeps the input's ---
    $nl = "`r`n"
    if ($Eol -eq 'LF')   { $nl = "`n" }
    elseif ($Eol -eq 'CRLF') { $nl = "`r`n" }
    else {
        $probe = New-Object byte[] 65536
        $fsP = [System.IO.File]::OpenRead($src.FullName)
        try { $nRead = $fsP.Read($probe, 0, $probe.Length) } finally { $fsP.Dispose() }
        $crlf = 0; $lf = 0
        for ($i = 0; $i -lt $nRead; $i++) {
            if ($probe[$i] -eq 10) {
                if ($i -gt 0 -and $probe[$i - 1] -eq 13) { $crlf++ } else { $lf++ }
            }
        }
        $nl = if ($lf -gt $crlf) { "`n" } else { "`r`n" }
    }

    Log ("START file={0} delimiter='{1}' charset={2} joinWith='{3}' eol={4}" -f `
         $src.Name, $Delimiter, $Encoding, $JoinWith, $(if ($nl -eq "`n") { 'LF' } else { 'CRLF' }))
    Log ("  target={0}" -f $(if ($inPlace) { 'in place' } else { $dest }))

    $reader = $null; $writer = $null
    $physical = 0; $records = 0; $rejoined = 0; $linesConsumed = 0
    $tooMany = 0; $abandoned = 0
    $listRejoin = New-Object 'System.Collections.Generic.List[string]'
    $listTooMany = New-Object 'System.Collections.Generic.List[string]'
    $expected = 0

    try {
        $reader = New-Object System.IO.StreamReader($src.FullName, $enc, $false, 1048576)
        $writer = New-Object System.IO.StreamWriter($dest, $false, $enc, 1048576)

        # --- header sets the reference ---
        $headerLine = $reader.ReadLine()
        if ($null -eq $headerLine) { throw "File is empty." }
        $physical++

        $h = Measure-CsvFields -Text $headerLine -Sep $sep
        if ($h.InQuote) { Log "  WARNING the header itself has an unterminated quote" }
        $expected = $h.Fields
        Log ("  headerFields={0}" -f $expected)

        $writer.Write($headerLine); $writer.Write($nl)

        # --- records ---
        $buf = $null
        $bufStart = 0
        $bufLines = 0

        while ($null -ne ($line = $reader.ReadLine())) {
            $physical++

            if ($null -eq $buf) {
                $buf = $line; $bufStart = $physical; $bufLines = 1
            }
            else {
                $buf = $buf + $JoinWith + $line
                $bufLines++
            }

            $m = Measure-CsvFields -Text $buf -Sep $sep

            if ($m.InQuote) {
                # Definitely incomplete: a quoted field is still open.
                if ($bufLines -ge $MaxLinesPerRecord) {
                    $abandoned++
                    $writer.Write($buf); $writer.Write($nl)
                    $records++
                    $buf = $null
                }
                continue
            }

            if ($m.Fields -lt $expected) {
                # Too few fields: the record continues on the next physical line.
                if ($bufLines -ge $MaxLinesPerRecord) {
                    $abandoned++
                    $writer.Write($buf); $writer.Write($nl)
                    $records++
                    $buf = $null
                }
                continue
            }

            if ($m.Fields -gt $expected) {
                # Not a line-break problem: an unescaped delimiter inside a value.
                # Joining cannot fix it, so it goes through untouched and is reported.
                $tooMany++
                if ($listTooMany.Count -lt $MaxReport) {
                    [void]$listTooMany.Add(('line {0}: {1} fields' -f $bufStart, $m.Fields))
                }
            }

            if ($bufLines -gt 1) {
                $rejoined++
                $linesConsumed += $bufLines
                if ($listRejoin.Count -lt $MaxReport) {
                    [void]$listRejoin.Add(('lines {0}-{1} ({2} physical)' -f $bufStart, $physical, $bufLines))
                }
            }

            $writer.Write($buf); $writer.Write($nl)
            $records++
            $buf = $null
        }

        # A trailing incomplete record is written rather than dropped: losing data
        # would be worse than leaving a row that still needs attention.
        if ($null -ne $buf) {
            $abandoned++
            $writer.Write($buf); $writer.Write($nl)
            $records++
        }
    }
    finally {
        if ($writer) { $writer.Flush(); $writer.Dispose() }
        if ($reader) { $reader.Dispose() }
    }

    Log "SUMMARY"
    Log ("  physicalLines={0:n0}" -f $physical)
    Log ("  records={0:n0}" -f $records)
    Log ("  recordsRejoined={0:n0}" -f $rejoined)
    Log ("  extraLinesRemoved={0:n0}" -f $(if ($rejoined -gt 0) { $linesConsumed - $rejoined } else { 0 }))
    Log ("  recordsWithTooManyFields={0:n0}" -f $tooMany)
    Log ("  recordsAbandonedIncomplete={0:n0}" -f $abandoned)

    if ($listRejoin.Count -gt 0) {
        Log "  rejoined:"
        foreach ($x in $listRejoin) { Log ("    {0}" -f $x) }
        if ($rejoined -gt $listRejoin.Count) { Log ("    ... +{0} more" -f ($rejoined - $listRejoin.Count)) }
    }
    if ($listTooMany.Count -gt 0) {
        Log "  too many fields (an unescaped delimiter, not a line break - not repaired):"
        foreach ($x in $listTooMany) { Log ("    {0}" -f $x) }
        if ($tooMany -gt $listTooMany.Count) { Log ("    ... +{0} more" -f ($tooMany - $listTooMany.Count)) }
    }
    if ($abandoned -gt 0) {
        Log ("  NOTE {0} record(s) never reached {1} fields within {2} lines and were written unchanged" -f `
             $abandoned, $expected, $MaxLinesPerRecord)
    }

    if ($inPlace) {
        if ($PSCmdlet.ShouldProcess($src.FullName, "Replace with $rejoined rejoined record(s)")) {
            if (-not $NoBackup) {
                Copy-Item -LiteralPath $src.FullName -Destination ($src.FullName + '.bak') -Force
                Log ("  backup={0}.bak" -f $src.Name)
            }
            Move-Item -LiteralPath $dest -Destination $src.FullName -Force
            Log ("  replaced={0}" -f $src.Name)
        }
        elseif (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force }
    }
    else {
        Log ("  out={0}" -f $dest)
    }

    if ($rejoined -gt 0 -or $tooMany -gt 0 -or $abandoned -gt 0) { $exitCode = 2 }
    else { Log "  every record already occupied exactly one line" }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
