<#
.SYNOPSIS
    Extracts the Nth occurrence of an element from a large XML file.

.DESCRIPTION
    Streams the file and returns the character data of the requested element
    occurrence, without loading the document into memory. Written for ELAR INDX
    files, where a single file can be 150 MB and a single ELAR:Content element can
    hold a whole Base64-encoded document.

    Text-level extraction, not XML parsing. Three reasons:

      - it works on files that are not well-formed, which is exactly the case when a
        legacy line break landed inside a tag name
      - it returns the bytes as they are on disk, including line breaks injected
        inside a value, so the corruption is visible rather than normalized away
      - it never materializes the document, and caps what it holds via -MaxChars

    Memory is bounded regardless of file size: while searching, only a few characters
    of look-behind are kept; while capturing an occurrence that was not requested,
    the content is discarded and only the end tag is looked for.

.PARAMETER Path
    File or wildcard pattern. Every matching file is scanned.

.PARAMETER Tag
    Element to extract. With a prefix ("ELAR:DSAK") the match is exact. Without one
    ("DSAK") any prefix matches.

.PARAMETER Index
    1-based occurrence to extract. Default 1. Ignored with -All or -Count.

.PARAMETER All
    Emit every occurrence instead of one.

.PARAMETER Count
    Only report how many occurrences exist. Fastest mode: nothing is captured.

.PARAMETER MaxChars
    Maximum characters captured per occurrence. Default 4000. Longer values are
    truncated and flagged. Raise it for ELAR:Content, or use -OutFile.

.PARAMETER Normalize
    Strip line breaks from the extracted value, showing the value as it was meant to
    be. Without this the raw form is shown and breaks are marked <LF> / <CR>.

.PARAMETER OutFile
    Write the extracted value to a file instead of stdout. With -DecodeBase64 the
    decoded bytes are written, which is how a document is recovered from an INDX.

.PARAMETER DecodeBase64
    Treat the value as Base64 and decode it. Requires -OutFile. Whitespace in the
    value is ignored, as any Base64 decoder does.

.PARAMETER Encoding
    Charset used to read. Defaults to windows-1252, which is what the legacy writer
    emitted; the XML declaration claims ISO-8859-1 and the two differ only over
    0x80-0x9F.

.EXAMPLE
    .\Get-ElarXmlElement.ps1 -Path "...INDX.C143500.xml" -Tag ELAR:DSAK -Index 5

.EXAMPLE
    .\Get-ElarXmlElement.ps1 -Path "...INDX.C143500.xml" -Tag DSAK -All

.EXAMPLE
    .\Get-ElarXmlElement.ps1 -Path "...INDX.C143500.xml" -Tag ELAR:Content -Count

.EXAMPLE
    .\Get-ElarXmlElement.ps1 -Path "...INDX.C143500.xml" -Tag ELAR:Content -Index 3 `
        -MaxChars 0 -DecodeBase64 -OutFile .\doc3.pdf

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Read-only.
    Exit codes: 0 found, 3 not found, 1 error.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $Tag,

    [int] $Index = 1,

    [switch] $All,

    [switch] $Count,

    [int] $MaxChars = 4000,

    [switch] $Normalize,

    [string] $OutFile,

    [switch] $DecodeBase64,

    [string] $Encoding = 'windows-1252'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

function Out-Value {
    param([string] $Text)
    [Console]::Out.WriteLine($Text)
    [Console]::Out.Flush()
}

function Emit-Occurrence {
    param(
        [System.IO.FileInfo] $File,
        [int]    $Occurrence,
        [int]    $Line,
        [long]   $Offset,
        [string] $Value,
        [bool]   $Truncated
    )

    $hasBreak = ($Value.IndexOf("`n") -ge 0 -or $Value.IndexOf("`r") -ge 0)

    Log ("  occurrence={0} line={1} offset={2} length={3}{4}{5}" -f `
         $Occurrence, $Line, $Offset, $Value.Length,
         $(if ($Truncated) { ' TRUNCATED' } else { '' }),
         $(if ($hasBreak)  { ' CONTAINS-LINE-BREAK' } else { '' }))

    if ($OutFile) {
        if ($DecodeBase64) {
            # Any Base64 decoder ignores whitespace, including the breaks the legacy
            # wrapper injected into the payload.
            $clean = [regex]::Replace($Value, '\s', '')
            $bytes = [Convert]::FromBase64String($clean)
            [System.IO.File]::WriteAllBytes($OutFile, $bytes)
            Log ("  decoded {0} byte(s) -> {1}" -f $bytes.Length, $OutFile)
        }
        else {
            $text = if ($Normalize) { $Value -replace "`r`n", '' -replace "[`r`n]", '' } else { $Value }
            [System.IO.File]::WriteAllText($OutFile, $text, $enc)
            Log ("  wrote {0} char(s) -> {1}" -f $text.Length, $OutFile)
        }
        return
    }

    $shown = if ($Normalize) {
                 $Value -replace "`r`n", '' -replace "[`r`n]", ''
             }
             else {
                 $Value -replace "`r", '<CR>' -replace "`n", '<LF>'
             }
    Out-Value $shown
}

if ($DecodeBase64 -and -not $OutFile) { throw "-DecodeBase64 requires -OutFile." }

try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

# ---------------------------------------------------------------------------
# Patterns. A start tag must be followed by whitespace, '/' or '>', so that
# searching for Content does not also match ContentName.
# ---------------------------------------------------------------------------

if ($Tag.Contains(':')) {
    $namePart  = [regex]::Escape($Tag)
    $endName   = [regex]::Escape($Tag)
}
else {
    $namePart  = '(?:[A-Za-z0-9_.\-]+:)?' + [regex]::Escape($Tag)
    $endName   = $namePart
}

$reStart = New-Object System.Text.RegularExpressions.Regex ('<' + $namePart + '(?=[\s/>])')
$reEnd   = New-Object System.Text.RegularExpressions.Regex ('</' + $endName + '\s*>')

$files = @()
foreach ($p in $Path) {
    if (Test-Path -LiteralPath $p -PathType Leaf) { $files += Get-Item -LiteralPath $p }
    else { $files += @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue) }
}
if ($files.Count -eq 0) { Log ("ERROR no file matched: {0}" -f ($Path -join ', ')); exit 1 }

$exitCode = 3
$captureLimit = if ($MaxChars -le 0) { [int]::MaxValue } else { $MaxChars }

foreach ($file in $files) {

    Log ("FILE {0}  {1:n1} MB  tag={2}" -f $file.Name, ($file.Length / 1MB), $Tag)

    $fs = $null; $reader = $null
    $occurrence = 0
    $found      = $false

    # Buffer state
    $buf        = ''
    $line       = 1
    $absOffset  = 0     # characters already discarded from the front of $buf
    $state      = 'SEARCH'
    $keep       = $reStart.ToString().Length + 8   # look-behind while searching
    $capture    = $null
    $capTrunc   = $false
    $capLine    = 0
    $capOffset  = 0
    $wanted     = $false

    try {
        $fs = New-Object System.IO.FileStream($file.FullName,
                  [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite, 1048576)
        $reader = New-Object System.IO.StreamReader($fs, $enc, $false, 1048576)

        $chunk = New-Object char[] 262144

        while ($true) {
            $n = $reader.Read($chunk, 0, $chunk.Length)
            if ($n -le 0 -and $buf.Length -eq 0) { break }
            if ($n -gt 0) { $buf += (New-Object string ($chunk, 0, $n)) }
            $eof = ($n -le 0)

            $progress = $true
            while ($progress) {
                $progress = $false

                if ($state -eq 'SEARCH') {
                    $m = $reStart.Match($buf)
                    if ($m.Success) {
                        # Find the '>' that closes the start tag.
                        $gt = $buf.IndexOf('>', $m.Index)
                        if ($gt -lt 0) { break }   # need more data

                        $occurrence++
                        $wanted = $Count.IsPresent -eq $false -and ($All -or $occurrence -eq $Index)

                        $capLine   = $line + ([regex]::Matches($buf.Substring(0, $m.Index), "`n")).Count
                        $capOffset = $absOffset + $m.Index

                        $selfClosing = ($buf[$gt - 1] -eq '/')
                        if ($selfClosing) {
                            if ($wanted) {
                                Emit-Occurrence -File $file -Occurrence $occurrence -Line $capLine `
                                                -Offset $capOffset -Value '' -Truncated $false
                                $found = $true
                                if (-not $All) { $state = 'DONE'; break }
                            }
                            $consumed = $gt + 1
                        }
                        else {
                            $capture  = New-Object System.Text.StringBuilder
                            $capTrunc = $false
                            $state    = 'CAPTURE'
                            $consumed = $gt + 1
                        }

                        $line += ([regex]::Matches($buf.Substring(0, $consumed), "`n")).Count
                        $absOffset += $consumed
                        $buf = $buf.Substring($consumed)
                        $progress = $true
                    }
                    else {
                        # Discard everything except a short look-behind.
                        if ($buf.Length -gt $keep) {
                            $cut = $buf.Length - $keep
                            $line += ([regex]::Matches($buf.Substring(0, $cut), "`n")).Count
                            $absOffset += $cut
                            $buf = $buf.Substring($cut)
                        }
                    }
                }
                elseif ($state -eq 'CAPTURE') {
                    $m = $reEnd.Match($buf)
                    if ($m.Success) {
                        if ($wanted -and $capture.Length -lt $captureLimit) {
                            $take = [Math]::Min($m.Index, $captureLimit - $capture.Length)
                            [void]$capture.Append($buf, 0, $take)
                            if ($take -lt $m.Index) { $capTrunc = $true }
                        }

                        if ($wanted) {
                            Emit-Occurrence -File $file -Occurrence $occurrence -Line $capLine `
                                            -Offset $capOffset -Value $capture.ToString() -Truncated $capTrunc
                            $found = $true
                            if (-not $All) { $state = 'DONE'; break }
                        }

                        $consumed = $m.Index + $m.Length
                        $line += ([regex]::Matches($buf.Substring(0, $consumed), "`n")).Count
                        $absOffset += $consumed
                        $buf = $buf.Substring($consumed)
                        $capture = $null
                        $state = 'SEARCH'
                        $progress = $true
                    }
                    else {
                        # End tag not in the buffer yet. Keep only a tail long enough
                        # to hold a split end tag; capture the rest if it is wanted.
                        $tail = 64
                        if ($buf.Length -gt $tail) {
                            $cut = $buf.Length - $tail
                            if ($wanted -and $capture.Length -lt $captureLimit) {
                                $take = [Math]::Min($cut, $captureLimit - $capture.Length)
                                [void]$capture.Append($buf, 0, $take)
                                if ($take -lt $cut) { $capTrunc = $true }
                            }
                            $line += ([regex]::Matches($buf.Substring(0, $cut), "`n")).Count
                            $absOffset += $cut
                            $buf = $buf.Substring($cut)
                        }
                    }
                }
            }

            if ($state -eq 'DONE') { break }
            if ($eof) { break }
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
        elseif ($fs) { $fs.Dispose() }
    }

    if ($Count) {
        Log ("  occurrences={0}" -f $occurrence)
        if ($occurrence -gt 0) { $exitCode = 0 }
        continue
    }

    if ($found) { $exitCode = 0 }
    else {
        Log ("  not found: {0} occurrence(s) of {1}, requested #{2}" -f $occurrence, $Tag, $Index)
    }
}

exit $exitCode
