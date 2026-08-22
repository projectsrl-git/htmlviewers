<#
.SYNOPSIS
    Extracts an INDX file back into the flat CSV and the content files it was built
    from, so the same input can be fed to a new implementation and the two outputs
    compared.

.DESCRIPTION
    Reverses the legacy pipeline. For every document in the INDX it collects the
    metadata elements into one CSV row, and optionally decodes ELAR:Content back into
    a file on disk, writing that file's path into the CSV. The result is an input set
    that should regenerate the original INDX byte for byte, modulo the defects the
    original carries.

    Reading is done with a streaming XML reader, so a 500 MB file costs a few MB of
    memory and the Base64 payload is decoded straight to disk without ever being held
    as a string. It follows that **the file must be well-formed**: repair it first if
    it is not. A parse failure here is a real finding, not a limitation of this
    script, and it is reported with a line and column.

    Three outputs, the last two optional:

      - the CSV itself, one row per document, one column per distinct element
      - the decoded content files, named after the document id
      - a properties file with an identity mapping, ready to drive a regeneration

.PARAMETER Path
    The INDX file to reverse.

.PARAMETER CsvOut
    Destination CSV. Defaults to the INDX name with .csv appended.

.PARAMETER ContentDir
    Directory for the decoded payloads. When given, ELAR:Content in the CSV holds the
    path of the decoded file rather than being left empty. The extension comes from
    ELAR:DSAK when present, otherwise .bin.

.PARAMETER PropertiesOut
    Writes a properties file with an identity mapping - tagNameMapping.<COLUMN> for
    every column found, plus input.doc_id_reference - as a starting point for
    regenerating the INDX.

.PARAMETER FamilyType
    Prefix used in the generated properties file. Default CLIAC@DT.

.PARAMETER DocIdTag
    Element treated as the document id, used to name content files and to fill
    input.doc_id_reference. Default RecordId; UniqueReportID is the usual alternative.

.PARAMETER Separator
    CSV separator. Default ';'.

.PARAMETER ContentElement
    Local name of the payload element. Default Content.

.PARAMETER DocElement
    Local name of the element delimiting one document. Default Doc.

.PARAMETER Encoding
    Charset used to read the INDX and write the CSV. Default windows-1252.

.PARAMETER MaxDocs
    Stop after this many documents. 0 means all. Useful for a quick sample.

.EXAMPLE
    .\Export-ElarIndxCsv.ps1 -Path "...INDX.C113924" -CsvOut .\rebuild.csv

.EXAMPLE
    .\Export-ElarIndxCsv.ps1 -Path "...INDX.C113924" -CsvOut .\rebuild.csv `
        -ContentDir .\payload -PropertiesOut .\rebuild.properties -MaxDocs 50

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Read-only on the INDX.
    Exit codes: 0 completed, 2 completed with warnings, 1 error.

    A value containing the separator is quoted. The legacy reader does not understand
    quoting and would mis-parse such a row, so the count of quoted values is reported:
    if it is not zero, the legacy tool could never have produced this data from a CSV,
    which is itself worth knowing.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [string] $CsvOut,

    [string] $ContentDir,

    [string] $PropertiesOut,

    [string] $FamilyType = 'CLIAC@DT',

    [string] $DocIdTag = 'RecordId',

    [string] $Separator = ';',

    [string] $ContentElement = 'Content',

    [string] $DocElement = 'Doc',

    [string] $Encoding = 'windows-1252',

    [int] $MaxDocs = 0
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
catch { throw "Unknown charset '$Encoding'. Try windows-1252, ISO-8859-1, or utf-8." }

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }
$src = Get-Item -LiteralPath $Path

if ([string]::IsNullOrEmpty($CsvOut)) { $CsvOut = $src.FullName + '.csv' }
if ($ContentDir -and -not (Test-Path -LiteralPath $ContentDir)) {
    New-Item -ItemType Directory -Path $ContentDir -Force | Out-Null
}

$exitCode = 0
$fs = $null; $sr = $null; $reader = $null

# Column order is the order of first appearance, which keeps the CSV readable and
# matches the order the template emits.
$columns = New-Object 'System.Collections.Generic.List[string]'
$colSet  = New-Object 'System.Collections.Generic.HashSet[string]'
$rows    = New-Object 'System.Collections.Generic.List[hashtable]'

$docs = 0; $quoted = 0; $payloads = 0; $payloadBytes = 0L
$emptyDocId = 0

try {
    Log ("START file={0} {1:n1} MB csv={2}" -f $src.Name, ($src.Length / 1MB), $CsvOut)

    $fs = New-Object System.IO.FileStream($src.FullName,
              [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
              [System.IO.FileShare]::ReadWrite, 1048576)
    # Explicit charset: the declaration in these files is not reliable.
    $sr = New-Object System.IO.StreamReader($fs, $enc, $false, 1048576)

    $set = New-Object System.Xml.XmlReaderSettings
    $set.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $set.IgnoreWhitespace = $true
    $set.IgnoreComments   = $true
    $reader = [System.Xml.XmlReader]::Create($sr, $set)

    $cur = $null
    $sw  = [System.Diagnostics.Stopwatch]::StartNew()
    $nextTick = 10.0

    while ($reader.Read()) {

        if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement -and
                $reader.LocalName -eq $DocElement -and $null -ne $cur) {
                $rows.Add($cur); $cur = $null
                if ($MaxDocs -gt 0 -and $rows.Count -ge $MaxDocs) { break }
            }
            continue
        }

        if ($reader.LocalName -eq $DocElement) {
            $docs++
            $cur = @{}
            if ($sw.Elapsed.TotalSeconds -ge $nextTick) {
                $nextTick = $sw.Elapsed.TotalSeconds + 10
                $pct = if ($src.Length -gt 0) { ($fs.Position / [double]$src.Length) * 100 } else { 0 }
                Log ("    ... {0:n0}% {1:n0} documents" -f $pct, $docs)
            }
            continue
        }

        if ($null -eq $cur) { continue }   # outside a document

        $name = $reader.Name          # keep the prefix: it is what the template uses
        $local = $reader.LocalName

        if (-not $colSet.Contains($name)) { [void]$colSet.Add($name); $columns.Add($name) }

        if ($local -eq $ContentElement) {
            if (-not $ContentDir) {
                # Skip the payload without materializing it.
                $reader.Skip()
                continue
            }

            # Name the file after the document id, which must already have been read:
            # the id element precedes Content in the template.
            $id = if ($cur.ContainsKey('ELAR:' + $DocIdTag)) { $cur['ELAR:' + $DocIdTag] }
                  else { ($cur.Values | Select-Object -First 1) }
            if ([string]::IsNullOrEmpty($id)) { $id = ('doc{0:D6}' -f $docs); $emptyDocId++ }

            $extRaw = if ($cur.ContainsKey('ELAR:DSAK')) { $cur['ELAR:DSAK'] } else { '' }
            $ext = if ($extRaw) { '.' + $extRaw.Trim().ToLowerInvariant() } else { '.bin' }

            $outFile = Join-Path $ContentDir (($id -replace '[\\/:*?"<>|]', '_') + $ext)

            $stream = [System.IO.File]::Create($outFile)
            try {
                $buf = New-Object byte[] 65536
                $n = 0
                # Decodes in chunks; whitespace inside the payload is ignored, so the
                # legacy line wrapping does not matter here.
                while (($n = $reader.ReadElementContentAsBase64($buf, 0, $buf.Length)) -gt 0) {
                    $stream.Write($buf, 0, $n)
                    $payloadBytes += $n
                }
            }
            finally { $stream.Dispose() }

            $payloads++
            $cur[$name] = $outFile
            continue
        }

        # Ordinary metadata element: take its text, empty if self-closed.
        if ($reader.IsEmptyElement) { $cur[$name] = '' }
        else { $cur[$name] = $reader.ReadElementContentAsString() }
    }

    if ($null -ne $cur) { $rows.Add($cur) }
}
catch [System.Xml.XmlException] {
    Log ("ERROR not well-formed at line {0} position {1}: {2}" -f `
         $_.Exception.LineNumber, $_.Exception.LinePosition, $_.Exception.Message)
    Log  "ERROR repair the file before reversing it"
    Log  "END exit=1"
    exit 1
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}
finally {
    if ($reader) { $reader.Dispose() }
    elseif ($sr) { $sr.Dispose() }
    elseif ($fs) { $fs.Dispose() }
}

# ---------------------------------------------------------------------------

$needsQuote = [char[]]@($Separator[0], '"', "`r", "`n")

function Format-Field {
    param([string] $Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    if ($Value.IndexOfAny($needsQuote) -ge 0) {
        $script:quoted++
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

$out = New-Object System.IO.StreamWriter($CsvOut, $false, $enc)
try {
    $out.Write(($columns -join $Separator)); $out.Write("`n")
    foreach ($r in $rows) {
        $line = New-Object System.Text.StringBuilder
        for ($k = 0; $k -lt $columns.Count; $k++) {
            if ($k -gt 0) { [void]$line.Append($Separator) }
            $v = if ($r.ContainsKey($columns[$k])) { [string]$r[$columns[$k]] } else { '' }
            [void]$line.Append((Format-Field -Value $v))
        }
        $out.Write($line.ToString()); $out.Write("`n")
    }
}
finally { $out.Dispose() }

if ($PropertiesOut) {
    $p = New-Object System.IO.StreamWriter($PropertiesOut, $false, $enc)
    try {
        $p.Write("# Generated by Export-ElarIndxCsv from " + $src.Name + "`n")
        $p.Write("# Identity mapping: CSV column name equals the element it fills.`n`n")
        $idCol = @($columns | Where-Object { $_ -like ('*:' + $DocIdTag) }) | Select-Object -First 1
        if (-not $idCol) { $idCol = $columns[0] }
        $p.Write(("{0}.input.doc_id_reference={1}`n" -f $FamilyType, $idCol))
        foreach ($c in $columns) {
            $p.Write(("{0}.tagNameMapping.{1}={1}`n" -f $FamilyType, $c))
        }
    }
    finally { $p.Dispose() }
}

Log "SUMMARY"
Log ("  documents={0}" -f $rows.Count)
Log ("  columns={0}" -f $columns.Count)
Log ("  quotedValues={0}" -f $quoted)
if ($ContentDir) {
    Log ("  payloadsDecoded={0}" -f $payloads)
    Log ("  payloadBytes={0:n0}" -f $payloadBytes)
}
if ($emptyDocId -gt 0) {
    Log ("  WARNING {0} document(s) had no {1}; content files were named by position" -f $emptyDocId, $DocIdTag)
    $exitCode = 2
}
if ($quoted -gt 0) {
    Log ("  NOTE {0} value(s) contain the separator and were quoted" -f $quoted)
    Log  "  NOTE the legacy reader does not understand quoting, so it could not have produced these rows"
    $exitCode = 2
}
Log ("  csv={0}" -f $CsvOut)
if ($PropertiesOut) { Log ("  properties={0}" -f $PropertiesOut) }
Log ("END exit={0}" -f $exitCode)

exit $exitCode
