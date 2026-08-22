<#
.SYNOPSIS
    Extracts INDX files back into the flat CSV and the content files they were built
    from, so the same inputs can be fed to a new implementation and the outputs
    compared.

.DESCRIPTION
    Reverses the legacy pipeline. For every document it collects the metadata
    elements into one CSV row, and optionally decodes ELAR:Content back into a file on
    disk, writing that file's path into the CSV. The result is an input set that
    should regenerate the original INDX, modulo the defects the original carries.

    Reading is streaming, so a 530 MB file costs a few MB of memory and the Base64
    payload is decoded straight to disk without ever being held as a string. It
    follows that the file must parse: repair it first if it does not. A parse failure
    here is a real finding, reported with line and column.

    Structure is not assumed. The INDX nests containers - ELAR:Data holds
    ELAR:Document, which holds ContentName, HashValue, DSAK and Content - so the
    reader takes text where text exists and ignores elements that only hold other
    elements. Column names are the fully qualified element names, in order of first
    appearance.

    Three outputs, the last two optional:

      - the CSV, one row per document, one column per element that carries a value
      - the decoded content files, named after the document id
      - a properties file with an identity mapping, ready to drive a regeneration

.PARAMETER Path
    File, directory, or wildcard. With several matches, one CSV is written per input
    file and -CsvOut must be omitted.

.PARAMETER Filter
    Wildcard applied when Path names a directory. Default *INDX*.

.PARAMETER CsvOut
    Destination CSV for a single input file. Defaults to the input name with .csv
    appended, alongside the input.

.PARAMETER ContentDir
    Directory for the decoded payloads. When given, ELAR:Content in the CSV holds the
    path of the decoded file instead of being left empty. The extension comes from
    ELAR:DSAK when present, otherwise .bin.

.PARAMETER PropertiesOut
    Writes a properties file with an identity mapping - tagNameMapping.<COLUMN> for
    every column found, plus input.doc_id_reference - as a starting point for a
    regeneration.

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
    Stop after this many documents per file. 0 means all.

.EXAMPLE
    .\Export-ElarIndxCsv.ps1 -Path "...INDX.C113925" -CsvOut .\rebuild.csv

.EXAMPLE
    .\Export-ElarIndxCsv.ps1 -Path "G:\ELAR\OUT\CMOD\S210969_CLIAC" `
        -ContentDir .\payload -PropertiesOut .\rebuild.properties -MaxDocs 50

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Read-only on the INDX.
    Exit codes: 0 completed, 2 completed with warnings, 1 error.

    A value containing the separator is quoted. The legacy reader does not understand
    quoting and would mis-parse such a row, so the count is reported: if it is not
    zero, the legacy tool could not have produced this data from a CSV.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [string] $Filter = '*INDX*',

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
$files = @($files | Where-Object { $_.Extension -notin @('.bak', '.tmp', '.ps1', '.csv') } |
                    Sort-Object FullName -Unique)

if ($files.Count -eq 0) {
    Log ("WARN no file matched: {0}" -f ($Path -join ', '))
    Log  "END exit=0"
    exit 0
}
if ($files.Count -gt 1 -and $CsvOut) {
    throw "-CsvOut names a single destination but $($files.Count) files matched. Omit it to write one CSV per input."
}

if ($ContentDir -and -not (Test-Path -LiteralPath $ContentDir)) {
    New-Item -ItemType Directory -Path $ContentDir -Force | Out-Null
}

$exitCode = 0

# ---------------------------------------------------------------------------

function Export-OneIndx {
    param(
        [System.IO.FileInfo]   $File,
        [string]               $Csv,
        [System.Text.Encoding] $Enc
    )

    $fs = $null; $sr = $null; $reader = $null

    $columns = New-Object 'System.Collections.Generic.List[string]'
    $colSet  = New-Object 'System.Collections.Generic.HashSet[string]'
    $rows    = New-Object 'System.Collections.Generic.List[hashtable]'

    $docs = 0; $payloads = 0; $payloadBytes = 0L; $emptyDocId = 0
    $parseError = $null

    try {
        $fs = New-Object System.IO.FileStream($File.FullName,
                  [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite, 1048576)
        # Explicit charset: the declaration in these files is not reliable.
        $sr = New-Object System.IO.StreamReader($fs, $Enc, $false, 1048576)

        $set = New-Object System.Xml.XmlReaderSettings
        $set.DtdProcessing    = [System.Xml.DtdProcessing]::Prohibit
        $set.IgnoreWhitespace = $true
        $set.IgnoreComments   = $true
        $reader = [System.Xml.XmlReader]::Create($sr, $set)

        $cur     = $null
        $pending = $null      # element awaiting its text, if it has any
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $nextTick = 10.0
        $stop = $false

        while (-not $stop -and $reader.Read()) {

            switch ($reader.NodeType) {

                ([System.Xml.XmlNodeType]::Element) {
                    $local = $reader.LocalName
                    $name  = $reader.Name

                    if ($local -eq $DocElement) {
                        $docs++
                        $cur = @{}
                        $pending = $null

                        if ($sw.Elapsed.TotalSeconds -ge $nextTick) {
                            $nextTick = $sw.Elapsed.TotalSeconds + 10
                            $pct = if ($File.Length -gt 0) { ($fs.Position / [double]$File.Length) * 100 } else { 0 }
                            Log ("    ... {0:n0}% {1:n0} documents" -f $pct, $docs)
                        }
                        break
                    }

                    if ($null -eq $cur) { break }   # outside a document

                    if ($local -eq $ContentElement) {
                        if (-not $ContentDir) {
                            if (-not $reader.IsEmptyElement) { $reader.Skip() }
                            break
                        }

                        # The id element precedes Content in the template, so it is
                        # already known by the time we get here.
                        $idKey = @($cur.Keys | Where-Object { $_ -like ('*:' + $DocIdTag) -or $_ -eq $DocIdTag }) |
                                 Select-Object -First 1
                        $id = if ($idKey) { [string]$cur[$idKey] } else { '' }
                        if ([string]::IsNullOrEmpty($id)) { $id = ('doc{0:D6}' -f $docs); $script:emptyIdSeen++ }

                        $dsakKey = @($cur.Keys | Where-Object { $_ -like '*:DSAK' -or $_ -eq 'DSAK' }) |
                                   Select-Object -First 1
                        $extRaw = if ($dsakKey) { [string]$cur[$dsakKey] } else { '' }
                        $ext = if ($extRaw.Trim()) { '.' + $extRaw.Trim().ToLowerInvariant() } else { '.bin' }

                        $outFile = Join-Path $ContentDir (($id -replace '[\\/:*?"<>|]', '_') + $ext)

                        if (-not $colSet.Contains($name)) { [void]$colSet.Add($name); $columns.Add($name) }

                        if ($reader.IsEmptyElement) { $cur[$name] = '' ; break }

                        $stream = [System.IO.File]::Create($outFile)
                        try {
                            $buf = New-Object byte[] 65536
                            $n = 0
                            # Whitespace inside the payload is ignored by the decoder,
                            # so the line wrapping does not matter here.
                            while (($n = $reader.ReadElementContentAsBase64($buf, 0, $buf.Length)) -gt 0) {
                                $stream.Write($buf, 0, $n)
                                $payloadBytes += $n
                            }
                        }
                        finally { $stream.Dispose() }

                        $payloads++
                        $cur[$name] = $outFile
                        break
                    }

                    if ($reader.IsEmptyElement) {
                        # Self-closed: a present but valueless element.
                        if (-not $colSet.Contains($name)) { [void]$colSet.Add($name); $columns.Add($name) }
                        $cur[$name] = ''
                        $pending = $null
                    }
                    else {
                        # Might be a leaf with text, might be a container. We find out
                        # when the next node arrives, so nothing is assumed here.
                        $pending = $name
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::Text) {
                    if ($null -ne $cur -and $null -ne $pending) {
                        if (-not $colSet.Contains($pending)) { [void]$colSet.Add($pending); $columns.Add($pending) }
                        $cur[$pending] = $reader.Value
                        $pending = $null
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::CDATA) {
                    if ($null -ne $cur -and $null -ne $pending) {
                        if (-not $colSet.Contains($pending)) { [void]$colSet.Add($pending); $columns.Add($pending) }
                        $cur[$pending] = $reader.Value
                        $pending = $null
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::EndElement) {
                    if ($reader.LocalName -eq $DocElement -and $null -ne $cur) {
                        $rows.Add($cur); $cur = $null; $pending = $null
                        if ($MaxDocs -gt 0 -and $rows.Count -ge $MaxDocs) { $stop = $true }
                    }
                    else {
                        # A container closing, or a leaf that turned out to be empty.
                        if ($null -ne $cur -and $null -ne $pending -and $reader.Name -eq $pending) {
                            if (-not $colSet.Contains($pending)) { [void]$colSet.Add($pending); $columns.Add($pending) }
                            $cur[$pending] = ''
                        }
                        $pending = $null
                    }
                    break
                }
            }
        }

        if ($null -ne $cur) { $rows.Add($cur) }
    }
    catch [System.Xml.XmlException] {
        $parseError = ("parse error at line {0} position {1}: {2}" -f `
                       $_.Exception.LineNumber, $_.Exception.LinePosition, $_.Exception.Message)
    }
    finally {
        if ($reader) { $reader.Dispose() }
        elseif ($sr) { $sr.Dispose() }
        elseif ($fs) { $fs.Dispose() }
    }

    if ($parseError) {
        return [pscustomobject]@{ Ok = $false; Error = $parseError; Rows = 0; Columns = 0; Quoted = 0
                                  Payloads = $payloads; PayloadBytes = $payloadBytes; Columns2 = $columns }
    }

    # --- write the CSV ---
    $needsQuote = [char[]]@($Separator[0], '"', "`r", "`n")
    $quoted = 0

    $out = New-Object System.IO.StreamWriter($Csv, $false, $Enc)
    try {
        $out.Write(($columns -join $Separator)); $out.Write("`n")
        foreach ($r in $rows) {
            $sb = New-Object System.Text.StringBuilder
            for ($k = 0; $k -lt $columns.Count; $k++) {
                if ($k -gt 0) { [void]$sb.Append($Separator) }
                $v = if ($r.ContainsKey($columns[$k])) { [string]$r[$columns[$k]] } else { '' }
                if ($v -and $v.IndexOfAny($needsQuote) -ge 0) {
                    $quoted++
                    [void]$sb.Append('"').Append($v.Replace('"', '""')).Append('"')
                }
                else { [void]$sb.Append($v) }
            }
            $out.Write($sb.ToString()); $out.Write("`n")
        }
    }
    finally { $out.Dispose() }

    return [pscustomobject]@{
        Ok = $true; Error = ''; Rows = $rows.Count; Columns = $columns.Count
        Quoted = $quoted; Payloads = $payloads; PayloadBytes = $payloadBytes; Columns2 = $columns
    }
}

# ---------------------------------------------------------------------------

$script:emptyIdSeen = 0
$allColumns = $null
$i = 0

Log ("START files={0} charset={1} contentDir={2}" -f $files.Count, $Encoding, $(if ($ContentDir) { $ContentDir } else { '(none)' }))

foreach ($f in $files) {
    $i++
    $csv = if ($CsvOut) { $CsvOut } else { $f.FullName + '.csv' }
    Log ("[{0}/{1}] {2}  {3:n1} MB -> {4}" -f $i, $files.Count, $f.Name, ($f.Length / 1MB), $csv)

    $r = Export-OneIndx -File $f -Csv $csv -Enc $enc

    if (-not $r.Ok) {
        Log ("    ERROR {0}" -f $r.Error)
        Log  "    repair the file before reversing it"
        $exitCode = 1
        continue
    }

    Log ("    documents={0} columns={1} quotedValues={2}" -f $r.Rows, $r.Columns, $r.Quoted)
    if ($ContentDir) {
        Log ("    payloadsDecoded={0} payloadBytes={1:n0}" -f $r.Payloads, $r.PayloadBytes)
    }
    if ($r.Quoted -gt 0) {
        Log ("    NOTE {0} value(s) contain the separator and were quoted" -f $r.Quoted)
        Log  "    NOTE the legacy reader does not understand quoting, so it could not have produced these rows"
        if ($exitCode -eq 0) { $exitCode = 2 }
    }
    if ($null -eq $allColumns) { $allColumns = $r.Columns2 }
}

if ($script:emptyIdSeen -gt 0) {
    Log ("WARN {0} document(s) had no {1}; content files were named by position" -f $script:emptyIdSeen, $DocIdTag)
    if ($exitCode -eq 0) { $exitCode = 2 }
}

if ($PropertiesOut -and $null -ne $allColumns -and $allColumns.Count -gt 0) {
    $p = New-Object System.IO.StreamWriter($PropertiesOut, $false, $enc)
    try {
        $p.Write("# Generated by Export-ElarIndxCsv`n")
        $p.Write("# Identity mapping: the CSV column name equals the element it fills.`n`n")
        $idCol = @($allColumns | Where-Object { $_ -like ('*:' + $DocIdTag) }) | Select-Object -First 1
        if (-not $idCol) { $idCol = $allColumns[0] }
        $p.Write(("{0}.input.doc_id_reference={1}`n" -f $FamilyType, $idCol))
        foreach ($c in $allColumns) {
            $p.Write(("{0}.tagNameMapping.{1}={1}`n" -f $FamilyType, $c))
        }
    }
    finally { $p.Dispose() }
    Log ("properties={0}" -f $PropertiesOut)
}

Log ("END exit={0}" -f $exitCode)
exit $exitCode
