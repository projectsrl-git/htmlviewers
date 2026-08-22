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

.PARAMETER UniqueKeyWhereCondition
    Name of the element holding the key, e.g. ELAR:RecordId or just RecordId. When
    given, a .txt is written next to the CSV containing a ready-to-paste SQL IN
    clause listing every key in the file, for querying the source system about
    exactly the documents this INDX carries.

    Values are single-quoted and embedded quotes are doubled. Duplicates are removed
    and counted: a repeated key in an INDX is itself a defect worth knowing about.

.PARAMETER WhereOut
    Destination for the IN clause. Defaults to the CSV path with .where.txt appended.
    Only valid with a single input file.

.PARAMETER WhereColumnName
    Column name to use inside the SQL clause. Defaults to the local name of the key
    element, so ELAR:RecordId becomes RecordId. Set it when the source column is
    named differently from the XML element.

.PARAMETER MaxInItems
    Maximum values per IN list. Default 1000, which is Oracle's limit on expressions
    in an IN clause. Beyond it the clause is split into several IN lists combined
    with OR, wrapped in parentheses, so the result stays valid however many keys
    there are.

.PARAMETER KeepPrefix
    Keep the namespace prefix in the CSV header, so the column is ELAR:RecordId
    rather than RecordId. Off by default: the header carries local names, and the
    generated properties file maps each one back to its fully qualified element.

.PARAMETER NewlineInValue
    What to do with a value that contains a line break.

      SPACE  (default) replace each break with a single space
      STRIP  remove the break, joining the halves
      KEEP   leave it, quoting the field

    KEEP produces a CSV that is formally correct but spans several physical lines per
    record, which a line-oriented reader - the legacy one included - splits into
    different records. SPACE keeps one line per record at the cost of altering the
    value; the count of altered values is always reported so the change is never
    silent.

    Note that a line break inside a metadata value is itself a known defect of the
    legacy writer, not real data. A non-zero count here usually means the source INDX
    carries corrupted values.

.PARAMETER NoVerify
    Skip the verification pass over the CSV just written.

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

    [int] $MaxDocs = 0,

    [string] $UniqueKeyWhereCondition,

    [string] $WhereOut,

    [string] $WhereColumnName,

    [int] $MaxInItems = 1000,

    [switch] $KeepPrefix,

    [ValidateSet('KEEP', 'SPACE', 'STRIP')]
    [string] $NewlineInValue = 'SPACE',

    [switch] $NoVerify
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
if ($files.Count -gt 1 -and $WhereOut) {
    throw "-WhereOut names a single destination but $($files.Count) files matched. Omit it to write one file per input."
}
if ($MaxInItems -lt 1) { throw "-MaxInItems must be at least 1." }

if ($ContentDir -and -not (Test-Path -LiteralPath $ContentDir)) {
    New-Item -ItemType Directory -Path $ContentDir -Force | Out-Null
}

$exitCode = 0

# ---------------------------------------------------------------------------

function Get-DisplayName {
    param([string] $Name)
    if ($KeepPrefix) { return $Name }
    $i = $Name.IndexOf(':')
    if ($i -ge 0) { return $Name.Substring($i + 1) }
    return $Name
}

<#
    Reads the CSV back with a proper quoted parser and checks two things: that every
    record has exactly as many fields as the header, and that the record count
    matches what was extracted. Together those catch a truncated write, a stray
    separator, and an unbalanced quote - the three ways this file can be wrong in a
    manner that is invisible until something downstream mis-parses it.
#>
function Test-CsvIntegrity {
    param(
        [string]               $Csv,
        [string]               $Sep,
        [int]                  $ExpectedRecords,
        [System.Text.Encoding] $Enc
    )

    $sepCh = $Sep[0]
    $reader = $null
    $header = -1; $records = 0; $bad = New-Object 'System.Collections.Generic.List[string]'
    $multiline = 0; $lastFieldCount = 0

    try {
        $reader = New-Object System.IO.StreamReader($Csv, $Enc, $false, 1048576)

        $fields = 1
        $inQuote = $false
        $sawAny = $false
        $recordSpansLines = $false
        $physLine = 1

        while (($ch = $reader.Read()) -ge 0) {
            $c = [char]$ch
            $sawAny = $true

            if ($inQuote) {
                if ($c -eq '"') {
                    if ([char]$reader.Peek() -eq '"') { [void]$reader.Read() }  # escaped quote
                    else { $inQuote = $false }
                }
                elseif ($c -eq "`n") { $recordSpansLines = $true; $physLine++ }
                continue
            }

            if ($c -eq '"')      { $inQuote = $true; continue }
            if ($c -eq $sepCh)   { $fields++; continue }
            if ($c -eq "`r")     { continue }

            if ($c -eq "`n") {
                $physLine++
                if ($header -lt 0) { $header = $fields }
                else {
                    $records++
                    if ($fields -ne $header) {
                        if ($bad.Count -lt 20) {
                            [void]$bad.Add(('record {0} has {1} field(s), header has {2}' -f $records, $fields, $header))
                        }
                    }
                    if ($recordSpansLines) { $multiline++ }
                }
                $lastFieldCount = $fields
                $fields = 1
                $recordSpansLines = $false
            }
        }

        # A final record without a trailing newline still counts.
        if ($sawAny -and $fields -gt 1) {
            if ($header -lt 0) { $header = $fields }
            else {
                $records++
                if ($fields -ne $header -and $bad.Count -lt 20) {
                    [void]$bad.Add(('record {0} has {1} field(s), header has {2}' -f $records, $fields, $header))
                }
                if ($recordSpansLines) { $multiline++ }
            }
        }

        $unbalanced = $inQuote
    }
    finally { if ($reader) { $reader.Dispose() } }

    return [pscustomobject]@{
        HeaderFields = $header
        Records      = $records
        Mismatches   = $bad
        Multiline    = $multiline
        Unbalanced   = $unbalanced
        Truncated    = ($records -ne $ExpectedRecords)
        Expected     = $ExpectedRecords
    }
}

function Write-WhereClause {
    param(
        [string[]]             $Values,
        [string]               $Column,
        [string]               $Destination,
        [int]                  $ChunkSize,
        [System.Text.Encoding] $Enc
    )

    $w = New-Object System.IO.StreamWriter($Destination, $false, $Enc)
    try {
        $chunks = [Math]::Ceiling($Values.Count / [double]$ChunkSize)

        # Several IN lists are combined with OR and wrapped, so the clause stays
        # valid past the database's limit on expressions in a single IN.
        if ($chunks -gt 1) { $w.Write("(`n") }

        for ($c = 0; $c -lt $chunks; $c++) {
            $slice = $Values[($c * $ChunkSize) .. ([Math]::Min(($c + 1) * $ChunkSize, $Values.Count) - 1)]

            if ($c -gt 0) { $w.Write("`nOR`n") }
            $w.Write($Column); $w.Write(" IN (`n")

            $line = New-Object System.Text.StringBuilder
            for ($k = 0; $k -lt $slice.Count; $k++) {
                $item = "'" + ($slice[$k].Replace("'", "''")) + "'"
                if ($k -lt $slice.Count - 1) { $item += ',' }
                # Wrap for readability: these files are read by people.
                if ($line.Length + $item.Length -gt 100) {
                    $w.Write($line.ToString()); $w.Write("`n")
                    $line.Length = 0
                }
                [void]$line.Append($item)
            }
            if ($line.Length -gt 0) { $w.Write($line.ToString()); $w.Write("`n") }
            $w.Write(")")
        }

        if ($chunks -gt 1) { $w.Write("`n)") }
        $w.Write("`n")
    }
    finally { $w.Dispose() }

    return $chunks
}

function Export-OneIndx {
    param(
        [System.IO.FileInfo]   $File,
        [string]               $Csv,
        [string]               $Where,
        [System.Text.Encoding] $Enc
    )

    $fs = $null; $sr = $null; $reader = $null

    $columns = New-Object 'System.Collections.Generic.List[string]'
    $colSet  = New-Object 'System.Collections.Generic.HashSet[string]'
    $colFull = @{}     # display name -> fully qualified element name
    $collide = New-Object 'System.Collections.Generic.List[string]'
    $rows    = New-Object 'System.Collections.Generic.List[hashtable]'
    $newlineValues = 0

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

                        $disp = Get-DisplayName -Name $name
                        if (-not $colSet.Contains($disp)) {
                            [void]$colSet.Add($disp); $columns.Add($disp); $colFull[$disp] = $name
                        }
                        elseif ($colFull[$disp] -ne $name -and $collide.Count -lt 10) {
                            [void]$collide.Add(('{0} <- {1} and {2}' -f $disp, $colFull[$disp], $name))
                        }

                        if ($reader.IsEmptyElement) { $cur[$disp] = '' ; break }

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
                        $cur[$disp] = $outFile
                        break
                    }

                    if ($reader.IsEmptyElement) {
                        # Self-closed: a present but valueless element.
                        $disp = Get-DisplayName -Name $name
                        if (-not $colSet.Contains($disp)) {
                            [void]$colSet.Add($disp); $columns.Add($disp); $colFull[$disp] = $name
                        }
                        $cur[$disp] = ''
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
                        $disp = Get-DisplayName -Name $pending
                        if (-not $colSet.Contains($disp)) {
                            [void]$colSet.Add($disp); $columns.Add($disp); $colFull[$disp] = $pending
                        }
                        $v = $reader.Value
                        if ($v.IndexOf("`n") -ge 0 -or $v.IndexOf("`r") -ge 0) {
                            $newlineValues++
                            if     ($NewlineInValue -eq 'SPACE') { $v = $v -replace "`r`n", ' ' -replace "[`r`n]", ' ' }
                            elseif ($NewlineInValue -eq 'STRIP') { $v = $v -replace "`r`n", ''  -replace "[`r`n]", '' }
                        }
                        $cur[$disp] = $v
                        $pending = $null
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::CDATA) {
                    if ($null -ne $cur -and $null -ne $pending) {
                        $disp = Get-DisplayName -Name $pending
                        if (-not $colSet.Contains($disp)) {
                            [void]$colSet.Add($disp); $columns.Add($disp); $colFull[$disp] = $pending
                        }
                        $cur[$disp] = $reader.Value
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
                            $disp = Get-DisplayName -Name $pending
                            if (-not $colSet.Contains($disp)) {
                                [void]$colSet.Add($disp); $columns.Add($disp); $colFull[$disp] = $pending
                            }
                            $cur[$disp] = ''
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

    # --- optional IN clause over the key column ---
    $keyTotal = 0; $keyUnique = 0; $keyEmpty = 0; $keyChunks = 0; $keyCol = ''

    if ($UniqueKeyWhereCondition) {
        $keyCol = @($columns | Where-Object {
                        $_ -eq $UniqueKeyWhereCondition -or
                        $_ -like ('*:' + $UniqueKeyWhereCondition)
                    }) | Select-Object -First 1

        if (-not $keyCol) {
            return [pscustomobject]@{
                Ok = $false
                Error = ("key element '{0}' not found; columns are: {1}" -f $UniqueKeyWhereCondition, ($columns -join ', '))
                Rows = $rows.Count; Columns = $columns.Count; Quoted = $quoted
                Payloads = $payloads; PayloadBytes = $payloadBytes; Columns2 = $columns
                ColFull = $colFull; Collisions = $collide; NewlineValues = $newlineValues
                KeyTotal = 0; KeyUnique = 0; KeyEmpty = 0; KeyChunks = 0; KeyColumn = ''
                Verify = $null
            }
        }

        $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
        $ordered  = New-Object 'System.Collections.Generic.List[string]'
        foreach ($r in $rows) {
            $v = if ($r.ContainsKey($keyCol)) { [string]$r[$keyCol] } else { '' }
            if ([string]::IsNullOrWhiteSpace($v)) { $keyEmpty++; continue }
            $keyTotal++
            # Order of first appearance, so the clause mirrors the file.
            if ($seenKeys.Add($v)) { $ordered.Add($v) }
        }
        $keyUnique = $ordered.Count

        $sqlName = if ($WhereColumnName) { $WhereColumnName }
                   else { $keyCol.Substring($keyCol.IndexOf(':') + 1) }

        if ($ordered.Count -gt 0) {
            $keyChunks = Write-WhereClause -Values $ordered.ToArray() -Column $sqlName `
                                           -Destination $Where -ChunkSize $MaxInItems -Enc $Enc
        }
    }

    # --- verify what was written ---
    $verify = $null
    if (-not $NoVerify) {
        $verify = Test-CsvIntegrity -Csv $Csv -Sep $Separator -ExpectedRecords $rows.Count -Enc $Enc
    }

    return [pscustomobject]@{
        Ok = $true; Error = ''; Rows = $rows.Count; Columns = $columns.Count
        Quoted = $quoted; Payloads = $payloads; PayloadBytes = $payloadBytes; Columns2 = $columns
        ColFull = $colFull; Collisions = $collide; NewlineValues = $newlineValues
        KeyTotal = $keyTotal; KeyUnique = $keyUnique; KeyEmpty = $keyEmpty
        KeyChunks = $keyChunks; KeyColumn = $keyCol
        Verify = $verify
    }
}

# ---------------------------------------------------------------------------

$script:emptyIdSeen = 0
$allColumns = $null
$allColFull = $null
$i = 0

Log ("START files={0} charset={1} contentDir={2}" -f $files.Count, $Encoding, $(if ($ContentDir) { $ContentDir } else { '(none)' }))

foreach ($f in $files) {
    $i++
    $csv = if ($CsvOut) { $CsvOut } else { $f.FullName + '.csv' }
    $whr = if ($WhereOut) { $WhereOut } else { $csv + '.where.txt' }
    Log ("[{0}/{1}] {2}  {3:n1} MB -> {4}" -f $i, $files.Count, $f.Name, ($f.Length / 1MB), $csv)

    $r = Export-OneIndx -File $f -Csv $csv -Where $whr -Enc $enc

    if (-not $r.Ok) {
        Log ("    ERROR {0}" -f $r.Error)
        Log  "    repair the file before reversing it"
        $exitCode = 1
        continue
    }

    Log ("    documents={0} columns={1} quotedValues={2}" -f $r.Rows, $r.Columns, $r.Quoted)

    if ($r.NewlineValues -gt 0) {
        Log ("    NOTE {0} value(s) contained a line break; policy={1}" -f $r.NewlineValues, $NewlineInValue)
        Log  "    NOTE a line break inside a metadata value is a defect of the source INDX, not data"
        if ($exitCode -eq 0) { $exitCode = 2 }
    }
    if (@($r.Collisions).Count -gt 0) {
        Log ("    WARNING {0} column name collision(s) after dropping the prefix:" -f @($r.Collisions).Count)
        foreach ($c in @($r.Collisions)) { Log ("      {0}" -f $c) }
        Log  "    WARNING re-run with -KeepPrefix to keep the columns distinct"
        $exitCode = 2
    }

    if ($null -ne $r.Verify) {
        $v = $r.Verify
        Log ("    verify headerFields={0} records={1} multilineRecords={2}" -f $v.HeaderFields, $v.Records, $v.Multiline)
        if ($v.Unbalanced) {
            Log  "    ERROR unbalanced quote: the CSV is truncated or malformed"
            $exitCode = 1
        }
        if ($v.Truncated) {
            Log ("    ERROR record count {0} does not match the {1} document(s) extracted" -f $v.Records, $v.Expected)
            $exitCode = 1
        }
        if (@($v.Mismatches).Count -gt 0) {
            Log ("    ERROR {0} record(s) do not have the header's field count:" -f @($v.Mismatches).Count)
            foreach ($m in @($v.Mismatches)) { Log ("      {0}" -f $m) }
            $exitCode = 1
        }
        if (-not $v.Unbalanced -and -not $v.Truncated -and @($v.Mismatches).Count -eq 0) {
            Log  "    verify=OK every record has the header's field count"
        }
    }
    if ($ContentDir) {
        Log ("    payloadsDecoded={0} payloadBytes={1:n0}" -f $r.Payloads, $r.PayloadBytes)
    }
    if ($UniqueKeyWhereCondition -and $r.KeyUnique -gt 0) {
        Log ("    key={0} values={1} unique={2} empty={3} inLists={4}" -f `
             $r.KeyColumn, $r.KeyTotal, $r.KeyUnique, $r.KeyEmpty, $r.KeyChunks)
        Log ("    where={0}" -f $whr)
        if ($r.KeyTotal -ne $r.KeyUnique) {
            Log ("    NOTE {0} duplicate key value(s): a repeated id in one INDX is itself a defect" -f `
                 ($r.KeyTotal - $r.KeyUnique))
            if ($exitCode -eq 0) { $exitCode = 2 }
        }
        if ($r.KeyEmpty -gt 0) {
            Log ("    NOTE {0} document(s) have no value for the key" -f $r.KeyEmpty)
            if ($exitCode -eq 0) { $exitCode = 2 }
        }
    }
    if ($r.Quoted -gt 0) {
        Log ("    NOTE {0} value(s) contain the separator and were quoted" -f $r.Quoted)
        Log  "    NOTE the legacy reader does not understand quoting, so it could not have produced these rows"
        if ($exitCode -eq 0) { $exitCode = 2 }
    }
    if ($null -eq $allColumns) { $allColumns = $r.Columns2; $allColFull = $r.ColFull }
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
            # The CSV header carries the local name; the mapping points it back at
            # the fully qualified element the template expects.
            $full = if ($allColFull -and $allColFull.ContainsKey($c)) { $allColFull[$c] } else { $c }
            $p.Write(("{0}.tagNameMapping.{1}={2}`n" -f $FamilyType, $c, $full))
        }
    }
    finally { $p.Dispose() }
    Log ("properties={0}" -f $PropertiesOut)
}

Log ("END exit={0}" -f $exitCode)
exit $exitCode
