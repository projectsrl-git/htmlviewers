<#
.SYNOPSIS
    Compares a key column from two CSV files and writes out only the keys that are
    not common to both.

.DESCRIPTION
    Reads one column from each file, compares the two sets, and produces a CSV with
    two columns: the name of the file the key came from, and the key itself. A key
    present in both files does not appear in the output; a key present in only one
    does, attributed to that file.

    The key column may be named differently in the two files, which is the usual case
    when comparing an extraction against a delivery.

    Comparison is on the trimmed value. Case sensitivity is off by default, matching
    how Windows and most identifier schemes behave; -CaseSensitive turns it on when
    the keys are genuinely case-bearing.

    Duplicates within one file are collapsed and counted: a repeated key is reported
    but does not multiply the output rows. That count is worth reading - a key that
    is meant to be unique and is not is a finding in its own right, independent of
    the comparison.

    Plain timestamped stdout, suitable as an OpenProteo exec step.

.PARAMETER Path1
    First CSV.

.PARAMETER Key1
    Key column in the first CSV.

.PARAMETER Path2
    Second CSV.

.PARAMETER Key2
    Key column in the second CSV. Defaults to -Key1 when the two files share the
    column name.

.PARAMETER OutFile
    Destination CSV. Defaults to keys-not-common.csv in the current directory.

.PARAMETER Delimiter
    Delimiter of the input files. Default ';'.

.PARAMETER OutDelimiter
    Delimiter of the output file. Defaults to -Delimiter.

.PARAMETER Encoding
    Charset used to read the inputs and write the output. Default windows-1252.

.PARAMETER CaseSensitive
    Compare keys case-sensitively.

.PARAMETER UseFullPath
    Write the full path in the file column instead of just the file name.

.PARAMETER MaxReport
    Maximum keys listed per side in the log. Default 20. The output file is complete
    regardless.

.EXAMPLE
    .\Compare-CsvKeys.ps1 -Path1 .\extraction.csv -Key1 UNIQUEREPORTID `
                          -Path2 .\delivered.csv  -Key2 UniqueReportID `
                          -OutFile .\missing.csv

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Read-only on both inputs.
    Exit codes: 0 the two key sets are identical, 2 differences found, 1 error.

    The exit code is deliberately non-zero when differences exist, so a workflow can
    branch on it without parsing the output.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path1,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $Key1,

    [Parameter(Mandatory = $true, Position = 2)]
    [string] $Path2,

    [string] $Key2,

    [string] $OutFile = '.\keys-not-common.csv',

    [string] $Delimiter = ';',

    [string] $OutDelimiter,

    [string] $Encoding = 'windows-1252',

    [switch] $CaseSensitive,

    [switch] $UseFullPath,

    [int] $MaxReport = 20
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

if (-not $Key2)         { $Key2 = $Key1 }
if (-not $OutDelimiter) { $OutDelimiter = $Delimiter }

$exitCode = 0

<#
    Reads one column and returns the distinct keys in order of first appearance,
    along with the counts needed to explain the result.
#>
function Read-KeyColumn {
    param(
        [string] $Path,
        [string] $Column,
        [string] $Delim,
        [string] $Enc,
        [bool]   $Sensitive
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }

    $cmp = if ($Sensitive) { [System.StringComparer]::Ordinal } else { [System.StringComparer]::OrdinalIgnoreCase }
    $set  = New-Object 'System.Collections.Generic.HashSet[string]' $cmp
    $list = New-Object 'System.Collections.Generic.List[string]'
    $rowsRead = 0; $blank = 0; $dupes = 0

    # Import-Csv understands quoting, so a value containing the delimiter is read as
    # one field instead of splitting the row.
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $Delim -Encoding ([string]$Enc) -ErrorAction Stop)
    if ($rows.Count -eq 0) { throw "No data rows in $Path" }

    $header = @($rows[0].PSObject.Properties.Name)
    if ($header -notcontains $Column) {
        throw ("Column '{0}' not found in {1}. Header is: {2}" -f $Column, (Split-Path -Leaf $Path), ($header -join ', '))
    }

    foreach ($r in $rows) {
        $rowsRead++
        $v = [string]$r.$Column
        if ([string]::IsNullOrWhiteSpace($v)) { $blank++; continue }
        $v = $v.Trim()
        if ($set.Add($v)) { $list.Add($v) } else { $dupes++ }
    }

    return [pscustomobject]@{
        Name = (Split-Path -Leaf $Path); Full = (Get-Item -LiteralPath $Path).FullName
        Set = $set; Keys = $list; Rows = $rowsRead; Blank = $blank; Duplicates = $dupes
    }
}

try {
    Log ("START key1={0} key2={1} delimiter='{2}' charset={3} caseSensitive={4}" -f `
         $Key1, $Key2, $Delimiter, $Encoding, $CaseSensitive.IsPresent)

    $a = Read-KeyColumn -Path $Path1 -Column $Key1 -Delim $Delimiter -Enc $Encoding -Sensitive $CaseSensitive.IsPresent
    Log ("  [1] {0}  rows={1:n0} distinctKeys={2:n0} blank={3} duplicates={4}" -f `
         $a.Name, $a.Rows, $a.Keys.Count, $a.Blank, $a.Duplicates)

    $b = Read-KeyColumn -Path $Path2 -Column $Key2 -Delim $Delimiter -Enc $Encoding -Sensitive $CaseSensitive.IsPresent
    Log ("  [2] {0}  rows={1:n0} distinctKeys={2:n0} blank={3} duplicates={4}" -f `
         $b.Name, $b.Rows, $b.Keys.Count, $b.Blank, $b.Duplicates)

    $label1 = if ($UseFullPath) { $a.Full } else { $a.Name }
    $label2 = if ($UseFullPath) { $b.Full } else { $b.Name }

    # Order of first appearance within each file, first file first: the output stays
    # readable against the sources instead of being sorted into an unrelated order.
    $onlyA = New-Object 'System.Collections.Generic.List[string]'
    foreach ($k in $a.Keys) { if (-not $b.Set.Contains($k)) { $onlyA.Add($k) } }

    $onlyB = New-Object 'System.Collections.Generic.List[string]'
    foreach ($k in $b.Keys) { if (-not $a.Set.Contains($k)) { $onlyB.Add($k) } }

    $common = $a.Keys.Count - $onlyA.Count

    $needsQuote = [char[]]@($OutDelimiter[0], '"', "`r", "`n")
    function Format-Field {
        param([string] $Value)
        if ([string]::IsNullOrEmpty($Value)) { return '' }
        if ($Value.IndexOfAny($needsQuote) -ge 0) { return '"' + $Value.Replace('"', '""') + '"' }
        return $Value
    }

    try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
    catch { throw "Unknown charset '$Encoding'." }

    $outDir = Split-Path -Parent $OutFile
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $w = New-Object System.IO.StreamWriter($OutFile, $false, $enc)
    try {
        $w.Write('file'); $w.Write($OutDelimiter); $w.Write('key'); $w.Write("`n")
        foreach ($k in $onlyA) {
            $w.Write((Format-Field $label1)); $w.Write($OutDelimiter); $w.Write((Format-Field $k)); $w.Write("`n")
        }
        foreach ($k in $onlyB) {
            $w.Write((Format-Field $label2)); $w.Write($OutDelimiter); $w.Write((Format-Field $k)); $w.Write("`n")
        }
    }
    finally { $w.Dispose() }

    Log "SUMMARY"
    Log ("  commonKeys={0:n0}" -f $common)
    Log ("  onlyIn[1] {0}={1:n0}" -f $a.Name, $onlyA.Count)
    Log ("  onlyIn[2] {0}={1:n0}" -f $b.Name, $onlyB.Count)
    Log ("  writtenRows={0:n0}" -f ($onlyA.Count + $onlyB.Count))
    Log ("  out={0}" -f $OutFile)

    if ($onlyA.Count -gt 0) {
        $shown = @($onlyA | Select-Object -First $MaxReport) -join ', '
        $more  = if ($onlyA.Count -gt $MaxReport) { (' ... +{0} more' -f ($onlyA.Count - $MaxReport)) } else { '' }
        Log ("  onlyIn[1]: {0}{1}" -f $shown, $more)
    }
    if ($onlyB.Count -gt 0) {
        $shown = @($onlyB | Select-Object -First $MaxReport) -join ', '
        $more  = if ($onlyB.Count -gt $MaxReport) { (' ... +{0} more' -f ($onlyB.Count - $MaxReport)) } else { '' }
        Log ("  onlyIn[2]: {0}{1}" -f $shown, $more)
    }

    if ($a.Duplicates -gt 0 -or $b.Duplicates -gt 0) {
        Log "  NOTE duplicate keys were collapsed; if the column is meant to be unique, that is a finding of its own"
    }
    if (($onlyA.Count + $onlyB.Count) -eq 0) {
        Log "  VERDICT the two key sets are identical"
    }
    else {
        Log "  VERDICT the key sets differ"
        $exitCode = 2
    }

    Log ("END exit={0}" -f $exitCode)
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
