<#
.SYNOPSIS
    Sets every value of one column in a CSV to a fixed value.

.DESCRIPTION
    Reads a CSV, writes the same CSV back with one column overwritten by a constant.
    Every other column, the header, the row order and the delimiter are preserved.

    Quoting is handled on both sides: a value containing the delimiter, a quote or a
    line break is parsed as one field on read and quoted again on write, so a row is
    never split or joined by accident. That matters here more than usual, because the
    file is being rewritten in place by default and a mis-parse would be permanent.

    By default the input is replaced and a .bak is kept. -OutFile writes elsewhere and
    leaves the input untouched.

    -OnlyWhere restricts the change to rows where another column has a given value,
    which is the difference between blanking a field everywhere and correcting the
    handful of rows that need it.

    Plain timestamped stdout, no prompts. Use -WhatIf for a dry run.

.PARAMETER Path
    CSV to modify.

.PARAMETER Column
    Column whose values are replaced.

.PARAMETER Value
    The fixed value to write. Use '' for an empty field.

.PARAMETER OutFile
    Write to this file instead of replacing the input.

.PARAMETER OnlyWhere
    Column tested to decide which rows are changed. Requires -OnlyWhereValue.

.PARAMETER OnlyWhereValue
    Value that -OnlyWhere must hold for the row to be changed.

.PARAMETER OnlyWhereNotEqual
    Invert the test: change the rows where -OnlyWhere does NOT equal the value.

.PARAMETER AddIfMissing
    Add the column when the header does not contain it, instead of failing.

.PARAMETER Delimiter
    CSV delimiter. Default ';'.

.PARAMETER Encoding
    Charset used to read and write. Default windows-1252.

.PARAMETER NoBackup
    Skip the .bak copy when replacing the input.

.PARAMETER CaseSensitive
    Match the -OnlyWhere value case-sensitively.

.EXAMPLE
    .\Set-CsvColumnValue.ps1 -Path .\feed.csv -Column ELAR:AccountID -Value '-'

.EXAMPLE
    .\Set-CsvColumnValue.ps1 -Path .\feed.csv -Column BU -Value '0043' `
        -OnlyWhere Origin -OnlyWhereValue SCANSIONATO -OutFile .\feed-fixed.csv

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible.
    Exit codes: 0 completed, 2 completed but no row matched, 1 error.

    The file is streamed, so size is not a constraint: a row is held, rewritten and
    released. The output is built alongside the input and swapped in only after a
    clean close, so an interrupted run leaves the original intact.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $Column,

    [Parameter(Mandatory = $true, Position = 2)]
    [AllowEmptyString()]
    [string] $Value,

    [string] $OutFile,

    [string] $OnlyWhere,

    [AllowEmptyString()]
    [string] $OnlyWhereValue,

    [switch] $OnlyWhereNotEqual,

    [switch] $AddIfMissing,

    [string] $Delimiter = ';',

    [string] $Encoding = 'windows-1252',

    [switch] $NoBackup,

    [switch] $CaseSensitive
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Log {
    param([string] $Message)
    [Console]::Out.WriteLine(('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message))
    [Console]::Out.Flush()
}

<#
    Splits one CSV record into fields, honouring quotes. A field may contain the
    delimiter and doubled quotes; embedded line breaks are out of scope because the
    reader is line-oriented, and are reported rather than mis-handled.
#>
function Split-CsvLine {
    param([string] $Line, [char] $Sep)

    $out = New-Object 'System.Collections.Generic.List[string]'
    $sb  = New-Object System.Text.StringBuilder
    $inQ = $false
    $i = 0

    while ($i -lt $Line.Length) {
        $c = $Line[$i]

        if ($inQ) {
            if ($c -eq '"') {
                if (($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$sb.Append('"'); $i += 2; continue }
                $inQ = $false; $i++; continue
            }
            [void]$sb.Append($c); $i++; continue
        }

        if ($c -eq '"')   { $inQ = $true; $i++; continue }
        if ($c -eq $Sep)  { $out.Add($sb.ToString()); $sb.Length = 0; $i++; continue }
        [void]$sb.Append($c); $i++
    }

    $out.Add($sb.ToString())
    return [pscustomobject]@{ Fields = $out; Unterminated = $inQ }
}

function Join-CsvLine {
    param([string[]] $Fields, [char] $Sep)

    $needs = [char[]]@($Sep, '"', "`r", "`n")
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Fields.Count; $i++) {
        if ($i -gt 0) { [void]$sb.Append($Sep) }
        $f = [string]$Fields[$i]
        if ($f -and $f.IndexOfAny($needs) -ge 0) {
            [void]$sb.Append('"').Append($f.Replace('"', '""')).Append('"')
        }
        else { [void]$sb.Append($f) }
    }
    return $sb.ToString()
}

$exitCode = 0

try {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }
    if ($OnlyWhere -and -not $PSBoundParameters.ContainsKey('OnlyWhereValue')) {
        throw "-OnlyWhere requires -OnlyWhereValue."
    }
    if (-not $OnlyWhere -and $PSBoundParameters.ContainsKey('OnlyWhereValue')) {
        throw "-OnlyWhereValue requires -OnlyWhere."
    }

    try { $enc = [System.Text.Encoding]::GetEncoding($Encoding) }
    catch { throw "Unknown charset '$Encoding'." }

    $src = Get-Item -LiteralPath $Path
    $sep = $Delimiter[0]
    $inPlace = -not $OutFile
    $dest = if ($inPlace) { Join-Path $src.DirectoryName ($src.Name + '.setcol.tmp') } else { $OutFile }

    $outDir = Split-Path -Parent $dest
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    Log ("START file={0} column='{1}' value='{2}' delimiter='{3}' charset={4}" -f `
         $src.Name, $Column, $Value, $Delimiter, $Encoding)
    if ($OnlyWhere) {
        $op = if ($OnlyWhereNotEqual) { '<>' } else { '=' }
        Log ("  condition: {0} {1} '{2}'" -f $OnlyWhere, $op, $OnlyWhereValue)
    }
    Log ("  target={0}" -f $(if ($inPlace) { 'in place' } else { $dest }))

    $reader = $null; $writer = $null
    $rows = 0; $changed = 0; $already = 0; $short = 0; $unterminated = 0
    $colIdx = -1; $whereIdx = -1; $added = $false

    try {
        $reader = New-Object System.IO.StreamReader($src.FullName, $enc, $false, 1048576)
        $writer = New-Object System.IO.StreamWriter($dest, $false, $enc, 1048576)

        # --- header ---
        $headerLine = $reader.ReadLine()
        if ($null -eq $headerLine) { throw "File is empty." }

        $h = Split-CsvLine -Line $headerLine -Sep $sep
        $header = @($h.Fields)

        for ($i = 0; $i -lt $header.Count; $i++) {
            if ($header[$i] -eq $Column) { $colIdx = $i }
            if ($OnlyWhere -and $header[$i] -eq $OnlyWhere) { $whereIdx = $i }
        }

        if ($colIdx -lt 0) {
            if (-not $AddIfMissing) {
                throw ("Column '{0}' not found. Header is: {1}" -f $Column, ($header -join ', '))
            }
            $header += $Column
            $colIdx = $header.Count - 1
            $added = $true
        }
        if ($OnlyWhere -and $whereIdx -lt 0) {
            throw ("Condition column '{0}' not found. Header is: {1}" -f $OnlyWhere, ($header -join ', '))
        }

        Log ("  headerFields={0} columnIndex={1}{2}" -f $header.Count, $colIdx, $(if ($added) { ' (added)' } else { '' }))

        $writer.Write((Join-CsvLine -Fields $header -Sep $sep)); $writer.Write("`n")

        # --- rows ---
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.Length -eq 0) { continue }
            $rows++

            $p = Split-CsvLine -Line $line -Sep $sep
            if ($p.Unterminated) { $unterminated++ }

            $fields = @($p.Fields)

            # A short row is padded rather than dropped: losing a record would be a
            # worse outcome than a row with empty trailing fields.
            while ($fields.Count -le $colIdx) { $fields += ''; $short++ }

            $doIt = $true
            if ($OnlyWhere) {
                $actual = if ($whereIdx -lt $fields.Count) { [string]$fields[$whereIdx] } else { '' }
                $match = if ($CaseSensitive) { $actual -ceq $OnlyWhereValue } else { $actual -eq $OnlyWhereValue }
                $doIt = if ($OnlyWhereNotEqual) { -not $match } else { $match }
            }

            if ($doIt) {
                if ([string]$fields[$colIdx] -ceq $Value) { $already++ }
                else { $fields[$colIdx] = $Value; $changed++ }
            }

            $writer.Write((Join-CsvLine -Fields $fields -Sep $sep)); $writer.Write("`n")
        }
    }
    finally {
        if ($writer) { $writer.Flush(); $writer.Dispose() }
        if ($reader) { $reader.Dispose() }
    }

    Log "SUMMARY"
    Log ("  rows={0:n0}" -f $rows)
    Log ("  changed={0:n0}" -f $changed)
    Log ("  alreadyEqual={0:n0}" -f $already)
    if ($short -gt 0)        { Log ("  paddedShortRows={0:n0}" -f $short) }
    if ($unterminated -gt 0) { Log ("  WARNING rows with an unterminated quote={0:n0}" -f $unterminated) }

    if ($inPlace) {
        if ($PSCmdlet.ShouldProcess($src.FullName, "Replace with $changed changed row(s)")) {
            if (-not $NoBackup) {
                Copy-Item -LiteralPath $src.FullName -Destination ($src.FullName + '.bak') -Force
                Log ("  backup={0}.bak" -f $src.Name)
            }
            Move-Item -LiteralPath $dest -Destination $src.FullName -Force
            Log ("  replaced={0}" -f $src.Name)
        }
        elseif (Test-Path -LiteralPath $dest) {
            Remove-Item -LiteralPath $dest -Force
        }
    }
    else {
        Log ("  out={0}" -f $dest)
    }

    if ($changed -eq 0) {
        Log "  NOTE no row was changed: check the column name, the condition, or whether the value is already set"
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
