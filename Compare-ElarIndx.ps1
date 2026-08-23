<#
.SYNOPSIS
    Compares a generated INDX against a reference INDX that ELAR has accepted, record
    by record, to certify that the generator produces equivalent output.

.DESCRIPTION
    The reference file is the ground truth: it was delivered and accepted. The
    candidate is produced by a new implementation and is never sent anywhere - this
    is a test of the generator, not of the data.

    Comparison is **semantic, not byte-level**, because byte equality is unattainable
    by construction: the filename counter derives from wall-clock time, the line
    wrapping width and break positions differ, and correct XML escaping differs from
    the reference wherever the reference was wrong. What is compared instead:

      - the set of record keys, matched by value rather than by position, so a
        different document order is not reported as a difference
      - per record, the set of elements and their values
      - per record, the order of elements, reported separately: the receiving schema
        is sequence-sensitive, so a reordering is a real finding even when every
        value matches
      - per record, the payload compared as the SHA-256 of the **decoded** bytes, so
        differences in Base64 line wrapping are correctly ignored while a single
        changed byte is caught

    Value differences are classified. A difference that disappears once line breaks
    are normalized is reported as WHITESPACE rather than VALUE: the reference carries
    known line-break corruption inside metadata values, so a candidate that emits the
    value cleanly is better than the reference, not different from it. Read those
    findings as expected improvements, not regressions.

    Memory is bounded by the metadata, not by file size: payloads are hashed in
    streaming and never held. Two 500 MB files with a few thousand records cost a few
    MB.

.PARAMETER ReferencePath
    The INDX that ELAR accepted.

.PARAMETER CandidatePath
    The generated INDX under test.

.PARAMETER KeyTag
    Element holding the unique record key, e.g. UniqueReportID or RecordId. With or
    without a prefix; without one, any prefix matches.

.PARAMETER IgnoreElements
    Comma-separated elements to exclude from the comparison, for values that are
    legitimately expected to differ.

.PARAMETER ContentElement
    Local name of the payload element. Default Content.

.PARAMETER DocElement
    Local name of the element delimiting one document. Default Doc.

.PARAMETER Encoding
    Charset used to read both files. Default windows-1252.

.PARAMETER MaskValues
    Withhold the values and print only element names and difference kinds. The diff
    shows values by default, because seeing them is the point of the exercise, but
    they contain personal data: use this switch when the output is to be shared or
    attached to a ticket.

.PARAMETER MaxValueChars
    Truncate printed values beyond this length. Default 200.

.PARAMETER CsvOut
    Write the findings to a CSV.

.PARAMETER MaxReport
    Maximum findings printed. Counts remain exact. Default 50.

.EXAMPLE
    .\Compare-ElarIndx.ps1 -ReferencePath .\accepted.INDX.C113925 `
                           -CandidatePath .\generated.INDX.C120000 `
                           -KeyTag UniqueReportID

.NOTES
    Windows PowerShell 5.1 and PowerShell 7 compatible. Read-only on both files.
    Exit codes: 0 equivalent, 2 differences found, 1 error.

    INSTRUMENTATION (this revision only, no behaviour change): the summary reports
    diffValueReclassifiableOnRemoval, the count of VALUE findings that would match if
    line breaks were removed instead of replaced with a space, broken down by element
    name. It measures how much of the current VALUE noise is the known line-break
    corruption rather than a real difference. Classification, findings, and exit codes
    are unchanged by this revision.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $ReferencePath,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $CandidatePath,

    [Parameter(Mandatory = $true, Position = 2)]
    [string] $KeyTag,

    [string] $IgnoreElements,

    [string] $ContentElement = 'Content',

    [string] $DocElement = 'Doc',

    [string] $Encoding = 'windows-1252',

    [switch] $MaskValues,

    [string] $CsvOut,

    [int] $MaxReport = 50,

    [int] $MaxValueChars = 200
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

foreach ($p in @($ReferencePath, $CandidatePath)) {
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { throw "File not found: $p" }
}

$ignored = @()
if ($IgnoreElements) {
    $ignored = @($IgnoreElements -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-Ignored {
    param([string] $Name)
    foreach ($ig in $ignored) {
        if ($Name -eq $ig -or $Name -like ('*:' + $ig)) { return $true }
    }
    return $false
}

<#
    Renders a value for the diff: line breaks made visible, long values truncated,
    and the whole thing withheld when -MaskValues is set.
#>
function Format-Value {
    param([string] $Value)

    if ($null -eq $Value) { return '(null)' }
    if ($Value -eq '')    { return '(empty)' }

    if ($MaskValues) {
        return ('({0} char(s), masked)' -f $Value.Length)
    }

    $v = $Value -replace "`r`n", '<CRLF>' -replace "`n", '<LF>' -replace "`r", '<CR>'
    if ($v.Length -gt $MaxValueChars) {
        $v = $v.Substring(0, $MaxValueChars) + ('... (+{0} char(s))' -f ($Value.Length - $MaxValueChars))
    }
    return $v
}

# Index of the first character that differs, so the eye goes straight to it on
# values that are long and nearly identical.
function Get-FirstDiff {
    param([string] $A, [string] $B)

    if ($null -eq $A -or $null -eq $B) { return -1 }
    $n = [Math]::Min($A.Length, $B.Length)
    for ($i = 0; $i -lt $n; $i++) {
        if ($A[$i] -cne $B[$i]) { return $i }
    }
    if ($A.Length -ne $B.Length) { return $n }
    return -1
}

function Get-Local {
    param([string] $Name)
    $i = $Name.IndexOf(':')
    if ($i -ge 0) { return $Name.Substring($i + 1) }
    return $Name
}

# ---------------------------------------------------------------------------

<#
    Streams one INDX and returns key -> record. A record holds the ordered element
    names, the values, and the payload digest. Payload bytes are hashed as they are
    decoded and never retained.
#>
function Read-IndxRecords {
    param(
        [string]               $Path,
        [System.Text.Encoding] $Enc,
        [string]               $Label
    )

    $file = Get-Item -LiteralPath $Path
    $map  = @{}
    $dupKeys = New-Object 'System.Collections.Generic.List[string]'
    $noKey = 0; $docs = 0

    $fs = $null; $sr = $null; $reader = $null

    try {
        $fs = New-Object System.IO.FileStream($file.FullName,
                  [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite, 1048576)
        # Explicit charset: the declaration in these files is not reliable.
        $sr = New-Object System.IO.StreamReader($fs, $Enc, $false, 1048576)

        $set = New-Object System.Xml.XmlReaderSettings
        $set.DtdProcessing    = [System.Xml.DtdProcessing]::Prohibit
        $set.IgnoreWhitespace = $true
        $set.IgnoreComments   = $true
        $reader = [System.Xml.XmlReader]::Create($sr, $set)

        $cur = $null; $order = $null; $pending = $null
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $nextTick = 10.0

        while ($reader.Read()) {

            switch ($reader.NodeType) {

                ([System.Xml.XmlNodeType]::Element) {
                    $local = $reader.LocalName
                    $name  = $reader.Name

                    if ($local -eq $DocElement) {
                        $docs++
                        $cur = @{}; $order = New-Object 'System.Collections.Generic.List[string]'
                        $pending = $null

                        if ($sw.Elapsed.TotalSeconds -ge $nextTick) {
                            $nextTick = $sw.Elapsed.TotalSeconds + 10
                            $pct = if ($file.Length -gt 0) { ($fs.Position / [double]$file.Length) * 100 } else { 0 }
                            Log ("    [{0}] {1:n0}% {2:n0} documents" -f $Label, $pct, $docs)
                        }
                        break
                    }

                    if ($null -eq $cur) { break }

                    if ($local -eq $ContentElement) {
                        $order.Add($name)
                        if ($reader.IsEmptyElement) { $cur[$name] = '#EMPTY'; break }

                        $sha = [System.Security.Cryptography.SHA256]::Create()
                        try {
                            $buf = New-Object byte[] 65536
                            $len = 0L
                            $n = 0
                            while (($n = $reader.ReadElementContentAsBase64($buf, 0, $buf.Length)) -gt 0) {
                                [void]$sha.TransformBlock($buf, 0, $n, $null, 0)
                                $len += $n
                            }
                            [void]$sha.TransformFinalBlock((New-Object byte[] 0), 0, 0)
                            $hex = ($sha.Hash | ForEach-Object { $_.ToString('x2') }) -join ''
                            # Length is carried too: it localises a difference faster
                            # than a digest alone.
                            $cur[$name] = ('sha256:{0} bytes:{1}' -f $hex, $len)
                        }
                        finally { $sha.Dispose() }
                        break
                    }

                    if ($reader.IsEmptyElement) {
                        $order.Add($name); $cur[$name] = ''; $pending = $null
                    }
                    else { $pending = $name }
                    break
                }

                ([System.Xml.XmlNodeType]::Text) {
                    if ($null -ne $cur -and $null -ne $pending) {
                        $order.Add($pending); $cur[$pending] = $reader.Value; $pending = $null
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::CDATA) {
                    if ($null -ne $cur -and $null -ne $pending) {
                        $order.Add($pending); $cur[$pending] = $reader.Value; $pending = $null
                    }
                    break
                }

                ([System.Xml.XmlNodeType]::EndElement) {
                    if ($reader.LocalName -eq $DocElement -and $null -ne $cur) {
                        $keyName = @($cur.Keys | Where-Object {
                                        $_ -eq $KeyTag -or $_ -like ('*:' + $KeyTag)
                                    }) | Select-Object -First 1
                        $key = if ($keyName) { [string]$cur[$keyName] } else { '' }

                        if ([string]::IsNullOrWhiteSpace($key)) { $noKey++ }
                        elseif ($map.ContainsKey($key)) {
                            if ($dupKeys.Count -lt 20) { [void]$dupKeys.Add($key) }
                        }
                        else {
                            $map[$key] = [pscustomobject]@{ Values = $cur; Order = $order }
                        }
                        $cur = $null; $order = $null; $pending = $null
                    }
                    else {
                        if ($null -ne $cur -and $null -ne $pending -and $reader.Name -eq $pending) {
                            $order.Add($pending); $cur[$pending] = ''
                        }
                        $pending = $null
                    }
                    break
                }
            }
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
        elseif ($sr) { $sr.Dispose() }
        elseif ($fs) { $fs.Dispose() }
    }

    return [pscustomobject]@{
        Map = $map; Documents = $docs; NoKey = $noKey; DuplicateKeys = $dupKeys
        SizeMB = [Math]::Round($file.Length / 1MB, 1); Name = $file.Name
    }
}

# ---------------------------------------------------------------------------

$exitCode = 0
$findings = New-Object 'System.Collections.Generic.List[psobject]'

try {
    Log ("START key={0} charset={1}" -f $KeyTag, $Encoding)
    if ($ignored.Count -gt 0) { Log ("  ignoring elements: {0}" -f ($ignored -join ', ')) }

    Log ("[reference] {0}" -f (Split-Path -Leaf $ReferencePath))
    $ref = Read-IndxRecords -Path $ReferencePath -Enc $enc -Label 'reference'
    Log ("    documents={0} keys={1} noKey={2}" -f $ref.Documents, $ref.Map.Count, $ref.NoKey)

    Log ("[candidate] {0}" -f (Split-Path -Leaf $CandidatePath))
    $cnd = Read-IndxRecords -Path $CandidatePath -Enc $enc -Label 'candidate'
    Log ("    documents={0} keys={1} noKey={2}" -f $cnd.Documents, $cnd.Map.Count, $cnd.NoKey)

    foreach ($pair in @(@{ N = 'reference'; R = $ref }, @{ N = 'candidate'; R = $cnd })) {
        if (@($pair.R.DuplicateKeys).Count -gt 0) {
            Log ("WARNING {0} has duplicate key value(s): {1}" -f $pair.N, (@($pair.R.DuplicateKeys) -join ','))
            Log  "WARNING only the first occurrence of each key was kept"
            $exitCode = 2
        }
        if ($pair.R.NoKey -gt 0) {
            Log ("WARNING {0} has {1} document(s) with no key; they cannot be matched" -f $pair.N, $pair.R.NoKey)
            $exitCode = 2
        }
    }

    # --- key sets ---
    $onlyRef = @($ref.Map.Keys | Where-Object { -not $cnd.Map.ContainsKey($_) })
    $onlyCnd = @($cnd.Map.Keys | Where-Object { -not $ref.Map.ContainsKey($_) })
    $common  = @($ref.Map.Keys | Where-Object { $cnd.Map.ContainsKey($_) })

    foreach ($k in $onlyRef) {
        $findings.Add([pscustomobject]@{ Key = $k; Kind = 'MISSING_IN_CANDIDATE'; Element = ''; Reference = ''; Candidate = '' })
    }
    foreach ($k in $onlyCnd) {
        $findings.Add([pscustomobject]@{ Key = $k; Kind = 'EXTRA_IN_CANDIDATE'; Element = ''; Reference = ''; Candidate = '' })
    }

    # --- per record ---
    $identical = 0; $differing = 0
    $countValue = 0; $countWs = 0; $countMissing = 0; $countExtra = 0; $countContent = 0; $countOrder = 0

    # --- instrumentation only, no behaviour change -------------------------
    # How many of today's VALUE findings would become WHITESPACE if line breaks
    # were REMOVED rather than replaced with a space. The legacy writer inserted
    # a break into a value ("P" + LF + "DF" for "PDF"), so removal is the
    # faithful inverse and substitution invents a character that was never there.
    $reclass = 0
    $reclassByElement = @{}
    # -----------------------------------------------------------------------

    foreach ($k in $common) {
        $a = $ref.Map[$k]; $b = $cnd.Map[$k]
        $diffHere = 0

        $namesA = @($a.Values.Keys | Where-Object { -not (Test-Ignored $_) })
        $namesB = @($b.Values.Keys | Where-Object { -not (Test-Ignored $_) })

        foreach ($n in $namesA) {
            if (-not $b.Values.ContainsKey($n)) {
                $findings.Add([pscustomobject]@{ Key = $k; Kind = 'ELEMENT_MISSING'; Element = $n; Reference = [string]$a.Values[$n]; Candidate = '' })
                $countMissing++; $diffHere++
                continue
            }

            $va = [string]$a.Values[$n]
            $vb = [string]$b.Values[$n]
            if ($va -ceq $vb) { continue }

            $isContent = ((Get-Local $n) -eq $ContentElement)
            if ($isContent) {
                $findings.Add([pscustomobject]@{ Key = $k; Kind = 'CONTENT_MISMATCH'; Element = $n; Reference = $va; Candidate = $vb })
                $countContent++; $diffHere++
                continue
            }

            # A difference that vanishes once line breaks are normalized is the known
            # corruption in the reference, not a fault in the candidate.
            $na = ($va -replace "[`r`n]+", ' ').Trim()
            $nb = ($vb -replace "[`r`n]+", ' ').Trim()
            if ($na -ceq $nb) {
                $findings.Add([pscustomobject]@{ Key = $k; Kind = 'WHITESPACE'; Element = $n; Reference = $va; Candidate = $vb })
                $countWs++
            }
            else {
                $findings.Add([pscustomobject]@{ Key = $k; Kind = 'VALUE'; Element = $n; Reference = $va; Candidate = $vb })
                $countValue++; $diffHere++

                # Instrumentation: would removal have matched them?
                if (($va -replace "[`r`n]+", '') -ceq ($vb -replace "[`r`n]+", '')) {
                    $reclass++
                    $el = Get-Local $n
                    if ($reclassByElement.ContainsKey($el)) { $reclassByElement[$el]++ }
                    else { $reclassByElement[$el] = 1 }
                }
            }
        }

        foreach ($n in $namesB) {
            if (-not $a.Values.ContainsKey($n)) {
                $findings.Add([pscustomobject]@{ Key = $k; Kind = 'ELEMENT_EXTRA'; Element = $n; Reference = ''; Candidate = [string]$b.Values[$n] })
                $countExtra++; $diffHere++
            }
        }

        # Order matters to the receiving schema, so it is checked even when every
        # value matches.
        $oa = @($a.Order | Where-Object { -not (Test-Ignored $_) })
        $ob = @($b.Order | Where-Object { -not (Test-Ignored $_) })
        if (($oa -join '|') -cne ($ob -join '|') -and $countMissing -eq 0 -and $countExtra -eq 0) {
            $findings.Add([pscustomobject]@{ Key = $k; Kind = 'ORDER'; Element = ''; Reference = ($oa -join ','); Candidate = ($ob -join ',') })
            $countOrder++; $diffHere++
        }

        if ($diffHere -eq 0) { $identical++ } else { $differing++ }
    }

    # --- report ---
    Log "SUMMARY"
    Log ("  referenceRecords={0}" -f $ref.Map.Count)
    Log ("  candidateRecords={0}" -f $cnd.Map.Count)
    Log ("  matched={0}" -f $common.Count)
    Log ("  missingInCandidate={0}" -f $onlyRef.Count)
    Log ("  extraInCandidate={0}" -f $onlyCnd.Count)
    Log ("  recordsIdentical={0}" -f $identical)
    Log ("  recordsDiffering={0}" -f $differing)
    Log ("  diffValue={0}" -f $countValue)
    Log ("  diffWhitespaceOnly={0}" -f $countWs)
    Log ("  diffElementMissing={0}" -f $countMissing)
    Log ("  diffElementExtra={0}" -f $countExtra)
    Log ("  diffContent={0}" -f $countContent)
    Log ("  diffOrder={0}" -f $countOrder)
    Log ("  diffValueReclassifiableOnRemoval={0}" -f $reclass)
    if ($reclass -gt 0) {
        foreach ($e in ($reclassByElement.GetEnumerator() | Sort-Object Value -Descending)) {
            Log ("    reclassifiable {0}={1}" -f $e.Key, $e.Value)
        }
        Log  "  NOTE these VALUE findings match once line breaks are removed rather than replaced with a space"
        Log  "  NOTE element names only, no values: this line is safe to share"
    }

    $blocking = $onlyRef.Count + $onlyCnd.Count + $countValue + $countMissing + $countExtra + $countContent + $countOrder
    if ($blocking -gt 0) { $exitCode = 2 }

    if ($findings.Count -gt 0) {
        Log ""
        Log "DIFFERENCES   - reference   + candidate"
        Log ""

        $shown = 0
        $stop  = $false

        foreach ($grp in ($findings | Group-Object Key)) {
            if ($stop) { break }

            $keyLabel = if ($grp.Name) { $grp.Name } else { '(no key)' }
            Log ("  record {0}   [{1} difference(s)]" -f $keyLabel, $grp.Count)

            foreach ($f in $grp.Group) {
                if ($shown -ge $MaxReport) {
                    Log ("  ... +{0} more finding(s); raise -MaxReport or use -CsvOut" -f ($findings.Count - $shown))
                    $stop = $true
                    break
                }
                $shown++

                switch ($f.Kind) {

                    'MISSING_IN_CANDIDATE' {
                        Log  "    - present in reference only"
                        Log  "    +"
                        break
                    }

                    'EXTRA_IN_CANDIDATE' {
                        Log  "    -"
                        Log  "    + present in candidate only"
                        break
                    }

                    'ORDER' {
                        Log  "    element order differs"
                        Log ("    - {0}" -f (Format-Value -Value $f.Reference))
                        Log ("    + {0}" -f (Format-Value -Value $f.Candidate))
                        break
                    }

                    'ELEMENT_MISSING' {
                        Log ("    {0}   [ELEMENT_MISSING]" -f (Get-Local $f.Element))
                        Log ("    - {0}" -f (Format-Value -Value $f.Reference))
                        Log  "    + (element absent)"
                        break
                    }

                    'ELEMENT_EXTRA' {
                        Log ("    {0}   [ELEMENT_EXTRA]" -f (Get-Local $f.Element))
                        Log  "    - (element absent)"
                        Log ("    + {0}" -f (Format-Value -Value $f.Candidate))
                        break
                    }

                    default {
                        # VALUE, WHITESPACE, CONTENT_MISMATCH
                        $pos = Get-FirstDiff -A $f.Reference -B $f.Candidate
                        $at  = if ($pos -ge 0) { ("  first differs at char {0}" -f ($pos + 1)) } else { '' }
                        Log ("    {0}   [{1}]{2}" -f (Get-Local $f.Element), $f.Kind, $at)
                        Log ("    - {0}" -f (Format-Value -Value $f.Reference))
                        Log ("    + {0}" -f (Format-Value -Value $f.Candidate))
                        break
                    }
                }
            }
            Log ""
        }

        if ($MaskValues) { Log "  values are masked; omit -MaskValues to see them" }
    }

    if ($CsvOut) {
        $export = if ($MaskValues) { $findings | Select-Object Key, Kind, Element } else { $findings }
        $export | Export-Csv -LiteralPath $CsvOut -NoTypeInformation -Encoding UTF8
        Log ("  report={0}" -f $CsvOut)
    }

    if ($countWs -gt 0) {
        Log "  NOTE WHITESPACE findings are values that match once line breaks are normalized"
        Log "  NOTE those are the reference's known corruption; the candidate is the better of the two"
    }
    if ($countContent -gt 0) {
        Log "  NOTE CONTENT_MISMATCH compares the digest of the decoded payload, so wrapping is not the cause"
    }
    if ($blocking -eq 0) {
        Log "  VERDICT equivalent: every record matches on keys, elements, order and payload"
    }
    else {
        Log "  VERDICT not equivalent: see the findings above"
    }

    Log ("END exit={0}" -f $exitCode)
}
catch [System.Xml.XmlException] {
    Log ("ERROR parse error at line {0} position {1}: {2}" -f `
         $_.Exception.LineNumber, $_.Exception.LinePosition, $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}
catch {
    Log ("ERROR {0}" -f $_.Exception.Message)
    Log  "END exit=1"
    exit 1
}

exit $exitCode
