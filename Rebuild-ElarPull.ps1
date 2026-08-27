<#
.SYNOPSIS
    Rebuilds the PULL half of already-delivered ELAR INDX/PULL pairs.

.DESCRIPTION
    One PULL per INDX found in -IndxDir. The PULL name is the INDX name with the LAST occurrence of
    INDX replaced by PULL; the name written INSIDE the PULL is the INDX file name with a trailing
    literal ".xml" removed and nothing else - the same two rules the executor uses
    (BatchNaming.partName and BatchNaming.stripExtension). Removing "the last dot-segment" instead
    would eat the .C103800 counter and produce a PULL referencing a file that does not exist, while
    both files still look like a pair in a directory listing.

    Two sources for the body:

      -ReferencePull   PREFERRED when a known-good PULL of the same family exists. Its text is
                       cloned and the index name it already carries is swapped for the new one, so
                       the result is byte-identical to its siblings except for the one string that
                       has to differ.

      -PullTemplate    The template with the [INDEX_NAME] placeholder. Correct, but note that the
                       executor PARSES the template and re-serialises it, generating the XML
                       declaration from the output charset rather than copying it. A text
                       substitution therefore reproduces the TEMPLATE's declaration, attribute order
                       and whitespace, which may not match what the executor emitted.

    Output is UTF-8 WITHOUT a BOM and LF-only, because that is what WrappingXmlOut writes
    (NL = char 10, never CRLF; declaration generated from the charset). Set-Content and Out-File get
    both of these wrong on Windows PowerShell 5.1, which is why every write here goes through
    [System.IO.File]::WriteAllText.

    Nothing is written without -Apply, and an existing target file is never overwritten.

.EXAMPLE
    .\Rebuild-ElarPull.ps1 -IndxDir D:\elar\out -ReferencePull D:\elar\good\...PULL.C101500 -OutDir D:\elar\rebuilt
    .\Rebuild-ElarPull.ps1 -IndxDir D:\elar\out -ReferencePull D:\elar\good\...PULL.C101500 -OutDir D:\elar\rebuilt -Apply
#>
[CmdletBinding(DefaultParameterSetName = 'Template')]
param(
    [Parameter(Mandatory = $true)]
    [string] $IndxDir,

    [Parameter(Mandatory = $true, ParameterSetName = 'Template')]
    [string] $PullTemplate,

    [Parameter(Mandatory = $true, ParameterSetName = 'Reference')]
    [string] $ReferencePull,

    [Parameter(Mandatory = $true)]
    [string] $OutDir,

    [string] $IndxFilter = '*INDX*',

    [switch] $Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$LF          = [string][char]10
$PLACEHOLDER = '[INDEX_NAME]'
$UTF8_NO_BOM = New-Object System.Text.UTF8Encoding($false)

function Read-TextLf([string] $path) {
    # ReadAllText, not Get-Content: Get-Content splits into lines and loses the trailing newline.
    # UTF8 here also strips a BOM if the file has one, so a BOM on input cannot reach the output.
    $t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    return $t.Replace([string][char]13 + $LF, $LF).Replace([string][char]13, $LF)
}

function Replace-LastToken([string] $name, [string] $from, [string] $to) {
    $i = $name.LastIndexOf($from, [System.StringComparison]::Ordinal)
    if ($i -lt 0) { return $null }
    return $name.Substring(0, $i) + $to + $name.Substring($i + $from.Length)
}

function Strip-XmlExtension([string] $name) {
    # ONLY a trailing literal .xml, exactly as BatchNaming.stripExtension does.
    if ($name.Length -gt 4 -and $name.Substring($name.Length - 4).ToLowerInvariant() -eq '.xml') {
        return $name.Substring(0, $name.Length - 4)
    }
    return $name
}

# ---------------------------------------------------------------- inputs

if (-not (Test-Path -LiteralPath $IndxDir -PathType Container)) {
    throw "INDX directory not found: $IndxDir"
}
$indxFull = (Resolve-Path -LiteralPath $IndxDir).ProviderPath
$outFull  = [System.IO.Path]::GetFullPath($OutDir)
if ($outFull.TrimEnd('\') -ieq $indxFull.TrimEnd('\')) {
    throw "OutDir must not be the INDX directory. Write the rebuilt files somewhere else, look at them, then move the old PULLs into a quarantine folder and copy the new ones in. Never overwrite a delivered file in place."
}

$body = $null
$refIndexName = $null

if ($PSCmdlet.ParameterSetName -eq 'Reference') {
    if (-not (Test-Path -LiteralPath $ReferencePull -PathType Leaf)) {
        throw "Reference PULL not found: $ReferencePull"
    }
    $body = Read-TextLf $ReferencePull
    $refName = [System.IO.Path]::GetFileName($ReferencePull)
    $refIndxFileName = Replace-LastToken $refName 'PULL' 'INDX'
    if ($null -eq $refIndxFileName) {
        throw "The reference file name does not contain PULL, so the index name it should carry cannot be derived: $refName"
    }
    $refIndexName = Strip-XmlExtension $refIndxFileName
    # A reference that does not name its own INDX is a broken pair, and cloning it would spread the
    # break to 22 files. This is the one check that makes the reference mode safe.
    if ($body.IndexOf($refIndexName, [System.StringComparison]::Ordinal) -lt 0) {
        throw "The reference PULL never names its own INDX ($refIndexName), so it is itself a broken pair and must not be used as a source."
    }
    Write-Host "Source     : reference PULL $refName"
    Write-Host "  it names : $refIndexName"
}
else {
    if (-not (Test-Path -LiteralPath $PullTemplate -PathType Leaf)) {
        throw "PULL template not found: $PullTemplate"
    }
    $body = Read-TextLf $PullTemplate
    # A template without the placeholder would produce 22 identical files, all naming nothing. Refuse.
    if ($body.IndexOf($PLACEHOLDER, [System.StringComparison]::Ordinal) -lt 0) {
        throw "The template does not contain $PLACEHOLDER anywhere, so every file produced would be identical and would name no INDX at all: $PullTemplate"
    }
    Write-Host "Source     : template $([System.IO.Path]::GetFileName($PullTemplate))"
    $firstLine = ($body -split $LF)[0]
    Write-Host "  its first line: $firstLine"
    if ($firstLine -notmatch '^<\?xml\s+version="1\.0"\s+encoding="UTF-8"\?>$') {
        Write-Warning "The executor GENERATES the declaration as <?xml version=`"1.0`" encoding=`"UTF-8`"?> from the output charset and never copies it from the template. The line above will be reproduced verbatim instead. Check it is what you want before -Apply."
    }
}

Write-Host "INDX dir   : $indxFull  (filter $IndxFilter)"
Write-Host "Out dir    : $outFull"
Write-Host ("Mode       : " + $(if ($Apply) { 'APPLY - files will be written' } else { 'DRY RUN - nothing will be written' }))
Write-Host ''

# ---------------------------------------------------------------- work

$candidates = @(Get-ChildItem -LiteralPath $indxFull -File -Filter $IndxFilter |
                Sort-Object -Property Name)

$planned = 0
$skipped = 0

foreach ($f in $candidates) {

    if ($f.Name.EndsWith('.part', [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Host "  SKIP  $($f.Name)  - a file still in flight, not a delivered INDX"
        $skipped++
        continue
    }

    $pullFileName = Replace-LastToken $f.Name 'INDX' 'PULL'
    if ($null -eq $pullFileName) {
        Write-Host "  SKIP  $($f.Name)  - no INDX token in the name"
        $skipped++
        continue
    }

    $indexName = Strip-XmlExtension $f.Name

    if ($PSCmdlet.ParameterSetName -eq 'Reference') {
        $text = $body.Replace($refIndexName, $indexName)
    }
    else {
        $text = $body.Replace($PLACEHOLDER, $indexName)
    }

    # The whole point of the file is that it names its INDX. Assert it rather than assume it.
    if ($text.IndexOf($indexName, [System.StringComparison]::Ordinal) -lt 0) {
        Write-Warning "  SKIP  $($f.Name)  - the result would not name it. Nothing written."
        $skipped++
        continue
    }
    if (-not $text.EndsWith($LF)) { $text = $text + $LF }   # write() ends with newLine()

    $target = Join-Path $outFull $pullFileName
    if (Test-Path -LiteralPath $target) {
        Write-Warning "  SKIP  $pullFileName  - already exists in the output directory. Nothing overwritten."
        $skipped++
        continue
    }

    if ($Apply) {
        if (-not (Test-Path -LiteralPath $outFull -PathType Container)) {
            New-Item -ItemType Directory -Path $outFull -Force | Out-Null
        }
        # WriteAllText, never Set-Content/Out-File: UTF-8 without a BOM, and the string written
        # exactly as built, so the LF-only line endings survive.
        [System.IO.File]::WriteAllText($target, $text, $UTF8_NO_BOM)
        Write-Host "  WROTE $pullFileName  -> names $indexName  ($($text.Length) chars)"
    }
    else {
        Write-Host "  would write $pullFileName  -> names $indexName  ($($text.Length) chars)"
    }
    $planned++
}

Write-Host ''
Write-Host ("$planned file(s) " + $(if ($Apply) { 'written' } else { 'planned' }) + ", $skipped skipped, out of $($candidates.Count) candidate(s).")
if (-not $Apply) { Write-Host 'Nothing was written. Re-run with -Apply once the list above is what you expect.' }
Write-Host ''
Write-Host 'Then, before delivering: move the old PULL files to a quarantine folder (do not delete them),'
Write-Host 'copy these beside their INDX, and run elarcheck with checkPull=true over that directory.'
Write-Host 'It reports PullMissing and PullUnreferenced, which is exactly the verification for this job.'
