#requires -version 5.1
<#
.SYNOPSIS
    Reads the whole elarxml step-log history of one or more workflows and writes a single CSV
    pairing each delivered INDX file with the original content file it carries.

.DESCRIPTION
    ELAR validates an INDX in full and rejects it in full, so the first question after a rejection is
    always "which documents were in that file". Since 2026-08-24 the elarxml executor answers it in
    its own step log, one line per document:

        O <TAB> 2026-08-26 12:00:00.123 <TAB> elarxml: <INDX> <- id=<ID> file=<NAME>

    This tool is the reader for those lines across every run of every named workflow. It is
    READ-ONLY: it writes nothing but its own output files, and it opens every log with
    FileShare.ReadWrite so it cannot block a feed that is running while it reads.

    THE RULE THAT DECIDES WHAT COMES OUT. The trace line above is written BEFORE the document is
    written, while the INDX only gets its final deliverable name when the batch closes. A batch that
    is aborted - an exception, the disk guard, a rollover - therefore leaves trace lines claiming
    documents inside a file that was never delivered. Harvesting the trace lines on their own would
    produce a CSV asserting deliveries that did not happen. So a pair is emitted only when the same
    log also shows that INDX reaching its final name, through either of the two lines that say so:

        elarxml: <INDX> delivered with <N> document(s), paired with <PULL>      (per batch, at close)
        elarxml: wrote <INDX> (<N> document(s))                                 (per run, at the end)

    The first is written only when logDocuments is on; the second only when the run reached its end.
    Either is accepted, and the traced count is reconciled against N. Traces without a delivery are
    counted and reported, never emitted and never silently dropped.

    COVERAGE. Runs older than 2026-08-24, and runs whose step set logDocuments=false, carry the
    delivery lines but no per-document trace. For those the mapping is not in the logs at all. The
    summary states how many such INDX files were seen, rather than letting them look like INDX files
    that carried nothing.

.PARAMETER Workflow
    One or more feed ids or workflow names. Wildcards are accepted and matched against both.

.PARAMETER BaseDir
    The feed base directory (orchestrator.default-base-dir): feeds live in <BaseDir>\<feedId>.

.PARAMETER WorkflowsDir
    The workflow XML directory (orchestrator.workflows-dir). Optional; it is what lets a workflow be
    named by its <workflow name="..."> instead of its feed id, and it is where a per-feed baseDir
    override is read from.

.PARAMETER PropertiesFile
    An application.properties to take BaseDir and WorkflowsDir from when they are not given.

.PARAMETER Out
    The CSV to write. Default .\elar-document-index.csv

.PARAMETER UndeliveredOut
    Optional CSV receiving the traces whose INDX never reached a final name, with the run they came
    from. Off by default: it is written only when asked for.

.PARAMETER Detailed
    Add feed_id, run_id, document_id and step columns to the main CSV. Default is the two columns
    that were asked for and nothing else.

.PARAMETER ExcludeTestRuns
    Skip runs whose id contains _test_. Off by default: a test run that wrote a real INDX delivered
    it, and a missing row is worse than an extra one for this question.

.PARAMETER From
    Only runs on or after this date (yyyyMMdd), taken from the run id.

.PARAMETER To
    Only runs on or before this date (yyyyMMdd), taken from the run id.

.PARAMETER NoHeader
    Omit the header line.

.PARAMETER NoBom
    Write UTF-8 without a byte order mark. The default includes one, because Excel needs it to read
    accented file names out of a semicolon-separated file.

.PARAMETER MaxTrackedFiles
    How many distinct content file names to remember for the duplicate-delivery check. Beyond this
    the check switches off and says so, rather than growing without bound. 0 disables it.

.EXAMPLE
    .\Get-ElarDocumentIndex.ps1 -Workflow CLIAC@DT -BaseDir D:\feeds -Out D:\temp\cliac.csv

.EXAMPLE
    .\Get-ElarDocumentIndex.ps1 -Workflow * -PropertiesFile C:\tomcat\config\application.properties
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Workflow,
    [string]$BaseDir,
    [string]$WorkflowsDir,
    [string]$PropertiesFile,
    [string]$Out = '.\elar-document-index.csv',
    [string]$UndeliveredOut,
    [switch]$Detailed,
    [switch]$ExcludeTestRuns,
    [string]$From,
    [string]$To,
    [switch]$NoHeader,
    [switch]$NoBom,
    [int]$MaxTrackedFiles = 1000000
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Built from character codes rather than from an escape sequence: the same discipline the JavaScript
# in this product follows, and it costs nothing here.
$CRLF = [string][char]13 + [string][char]10

# The markers, exactly as ElarRun writes them.
$MARK_ELARXML   = 'elarxml: '
$MARK_TRACE     = ' <- id='
$MARK_FILE      = ' file='
$MARK_DELIVERED = ' delivered with '
$MARK_PAIRED    = ' document(s), paired with '
$MARK_WROTE     = 'wrote '
$MARK_DOCS_TAIL = ' document(s))'

# ------------------------------------------------------------------ helpers

function Read-PropertyFile {
    param([string]$Path)
    $map = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $map }
    foreach ($line in [System.IO.File]::ReadAllLines($Path)) {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t.StartsWith('#') -or $t.StartsWith('!')) { continue }
        $eq = $t.IndexOf('=')
        if ($eq -lt 1) { continue }
        $map[$t.Substring(0, $eq).Trim()] = $t.Substring($eq + 1).Trim()
    }
    return $map
}

# Opened with FileShare.ReadWrite on purpose. This tool is meant to be pointed at a live feed
# directory, and a reader that locks a log the engine is still appending to would turn an analysis
# into an outage.
function Open-SharedReader {
    param([string]$Path)
    $fs = New-Object System.IO.FileStream(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite)
    return New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
}

function ConvertTo-CsvField {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value.IndexOf(';') -ge 0 -or $Value.IndexOf('"') -ge 0 `
        -or $Value.IndexOf([char]10) -ge 0 -or $Value.IndexOf([char]13) -ge 0) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

# .NET's current directory is NOT PowerShell's: a relative path handed to a StreamWriter lands
# wherever the process happened to start, which on a server is somewhere nobody will look for it.
function Resolve-OutPath {
    param([string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) { return [System.IO.Path]::GetFullPath($Path) }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $Path))
}

function New-CsvWriter {
    param([string]$Path, [bool]$Bom)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir }
    $enc = New-Object System.Text.UTF8Encoding($Bom)
    $w = New-Object System.IO.StreamWriter($Path, $false, $enc)
    $w.NewLine = $CRLF
    return $w
}

# The date a run id carries: <feedId>_yyyyMMdd-HHmmss_NNN, or <feedId>_test_yyyyMMdd-HHmmss_NNN.
function Get-RunDate {
    param([string]$RunId)
    $m = [regex]::Match($RunId, '_(\d{8})-\d{6}_')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

# ------------------------------------------------------------- resolve dirs

if ($PropertiesFile) {
    $props = Read-PropertyFile -Path $PropertiesFile
    if (-not $BaseDir -and $props.ContainsKey('orchestrator.default-base-dir')) {
        $BaseDir = $props['orchestrator.default-base-dir']
    }
    if (-not $WorkflowsDir -and $props.ContainsKey('orchestrator.workflows-dir')) {
        $WorkflowsDir = $props['orchestrator.workflows-dir']
    }
}

if (-not $BaseDir) {
    Write-Error "BaseDir is required (give -BaseDir, or -PropertiesFile carrying orchestrator.default-base-dir)"
    exit 2
}
if (-not (Test-Path -LiteralPath $BaseDir)) {
    Write-Error "BaseDir does not exist: $BaseDir"
    exit 2
}

# ------------------------------------------------------------- resolve feeds

# feedId -> @{ FeedDir; Name; HasElarStep }
$feeds = New-Object 'System.Collections.Specialized.OrderedDictionary'

function Add-Feed {
    param([string]$FeedId, [string]$FeedDir, [string]$Name, $HasElarStep)
    if ($feeds.Contains($FeedId)) { return }
    $feeds.Add($FeedId, @{ FeedDir = $FeedDir; Name = $Name; HasElarStep = $HasElarStep })
}

$definitions = @()
if ($WorkflowsDir -and (Test-Path -LiteralPath $WorkflowsDir)) {
    foreach ($f in (Get-ChildItem -LiteralPath $WorkflowsDir -Filter '*.xml' -File)) {
        try {
            $doc = New-Object System.Xml.XmlDocument
            $doc.Load($f.FullName)
        } catch {
            Write-Warning ("could not parse " + $f.Name + ": " + $_.Exception.Message)
            continue
        }
        $root = $doc.DocumentElement
        if ($null -eq $root -or $root.LocalName -ne 'workflow') { continue }
        $hasElar = $false
        foreach ($s in $doc.SelectNodes('//step')) {
            if ($s.GetAttribute('exec') -eq 'elarxml') { $hasElar = $true }
        }
        $definitions += ,@{
            FeedId   = $root.GetAttribute('feedId')
            Name     = $root.GetAttribute('name')
            BaseDir  = $root.GetAttribute('baseDir')
            HasElar  = $hasElar
            File     = $f.Name
        }
    }
} elseif ($WorkflowsDir) {
    Write-Warning "WorkflowsDir does not exist, so workflows can only be named by feed id: $WorkflowsDir"
}

foreach ($token in $Workflow) {
    $matched = $false
    foreach ($d in $definitions) {
        if (($d.FeedId -like $token) -or ($d.Name -and ($d.Name -like $token))) {
            $dir = Join-Path $BaseDir $d.FeedId
            if ($d.BaseDir) {
                if ([System.IO.Path]::IsPathRooted($d.BaseDir)) {
                    $dir = Join-Path $d.BaseDir $d.FeedId
                } else {
                    # The product resolves a relative baseDir against the JVM's working directory,
                    # which is Tomcat's and is not knowable from here. Saying so beats guessing.
                    Write-Warning ("workflow " + $d.FeedId + " declares a relative baseDir '" + $d.BaseDir `
                        + "'; using -BaseDir instead, which may be the wrong tree")
                }
            }
            Add-Feed -FeedId $d.FeedId -FeedDir $dir -Name $d.Name -HasElarStep $d.HasElar
            $matched = $true
        }
    }
    # A feed whose definition has been deleted still has its history on disk, so the directory is
    # also matched directly. That is the case this tool is most likely to be run for.
    foreach ($dir in (Get-ChildItem -LiteralPath $BaseDir -Directory | Where-Object { $_.Name -like $token })) {
        Add-Feed -FeedId $dir.Name -FeedDir $dir.FullName -Name $null -HasElarStep $null
        $matched = $true
    }
    if (-not $matched) { Write-Warning "no workflow or feed directory matched: $token" }
}

if ($feeds.Count -eq 0) {
    Write-Error "nothing to scan: no workflow or feed directory matched"
    exit 2
}

foreach ($id in $feeds.Keys) {
    $f = $feeds[$id]
    if ($f.HasElarStep -eq $false) {
        Write-Warning ("$id has no elarxml step in its CURRENT definition; its history is scanned anyway")
    }
}

# ------------------------------------------------------------------- scan

$Out = Resolve-OutPath $Out
if ($UndeliveredOut) { $UndeliveredOut = Resolve-OutPath $UndeliveredOut }
$writer = New-CsvWriter -Path $Out -Bom (-not $NoBom)
$undelWriter = $null
if ($UndeliveredOut) { $undelWriter = New-CsvWriter -Path $UndeliveredOut -Bom (-not $NoBom) }

try {
    if (-not $NoHeader) {
        if ($Detailed) { $writer.WriteLine('indx_file;original_file;document_id;feed_id;run_id;step') }
        else { $writer.WriteLine('indx_file;original_file') }
        if ($undelWriter) { $undelWriter.WriteLine('indx_file;original_file;document_id;feed_id;run_id;step') }
    }

    $stat = @{
        Feeds = 0; Runs = 0; RunsSkipped = 0; Logs = 0; ElarLogs = 0; Lines = 0
        Rows = 0; RowsFromTestRuns = 0
        DeliveredIndx = 0; IndxWithoutTrace = 0
        OrphanTraces = 0
        Mismatches = 0; Duplicates = 0
    }
    $mismatchNotes = New-Object System.Collections.ArrayList
    $noTraceNotes  = New-Object System.Collections.ArrayList
    $orphanNotes   = New-Object System.Collections.ArrayList
    $dupNotes      = New-Object System.Collections.ArrayList
    $seenFiles = @{}
    $trackingOn = ($MaxTrackedFiles -gt 0)
    $started = Get-Date

    foreach ($feedId in $feeds.Keys) {
        $feed = $feeds[$feedId]
        $runsRoot = Join-Path (Join-Path $feed.FeedDir '_logs') 'runs'
        if (-not (Test-Path -LiteralPath $runsRoot)) {
            Write-Warning ("$feedId has no run history at " + $runsRoot)
            continue
        }
        $stat.Feeds++

        foreach ($runDir in (Get-ChildItem -LiteralPath $runsRoot -Directory | Sort-Object Name)) {
            $runId = $runDir.Name
            $isTest = $runId.Contains('_test_')
            if ($ExcludeTestRuns -and $isTest) { $stat.RunsSkipped++; continue }
            $runDate = Get-RunDate -RunId $runId
            if ($From -and $runDate -and ($runDate -lt $From)) { $stat.RunsSkipped++; continue }
            if ($To -and $runDate -and ($runDate -gt $To))   { $stat.RunsSkipped++; continue }
            $stat.Runs++

            foreach ($logFile in (Get-ChildItem -LiteralPath $runDir.FullName -Filter '*.log' -File | Sort-Object Name)) {
                $stat.Logs++
                $step = [System.IO.Path]::GetFileNameWithoutExtension($logFile.Name)

                # PASS 1 - which INDX files reached a final name, and how many documents each says
                # it carries. Two passes rather than one because the delivery line comes AFTER the
                # traces it validates; buffering them instead would hold a whole batch in memory.
                $delivered = @{}      # indx -> declared document count
                $traceSeen = $false
                $isElarLog = $false

                $r = Open-SharedReader -Path $logFile.FullName
                try {
                    while ($null -ne ($line = $r.ReadLine())) {
                        $stat.Lines++
                        $i = $line.IndexOf($MARK_ELARXML)
                        if ($i -lt 0) { continue }
                        $isElarLog = $true
                        $rest = $line.Substring($i + $MARK_ELARXML.Length)

                        if ($rest.IndexOf($MARK_TRACE) -ge 0) { $traceSeen = $true; continue }

                        $d = $rest.IndexOf($MARK_DELIVERED)
                        if ($d -ge 0) {
                            $indx = $rest.Substring(0, $d)
                            $tail = $rest.Substring($d + $MARK_DELIVERED.Length)
                            $p = $tail.IndexOf($MARK_PAIRED)
                            if ($p -ge 0) {
                                $n = 0
                                if ([int]::TryParse($tail.Substring(0, $p), [ref]$n)) { $delivered[$indx] = $n }
                            }
                            continue
                        }

                        if ($rest.StartsWith($MARK_WROTE)) {
                            $t = $rest.Substring($MARK_WROTE.Length)
                            # a PULL is written on its own line with no count: not an INDX
                            if (-not $t.EndsWith($MARK_DOCS_TAIL)) { continue }
                            $p = $t.LastIndexOf(' (')
                            if ($p -lt 1) { continue }
                            $indx = $t.Substring(0, $p)
                            # the count sits between ' (' and the ' document(s))' tail, and nothing else
                            $numTxt = $t.Substring($p + 2, $t.Length - $MARK_DOCS_TAIL.Length - $p - 2)
                            $n = 0
                            if ([int]::TryParse($numTxt, [ref]$n)) {
                                if (-not $delivered.ContainsKey($indx)) { $delivered[$indx] = $n }
                            }
                        }
                    }
                } finally { $r.Dispose() }

                if (-not $isElarLog) { continue }
                $stat.ElarLogs++
                $stat.DeliveredIndx += $delivered.Count

                if (-not $traceSeen) {
                    # Delivered, but the mapping is not in this log: either the run predates the
                    # per-document trace (2026-08-24) or the step had logDocuments off.
                    if ($delivered.Count -gt 0) {
                        $stat.IndxWithoutTrace += $delivered.Count
                        $null = $noTraceNotes.Add("$feedId/$runId/$step : " + $delivered.Count + " delivered INDX, no per-document trace")
                    }
                    continue
                }

                # PASS 2 - emit the pairs whose INDX is known to have been delivered.
                $tracedPerIndx = @{}
                $r = Open-SharedReader -Path $logFile.FullName
                try {
                    while ($null -ne ($line = $r.ReadLine())) {
                        $i = $line.IndexOf($MARK_ELARXML)
                        if ($i -lt 0) { continue }
                        $rest = $line.Substring($i + $MARK_ELARXML.Length)
                        $j = $rest.IndexOf($MARK_TRACE)
                        if ($j -lt 0) { continue }

                        $indx = $rest.Substring(0, $j)
                        $tail = $rest.Substring($j + $MARK_TRACE.Length)
                        # LAST occurrence: an id may contain anything, the file name ends the line
                        $k = $tail.LastIndexOf($MARK_FILE)
                        if ($k -lt 0) { continue }
                        $docId = $tail.Substring(0, $k)
                        $file  = $tail.Substring($k + $MARK_FILE.Length)

                        if ($delivered.ContainsKey($indx)) {
                            if ($tracedPerIndx.ContainsKey($indx)) { $tracedPerIndx[$indx]++ }
                            else { $tracedPerIndx[$indx] = 1 }

                            if ($Detailed) {
                                $writer.WriteLine((ConvertTo-CsvField $indx) + ';' + (ConvertTo-CsvField $file) + ';' `
                                    + (ConvertTo-CsvField $docId) + ';' + (ConvertTo-CsvField $feedId) + ';' `
                                    + (ConvertTo-CsvField $runId) + ';' + (ConvertTo-CsvField $step))
                            } else {
                                $writer.WriteLine((ConvertTo-CsvField $indx) + ';' + (ConvertTo-CsvField $file))
                            }
                            $stat.Rows++
                            if ($isTest) { $stat.RowsFromTestRuns++ }

                            if ($trackingOn) {
                                if ($seenFiles.ContainsKey($file)) {
                                    if ($seenFiles[$file] -ne $indx) {
                                        $stat.Duplicates++
                                        if ($dupNotes.Count -lt 10) {
                                            $null = $dupNotes.Add("$file : " + $seenFiles[$file] + " and " + $indx)
                                        }
                                    }
                                } else {
                                    if ($seenFiles.Count -ge $MaxTrackedFiles) {
                                        $trackingOn = $false
                                        Write-Warning ("more than $MaxTrackedFiles distinct content files: the" `
                                            + " duplicate-delivery check is off from here on. Raise -MaxTrackedFiles to keep it.")
                                    } else {
                                        $seenFiles[$file] = $indx
                                    }
                                }
                            }
                        } else {
                            # traced into a batch that never reached a final name
                            $stat.OrphanTraces++
                            if ($orphanNotes.Count -lt 10) {
                                $null = $orphanNotes.Add("$feedId/$runId/$step : $indx <- $file")
                            }
                            if ($undelWriter) {
                                $undelWriter.WriteLine((ConvertTo-CsvField $indx) + ';' + (ConvertTo-CsvField $file) + ';' `
                                    + (ConvertTo-CsvField $docId) + ';' + (ConvertTo-CsvField $feedId) + ';' `
                                    + (ConvertTo-CsvField $runId) + ';' + (ConvertTo-CsvField $step))
                            }
                        }
                    }
                } finally { $r.Dispose() }

                foreach ($indx in $delivered.Keys) {
                    $traced = 0
                    if ($tracedPerIndx.ContainsKey($indx)) { $traced = $tracedPerIndx[$indx] }
                    if ($traced -ne $delivered[$indx]) {
                        $stat.Mismatches++
                        if ($mismatchNotes.Count -lt 10) {
                            $null = $mismatchNotes.Add("$feedId/$runId/$step : $indx says " + $delivered[$indx] `
                                + " document(s), " + $traced + " traced")
                        }
                    }
                }
            }
        }
    }

    $writer.Flush()
    if ($undelWriter) { $undelWriter.Flush() }

    # ------------------------------------------------------------- summary
    $elapsed = (Get-Date) - $started
    Write-Host ''
    Write-Host ("Get-ElarDocumentIndex - " + $stat.Rows + " pair(s) written to " + $Out)
    Write-Host ("  scanned      : " + $stat.Feeds + " feed(s), " + $stat.Runs + " run(s), " `
        + $stat.ElarLogs + " elarxml log(s) of " + $stat.Logs + ", " + $stat.Lines + " line(s) in " `
        + [math]::Round($elapsed.TotalSeconds, 1) + "s")
    if ($stat.RunsSkipped -gt 0) { Write-Host ("  skipped      : " + $stat.RunsSkipped + " run(s) by filter") }
    Write-Host ("  delivered    : " + $stat.DeliveredIndx + " INDX file(s)")
    if ($stat.RowsFromTestRuns -gt 0) {
        Write-Host ("  test runs    : " + $stat.RowsFromTestRuns + " row(s) come from _test_ runs (use -ExcludeTestRuns to drop them)")
    }

    if ($stat.IndxWithoutTrace -gt 0) {
        Write-Host ''
        Write-Host ("  " + $stat.IndxWithoutTrace + " delivered INDX carry NO per-document trace, so their contents are")
        Write-Host ("  not in the logs at all. Either the run predates 2026-08-24, or the step had")
        Write-Host ("  logDocuments off. They are absent from the CSV, not empty in it:")
        foreach ($n in ($noTraceNotes | Select-Object -First 10)) { Write-Host ("    " + $n) }
        if ($noTraceNotes.Count -gt 10) { Write-Host ("    ... and " + ($noTraceNotes.Count - 10) + " more") }
    }

    if ($stat.OrphanTraces -gt 0) {
        Write-Host ''
        Write-Host ("  " + $stat.OrphanTraces + " document(s) were traced into an INDX that never reached a final")
        Write-Host ("  name - an aborted batch. They are NOT in the CSV, because they were not delivered:")
        foreach ($n in $orphanNotes) { Write-Host ("    " + $n) }
        if ($stat.OrphanTraces -gt $orphanNotes.Count) { Write-Host ("    ... and " + ($stat.OrphanTraces - $orphanNotes.Count) + " more") }
        if (-not $UndeliveredOut) { Write-Host ("    (-UndeliveredOut <file> writes them out in full)") }
    }

    if ($stat.Mismatches -gt 0) {
        Write-Host ''
        Write-Host ("  " + $stat.Mismatches + " INDX file(s) traced a different number of documents than they say they")
        Write-Host ("  carry. Report this: the two lines are written by the same executor and must agree.")
        foreach ($n in $mismatchNotes) { Write-Host ("    " + $n) }
    }

    if ($stat.Duplicates -gt 0) {
        Write-Host ''
        Write-Host ("  " + $stat.Duplicates + " content file(s) were delivered in MORE THAN ONE INDX. Both rows are in")
        Write-Host ("  the CSV and both are true: this is the duplicate-delivery shape a re-run produces.")
        foreach ($n in $dupNotes) { Write-Host ("    " + $n) }
        if ($stat.Duplicates -gt $dupNotes.Count) { Write-Host ("    ... and " + ($stat.Duplicates - $dupNotes.Count) + " more") }
    }
    Write-Host ''
} finally {
    if ($writer) { $writer.Dispose() }
    if ($undelWriter) { $undelWriter.Dispose() }
}

exit 0
