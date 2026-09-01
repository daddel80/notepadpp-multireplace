# Generates the scale/UI QA corpus for MultiReplace.
# Covers: locale thousand separators in the dock/status summaries, "Collapse to
# file list", search in open buffers, the per-file CSV delimiter rescan and the
# CSV header change (find/replace searches header lines again).
#
# Usage: powershell -ExecutionPolicy Bypass -File gen_scale_corpus.ps1 [-Dir <path>] [-Clean]
#
# Files are deliberately tiny; the counts (not the bytes) push every summary
# number past 1000 so the locale separator becomes visible.

param(
    [string]$Dir             = "$PSScriptRoot\scale_corpus",
    [int]   $BulkFiles       = 1200,   # 1 hit each  -> hits AND files > 1000
    [int]   $BinFiles        = 1050,   # NUL files   -> "skipped" > 1000
    [int]   $HitsPerBigFile  = 1500,   # per-file hit count > 1000
    [int]   $LargeMB         = 2,      # for the "too large" skip (limit is OFF by default)
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

# ASCII marker, absent from real-world text. KOLX marks the CSV column tests.
$TOKEN    = 'ZTOKEN'
$CSVTOKEN = 'KOLX'

# Encoding-proof accented word (works whether or not this .ps1 has a BOM)
$E        = [char]0x00C9                      # E with acute
$ACCENTED = "PR$($E)C$($E)DENT"

# ---------------------------------------------------------------- helpers ---
function Reset-Dir([string]$path) {
    if (Test-Path $path) {
        # read-only files from a previous run would block Remove-Item
        Get-ChildItem $path -Recurse -File -Force -ErrorAction SilentlyContinue |
            ForEach-Object { $_.IsReadOnly = $false }
        Remove-Item $path -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $path | Out-Null
}

function Write-Text([string]$path, [string]$text, [System.Text.Encoding]$enc = $null) {
    if (-not $enc) { $enc = New-Object System.Text.UTF8Encoding($false) }  # UTF-8, no BOM
    [System.IO.File]::WriteAllText($path, $text, $enc)
}

$u8      = New-Object System.Text.UTF8Encoding($false)
$u8bom   = New-Object System.Text.UTF8Encoding($true)
$u16le   = New-Object System.Text.UnicodeEncoding($false, $true)   # LE + BOM
$u16be   = New-Object System.Text.UnicodeEncoding($true,  $true)   # BE + BOM
$ansi    = [System.Text.Encoding]::GetEncoding(1252)

if ($Clean) { Reset-Dir $Dir } else { New-Item -ItemType Directory -Force -Path $Dir | Out-Null }

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# ------------------------------------------------------- 01 bulk (1 hit) ----
# Drives both big numbers at once: one hit per file, one file per hit.
$d = Join-Path $Dir '01_bulk'; New-Item -ItemType Directory -Force -Path $d | Out-Null
for ($i = 1; $i -le $BulkFiles; $i++) {
    $n = '{0:D4}' -f $i
    Write-Text (Join-Path $d "bulk_$n.txt") "line one`r`nid $n $TOKEN here`r`nline three`r`n"
}
Write-Host ("01_bulk       : {0,6} files x 1 hit" -f $BulkFiles)

# ------------------------------------------- 02 many hits in a single file ---
# Makes the PER-FILE count in the dock cross 1000: "name.txt (1 500 hits)".
$d = Join-Path $Dir '02_hits'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$sb = New-Object System.Text.StringBuilder
for ($i = 1; $i -le $HitsPerBigFile; $i++) { [void]$sb.Append("row $i $TOKEN`r`n") }
$bigText = $sb.ToString()
$BigFiles = 3
for ($i = 1; $i -le $BigFiles; $i++) { Write-Text (Join-Path $d "many_$i.txt") $bigText }
Write-Host ("02_hits       : {0,6} files x {1} hits" -f $BigFiles, $HitsPerBigFile)

# ------------------------------------------------ 03 binary (NUL) files -----
# Skipped with "Skip binary files" ON -> pushes the skipped count past 1000.
# With the option OFF each one contributes exactly 1 hit.
$d = Join-Path $Dir '03_binary'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$binPrefix = [byte[]](0x4D,0x5A,0x00,0x00,0x00,0x01)
$binSuffix = $u8.GetBytes(" $TOKEN ")
for ($i = 1; $i -le $BinFiles; $i++) {
    $n = '{0:D4}' -f $i
    [System.IO.File]::WriteAllBytes((Join-Path $d "bin_$n.bin"), ($binPrefix + $binSuffix))
}
Write-Host ("03_binary     : {0,6} files (NUL bytes, 1 hit each)" -f $BinFiles)

# --------------------------------------------------------- 04 large files ---
# The size limit is OFF by default: these are searched normally until you turn
# "Limit file size" ON (set it to 1 MB) - then they appear as "too large".
$d = Join-Path $Dir '04_large'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$filler = ('x' * 78 + "`r`n") * 128                     # 10 KB per chunk
$LargeFiles = 2
for ($i = 1; $i -le $LargeFiles; $i++) {
    $p  = Join-Path $d "large_$i.txt"
    $fs = [System.IO.File]::CreateText($p)
    try {
        $fs.Write("head $TOKEN`r`n")                     # exactly one hit
        for ($k = 0; $k -lt ($LargeMB * 100); $k++) { $fs.Write($filler) }
    } finally { $fs.Close() }
}
Write-Host ("04_large      : {0,6} files x ~{1} MB (1 hit each)" -f $LargeFiles, $LargeMB)

# ------------------------------------------------------ 05 read-only files ---
# Found by Find in Files, refused by Replace in Files ("N file(s) skipped: read-only").
$d = Join-Path $Dir '05_readonly'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$RoFiles = 3
for ($i = 1; $i -le $RoFiles; $i++) {
    $p = Join-Path $d "ro_$i.txt"
    # a previous run left these read-only - clear the flag before rewriting
    if (Test-Path $p) { (Get-Item $p).IsReadOnly = $false }
    Write-Text $p "read only $TOKEN`r`n"
    (Get-Item $p).IsReadOnly = $true
}
Write-Host ("05_readonly   : {0,6} files (ReadOnly attribute)" -f $RoFiles)

# --------------------------------------------------------- 06 encodings -----
# One hit each. enc_nul_utf8.txt is the Mark_Style.txt analogue: structurally
# valid UTF-8 *with* NUL bytes -> classified binary, but readable with the
# option OFF (accented word must appear without mojibake).
$d = Join-Path $Dir '06_encodings'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$encText = "$ACCENTED line`r`nmarker $TOKEN`r`n"
[System.IO.File]::WriteAllText((Join-Path $d 'enc_utf8.txt'),     $encText, $u8)
[System.IO.File]::WriteAllText((Join-Path $d 'enc_utf8bom.txt'),  $encText, $u8bom)
[System.IO.File]::WriteAllText((Join-Path $d 'enc_utf16le.txt'),  $encText, $u16le)
[System.IO.File]::WriteAllText((Join-Path $d 'enc_utf16be.txt'),  $encText, $u16be)
[System.IO.File]::WriteAllText((Join-Path $d 'enc_ansi1252.txt'), $encText, $ansi)
[System.IO.File]::WriteAllBytes((Join-Path $d 'enc_nul_utf8.txt'),
    ($u8.GetBytes("$ACCENTED table`r`n") + [byte[]](0x00,0x00,0x00) + $u8.GetBytes("marker $TOKEN`r`n")))
$EncFiles       = 6
$EncBinaryFiles = 1        # enc_nul_utf8.txt
Write-Host ("06_encodings  : {0,6} files (UTF-8/BOM, UTF-16 LE+BE, ANSI, NUL+UTF-8)" -f $EncFiles)

# --------------------------------------------------------------- 07 CSV -----
# Three different column layouts. Column-mode search must rescan the delimiters
# for EVERY file - with the old bug, files 2..N were filtered using file 1's
# frozen column positions. csv_A's header carries a token on purpose: header
# lines are searched again (they are only protected from sort/duplicates).
$d = Join-Path $Dir '07_csv'; New-Item -ItemType Directory -Force -Path $d | Out-Null

# Row 1 of each file is the header line. csv_A carries a token there on purpose.
$csvDefs = [ordered]@{
    'csv_A.csv' = @(                     # 3 columns
        "id;$CSVTOKEN;note"                  # header -> col2
        "1;$CSVTOKEN;alpha"                  # col2
        "2;beta;$CSVTOKEN"                   # col3
        "3;$CSVTOKEN;gamma"                  # col2
        "4;delta;epsilon"
        "5;$CSVTOKEN;zeta"                   # col2
        "6;eta;$CSVTOKEN"                    # col3
    )
    'csv_B.csv' = @(                     # 7 columns
        "c1;c2;c3;c4;c5;c6;c7"
        "a1;a2;a3;a4;$CSVTOKEN;a6;a7"        # col5
        "b1;$CSVTOKEN;b3;b4;b5;b6;b7"        # col2
        "c1;c2;c3;c4;$CSVTOKEN;c6;c7"        # col5
        "d1;d2;d3;d4;d5;d6;$CSVTOKEN"        # col7
        "e1;e2;e3;e4;$CSVTOKEN;e6;e7"        # col5
    )
    'csv_C.csv' = @(                     # 2 columns
        "k;v"
        "1;$CSVTOKEN"                        # col2
        "2;$CSVTOKEN"                        # col2
        "$CSVTOKEN;3"                        # col1
        "4;$CSVTOKEN"                        # col2
    )
}

# Counts are derived from the rows above - the table below can never drift
# away from the files that were actually written.
function Count-InColumn([string[]]$rows, [int]$col, [string]$tok) {
    $n = 0
    foreach ($r in $rows) {
        $cells = $r.Split(';')
        if ($col -eq 0) { $n += ([regex]::Matches($r, [regex]::Escape($tok))).Count }
        elseif ($cells.Count -ge $col) {
            $n += ([regex]::Matches($cells[$col - 1], [regex]::Escape($tok))).Count
        }
    }
    return $n
}

foreach ($name in $csvDefs.Keys) {
    Write-Text (Join-Path $d $name) (($csvDefs[$name]) -join "`r`n")
}
$CsvFiles   = $csvDefs.Count
$CsvColumns = @(0, 1, 2, 5)                  # 0 = no column mode (whole line)
$csvCount   = @{}                            # "file|col" -> hits
foreach ($name in $csvDefs.Keys) {
    foreach ($c in $CsvColumns) {
        $csvCount["$name|$c"] = Count-InColumn $csvDefs[$name] $c $CSVTOKEN
    }
}
$csvHeaderCol2 = Count-InColumn @($csvDefs['csv_A.csv'][0]) 2 $CSVTOKEN   # token in csv_A's header
Write-Host ("07_csv        : {0,6} files (3 / 7 / 2 columns)" -f $CsvFiles)

# ------------------------------------------------------- 08 open buffers -----
# Open these in Notepad++, edit without saving, then search/replace.
$d = Join-Path $Dir '08_buffer'; New-Item -ItemType Directory -Force -Path $d | Out-Null
$BufFiles = 4
for ($i = 1; $i -le $BufFiles; $i++) {
    Write-Text (Join-Path $d "buf_$i.txt") "on disk $TOKEN`r`nedit this line in N++ without saving`r`n"
}
Write-Host ("08_buffer     : {0,6} files (for the unsaved-buffer test)" -f $BufFiles)

$sw.Stop()

# ------------------------------------------------------------ expectations ---
$txtFiles   = $BulkFiles + $BigFiles + $LargeFiles + $RoFiles + $EncFiles + $BufFiles
$txtHitFile = $BulkFiles + $BigFiles + $LargeFiles + $RoFiles + ($EncFiles - $EncBinaryFiles) + $BufFiles
$txtHits    = $BulkFiles + ($BigFiles * $HitsPerBigFile) + $LargeFiles + $RoFiles + ($EncFiles - $EncBinaryFiles) + $BufFiles
$allFiles   = $txtFiles + $BinFiles + $CsvFiles

$t1Hits = $txtHits;                 $t1Files = $txtHitFile;  $t1Searched = $txtFiles - $EncBinaryFiles
$t2Hits = $txtHits - $LargeFiles;   $t2Files = $txtHitFile - $LargeFiles
$t2Searched = $txtFiles - $EncBinaryFiles - $LargeFiles
$t3Searched = $allFiles - $BinFiles - $EncBinaryFiles
$t3Skipped  = $BinFiles + $EncBinaryFiles
$t4Hits = $txtHits + $BinFiles + $EncBinaryFiles
$t4Files = $txtHitFile + $BinFiles + $EncBinaryFiles

function N([int]$v) { '{0:N0}' -f $v }   # shown with the generating machine's locale

# CSV result table, built from the rows that were actually written
$csvNames = @($csvDefs.Keys)
$csvNames = @($csvDefs.Keys)
$csvTable  = "| Column mode    | " + (($csvNames | ForEach-Object { '{0,-5}' -f ($_ -replace '\.csv$','') }) -join ' | ') + " | total         |`r`n"
$csvTable += "|----------------|" + (($csvNames | ForEach-Object { '-------' }) -join '|') + "|---------------|`r`n"
foreach ($c in $CsvColumns) {
    $label = if ($c -eq 0) { 'off (all text)' } else { "column $c" }
    $cells = @(); $sum = 0; $nfile = 0
    foreach ($n in $csvNames) {
        $v = $csvCount["$n|$c"]
        $cells += ('{0,5}' -f $v)
        $sum += $v
        if ($v -gt 0) { $nfile++ }
    }
    $totalTxt = '{0} in {1} file{2}' -f $sum, $nfile, $(if ($nfile -eq 1) { '' } else { 's' })
    $csvTable += ("| {0,-14} | " -f $label) + ($cells -join ' | ') + (" | {0,-13} |`r`n" -f $totalTxt)
}
$csvTable = $csvTable.TrimEnd()

$expected = @"
# MultiReplace - scale corpus, expected results

Corpus: ``$Dir``
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') - $allFiles files total

Search term: ``$TOKEN`` (Match case ON, subfolders ON). CSV tests use ``$CSVTOKEN``.
Numbers below are shown with this machine's locale; the plugin must format the
summary the same way (FR: ``24 581``, DE: ``24.581``, EN: ``24,581``).

## T1 - separators, default settings
Filter ``*.txt``, Skip binary ON, Limit file size OFF.

    ($(N $t1Hits) hits in $(N $t1Files) file(s)) [$(N $t1Searched) file(s) searched, 1 skipped: 1 binary]

Check: hits, files and searched all carry a separator. Each
``02_hits\many_*.txt`` must read ``($(N $HitsPerBigFile) hits)`` - the per-file
count uses the same format as the header.

## T2 - "too large" skip
Same as T1, plus Settings > File Search: Limit file size ON, 1 MB.

    ($(N $t2Hits) hits in $(N $t2Files) file(s)) [$(N $t2Searched) file(s) searched, 3 skipped: 1 binary, 2 too large]

## T3 - separator inside the skip breakdown
Filter ``*.*``, Skip binary ON, Limit file size OFF.

    ($(N $t1Hits) hits in $(N $t1Files) file(s)) [$(N $t3Searched) file(s) searched, $(N $t3Skipped) skipped: $(N $t3Skipped) binary]

## T4 - Skip binary OFF
Filter ``*.*``, Skip binary OFF, Limit file size OFF.

    ($(N $t4Hits) hits in $(N $t4Files) file(s)) [$(N $allFiles) file(s) searched]

The $(N $BinFiles) ``.bin`` files and ``enc_nul_utf8.txt`` now contribute 1 hit each.
Search ``$ACCENTED`` instead: the accented word must be found in every
``06_encodings`` file with no mojibake (that is the Mark_Style.txt case).

## T5 - CSV: per-file delimiter rescan and header lines
Filter ``*.csv``, CSV scope, delimiter ``;``, quote none. Search ``$CSVTOKEN``.

$csvTable

Column 5 is the regression test: only csv_B has a fifth column. If files 2..N
were scanned with file 1's delimiter positions again, csv_A/csv_C would report
hits here (or csv_B would lose them).
csv_A in column 2 must report **$($csvCount['csv_A.csv|2'])** hits, $csvHeaderCol2 of them in the header line: find/replace searches header lines again, the header count only protects sort and duplicate detection now. A count of $($csvCount['csv_A.csv|2'] - $csvHeaderCol2) there means header exclusion came back.

## T6 - Collapse to file list
Run T1, then right-click in the result panel > **Collapse to file list**.
Expect: $(N $t1Files) file lines with their hit counts, no hit lines, the search
header still open. Then "Unfold all" restores everything.

## T7 - Collapse with several search blocks (regression)
Result panel context menu: make sure **Purge for next search** is OFF.
1. Run T1 - one block.
2. Run T1 again with filter ``bulk_1*.txt`` - the older block is collapsed automatically.
3. **Collapse to file list**.
Expect: the newest block becomes a file list, **the older block stays collapsed**.
It must not pop open - the entry only folds file/criteria lines, it never
changes which search blocks are open.

## T8 - open buffers (Settings > File Search: "Search open documents ...", ON by default)
1. Open all four ``08_buffer\buf_*.txt`` in Notepad++.
2. In ``buf_1.txt`` add two more ``$TOKEN`` and do **not** save.
3. Find in Files, filter ``buf_*.txt``: expect **6 hits in 4 files** (2 + 1 + 1 + 1).
   Turn the option off and repeat: **4 hits in 4 files** (on-disk state).
4. Replace in Files ``$TOKEN`` -> ``DONE`` over the same filter: the unsaved file
   is not written and is reported as ``1 file(s) skipped: open with unsaved changes``.
5. ``05_readonly``: Replace in Files reports ``$RoFiles file(s) skipped: read-only``.
"@

$expectedPath = Join-Path (Split-Path $Dir -Parent) ((Split-Path $Dir -Leaf) + '_EXPECTED.md')
Write-Text $expectedPath $expected      # kept OUTSIDE the corpus so it is never searched

Write-Host ""
Write-Host ("Corpus written to : {0}" -f $Dir)
Write-Host ("Expectations      : {0}" -f $expectedPath)
Write-Host ("Total files       : {0}   ({1:N1}s)" -f $allFiles, $sw.Elapsed.TotalSeconds)
Write-Host ""
Write-Host "Quick check - T1 (filter *.txt, Skip binary ON, size limit OFF):"
Write-Host ("  ({0} hits in {1} file(s)) [{2} file(s) searched, 1 skipped: 1 binary]" -f (N $t1Hits), (N $t1Files), (N $t1Searched))
Write-Host ""
Write-Host "Re-run with -Clean to rebuild from scratch (clears the read-only flags first)."
