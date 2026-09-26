# Generates the file-search QA corpus for MultiReplace vs. N++ comparison.
# Usage: powershell -ExecutionPolicy Bypass -File gen_search_corpus.ps1 [targetDir]
# Search "Fi" (Regex + Match case, filter *.*) and compare against EXPECTED below.
param([string]$Dir = "$PSScriptRoot\search_corpus")

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dir | Out-Null

# 4 occurrences of "Fi" per text file: First, Fight, Fi, Fi
$text = "First Fight`r`nFi second`r`nlast Fi`r`n"

function Bytes([string]$s, [System.Text.Encoding]$enc, [byte[]]$bom) {
    $b = $enc.GetBytes($s)
    if ($bom) { return $bom + $b } else { return $b }
}

$u8    = [System.Text.Encoding]::UTF8
$u16le = [System.Text.Encoding]::Unicode
$u16be = [System.Text.Encoding]::BigEndianUnicode
$ansi  = [System.Text.Encoding]::GetEncoding(1252)

[System.IO.File]::WriteAllBytes("$Dir\ansi.txt",         (Bytes ($text + "Gruesse: $([char]0xE4)$([char]0xF6)$([char]0xFC)`r`n") $ansi $null))
[System.IO.File]::WriteAllBytes("$Dir\utf8.txt",         (Bytes $text $u8 $null))
[System.IO.File]::WriteAllBytes("$Dir\utf8-bom.txt",     (Bytes $text $u8 @(0xEF,0xBB,0xBF)))
[System.IO.File]::WriteAllBytes("$Dir\utf16le-bom.txt",  (Bytes $text $u16le @(0xFF,0xFE)))
[System.IO.File]::WriteAllBytes("$Dir\utf16be-bom.txt",  (Bytes $text $u16be @(0xFE,0xFF)))
[System.IO.File]::WriteAllBytes("$Dir\utf16le-nobom.txt",(Bytes $text $u16le $null))
[System.IO.File]::WriteAllBytes("$Dir\utf16be-nobom.txt",(Bytes $text $u16be $null))
[System.IO.File]::WriteAllBytes("$Dir\utf16-odd.txt",    ((Bytes $text $u16le @(0xFF,0xFE)) + [byte]0x58))
[System.IO.File]::WriteAllBytes("$Dir\noext",            (Bytes $text $u8 $null))

# Binary: early NULs, contains one "Fi" byte pair after the NUL block
$bin = [byte[]](0x4D,0x5A,0x00,0x00) + (1..64 | ForEach-Object { [byte]0x07 }) + $u8.GetBytes("Fi")
[System.IO.File]::WriteAllBytes("$Dir\binary.bin", $bin)

# Hidden file + file inside hidden folder
[System.IO.File]::WriteAllBytes("$Dir\hidden.txt", (Bytes $text $u8 $null))
(Get-Item "$Dir\hidden.txt").Attributes = 'Hidden'
New-Item -ItemType Directory -Force -Path "$Dir\hiddendir" | Out-Null
[System.IO.File]::WriteAllBytes("$Dir\hiddendir\inside.txt", (Bytes $text $u8 $null))
(Get-Item "$Dir\hiddendir").Attributes = 'Directory,Hidden'

Write-Host ""
Write-Host "Corpus written to: $Dir"
Write-Host ""
Write-Host "EXPECTED for search 'Fi' (Regex + Match case, filter *.*, subfolders ON, hidden folders OFF)."
Write-Host "Each text file contains 4 hits. 'noext' counts only if PathMatchSpecW matches it for *.*;"
Write-Host "that behavior is identical in N++ (same API) - note the outcome, it settles finding K2."
Write-Host ""
Write-Host "  Skip binary ON : 36 hits in 9 files [9 file(s) searched, 2 skipped: 2 binary]"
Write-Host "                   files: ansi, utf8, utf8-bom, utf16le-bom, utf16be-bom, utf16le-nobom, utf16-odd, noext, hidden.txt"
Write-Host "                   skipped: binary.bin + utf16be-nobom (binary; BE-noBOM undetected = N++ parity)"
Write-Host "                   utf16-odd: the dangling last byte is ignored, as in N++"
Write-Host "                   (without noext: 32 hits in 8 files, 8 file(s) searched)"
Write-Host "  Skip binary OFF: +binary.bin as raw bytes (1 hit) and +utf16be-nobom raw (0 hits)"
Write-Host "                   -> 37 hits in 10 files [11 file(s) searched]"
Write-Host "  Hidden folders ON: additionally hiddendir\inside.txt (+4 hits, +1 file)."
Write-Host "  Cross-check vs. N++ native Find in Files (same folder, *.*): with skip OFF the FILE SET"
Write-Host "  must be identical; hit deltas are acceptable only inside binary.bin and utf16be-nobom."
Write-Host "  Replace in Files on this folder: utf16-odd is skipped as not decodable (its dangling"
Write-Host "  byte cannot be written back) and stays untouched."
Write-Host "  Zero-length check: regex search '^' must report one hit per line, same count as N++."
