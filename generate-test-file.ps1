<#
.SYNOPSIS
    Generate a random file of a specified target size.

.DESCRIPTION
    PowerShell equivalent of generate-test-file.sh.
    Accepts human-readable sizes like "500 MB", "2GB", "10 KB", "1.5 GB".

.PARAMETER Size
    Target size with unit. Supported units: B, KB, MB, GB, TB.

.PARAMETER OutputFile
    Optional output path. Default: testfile-<normalized>.bin

.EXAMPLE
    .\generate-test-file.ps1 "10 GB"

.EXAMPLE
    .\generate-test-file.ps1 500MB my-large-file.bin

.EXAMPLE
    .\generate-test-file.ps1 "1.5 GB" uploads\payload.bin
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Size,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-Usage {
    $scriptName = Split-Path -Leaf $PSCommandPath
    if (-not $scriptName) {
        $scriptName = 'generate-test-file.ps1'
    }

    Write-Host "Usage: .\$scriptName <size> [output_file]"
    Write-Host '  <size>  Target size with unit (B, KB, MB, GB, TB). Examples: "10 GB", 500MB, "1.5 GB"'
    Write-Host '  [output_file]  Optional output path (default: testfile-<normalized>.bin)'
}

function Parse-Size {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputSize
    )

    $normalized = ($InputSize -replace '\s+', '').ToUpperInvariant()

    if ($normalized -notmatch '^([0-9]+\.?[0-9]*)([KMGT]?B)$') {
        throw "Invalid size format '$InputSize'. Use a number followed by B, KB, MB, GB, or TB."
    }

    $number = [double]$matches[1]
    $unit = $matches[2]

    $multiplier = switch ($unit) {
        'B'  { 1L }
        'KB' { 1024L }
        'MB' { 1024L * 1024L }
        'GB' { 1024L * 1024L * 1024L }
        'TB' { 1024L * 1024L * 1024L * 1024L }
        default { throw "Unsupported unit '$unit'." }
    }

    $total = [math]::Round($number * $multiplier, 0)
    if ($total -lt 0) {
        throw "Size must be zero or greater. Received: $InputSize"
    }

    return [long]$total
}

try {
    $totalBytes = Parse-Size -InputSize $Size
} catch {
    Write-Error $_.Exception.Message
    Show-Usage
    exit 1
}

if (-not $OutputFile) {
    $sanitized = ($Size -replace '\s+', '').ToLowerInvariant()
    $OutputFile = "testfile-$sanitized.bin"
}

$outputParent = Split-Path -Parent $OutputFile
if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
}

$blockSize = 1MB
$fullBlocks = [long]($totalBytes / $blockSize)
$remainder = [int]($totalBytes % $blockSize)

Write-Host "Generating $OutputFile ($totalBytes bytes) ..."

$fileStream = $null
$rng = $null

try {
    $fileStream = [System.IO.File]::Open($OutputFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    if ($fullBlocks -gt 0) {
        $buffer = New-Object byte[] $blockSize
        for ($i = 0L; $i -lt $fullBlocks; $i++) {
            $rng.GetBytes($buffer)
            $fileStream.Write($buffer, 0, $buffer.Length)

            if (($i % 64L) -eq 0L -or $i -eq ($fullBlocks - 1L)) {
                $written = (($i + 1L) * $blockSize)
                if ($remainder -gt 0) {
                    $written += 0
                }
                $percent = if ($totalBytes -gt 0) { [int](($written * 100L) / $totalBytes) } else { 100 }
                if ($percent -gt 100) { $percent = 100 }
                Write-Progress -Activity 'Generating random file' -Status "$written / $totalBytes bytes" -PercentComplete $percent
            }
        }
    }

    if ($remainder -gt 0) {
        $tail = New-Object byte[] $remainder
        $rng.GetBytes($tail)
        $fileStream.Write($tail, 0, $tail.Length)
    }

    Write-Progress -Activity 'Generating random file' -Completed
}
finally {
    if ($rng) {
        $rng.Dispose()
    }
    if ($fileStream) {
        $fileStream.Dispose()
    }
}

$actual = (Get-Item -LiteralPath $OutputFile).Length
Write-Host "Done. $OutputFile is $actual bytes."

if ($actual -ne $totalBytes) {
    Write-Error "Warning: expected $totalBytes bytes but got $actual."
    exit 1
}

exit 0