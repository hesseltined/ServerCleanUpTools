#Requires -Version 5.1
# Purpose: Recursively find files older than N years (last write time) and plan or move them to an archive root with mirrored paths.
# Author: Doug Hesseltine
# Created: 2026-03-27
# Modified: 2026-03-27
# Version: 1.0.0
<#
.SYNOPSIS
    Finds files older than N years (by last write time) under an input folder and plans or moves them to an archive root with mirrored paths.

.DESCRIPTION
    By default runs in preview mode: no files are moved and no folders are created on the archive.
    Use -Commit to perform moves. With -Commit, the archive folder is created if it does not exist.
    The archive path must not be inside the input tree. Preview mode requires the archive folder to already exist.

.PARAMETER InputPath
    Root folder to scan recursively.

.PARAMETER ArchivePath
    Root folder where mirrored relative paths are created. Must not be equal to or under InputPath.

.PARAMETER Years
    Files with LastWriteTime older than (now minus this many years) are selected.

.PARAMETER Commit
    Perform Move-Item and create destination directories. Without this switch, only a plan is reported.

.PARAMETER Output
    Text or CSV. If omitted, you are prompted at the end. CSV is written under ArchivePath.

.PARAMETER RemoveEmptyFolders
    Yes or No. If omitted and -Commit was used, you are prompted whether to remove empty directories under InputPath (input root is never removed).

.NOTES
    Purpose: Age-based archival of old files to a separate archive location.
    Author: Doug Hesseltine
    Created: 2026-03-27
    Modified: 2026-03-27
    Version: 1.0.0

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7

    Preview only: lists planned moves, no changes.

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7 -Commit -Output CSV
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$ArchivePath,

    [Parameter(Mandatory = $true, Position = 2)]
    [double]$Years,

    [Parameter(Mandatory = $false)]
    [switch]$Commit,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Text', 'CSV')]
    [string]$Output,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Yes', 'No')]
    [string]$RemoveEmptyFolders
)

$ErrorActionPreference = 'Stop'

# Script metadata (bump Version when making meaningful changes)
$script:Version = '1.0.0'

function Get-ResolvedPath {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Path does not exist or is not reachable: $Path"
    }
    return (Get-Item -LiteralPath $Path).FullName.TrimEnd('\')
}

function Test-ArchiveUnderInput {
    param(
        [string]$InputResolved,
        [string]$ArchiveResolved
    )
    $in = $InputResolved.TrimEnd('\')
    $ar = $ArchiveResolved.TrimEnd('\')
    if ([string]::Compare($ar, $in, $true) -eq 0) {
        return $true
    }
    $prefix = $in + '\'
    return $ar.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-UniqueDestinationFilePath {
    param([string]$InitialDestPath)
    if (-not (Test-Path -LiteralPath $InitialDestPath)) {
        return $InitialDestPath
    }
    $directory = Split-Path -Parent $InitialDestPath
    $fileName = [System.IO.Path]::GetFileName($InitialDestPath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [System.IO.Path]::GetExtension($fileName)
    $n = 2
    while ($true) {
        $candidateName = '{0}_{1}{2}' -f $base, $n, $extension
        $candidate = Join-Path $directory $candidateName
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
        $n++
    }
}

function Get-RelativePathFromRoot {
    param(
        [string]$FileFullName,
        [string]$RootFullName
    )
    $root = $RootFullName.TrimEnd('\')
    if (-not $FileFullName.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "File is not under input root: $FileFullName"
    }
    $tail = $FileFullName.Substring($root.Length).TrimStart('\')
    return $tail
}

# --- Main ---

try {
    $inputResolved = Get-ResolvedPath -Path $InputPath
}
catch {
    Write-Error $_
    exit 1
}

try {
    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        if ($Commit) {
            $null = New-Item -ItemType Directory -Path $ArchivePath -Force -ErrorAction Stop
        }
        else {
            throw "ArchivePath does not exist. Create the folder first for preview mode, or use -Commit to create it."
        }
    }
    $archiveResolved = (Get-Item -LiteralPath $ArchivePath).FullName.TrimEnd('\')
}
catch {
    Write-Error $_
    exit 1
}

if (Test-ArchiveUnderInput -InputResolved $inputResolved -ArchiveResolved $archiveResolved) {
    Write-Error "ArchivePath must not be the same as or inside InputPath. Input: $inputResolved  Archive: $archiveResolved"
    exit 1
}

$cutoff = (Get-Date).AddYears(-$Years)
Write-Host "Cutoff (files older than this LastWriteTime): $cutoff" -ForegroundColor Cyan
Write-Host "Mode: $(if ($Commit) { 'COMMIT (files will be moved)' } else { 'PREVIEW (no changes)' })" -ForegroundColor $(if ($Commit) { 'Yellow' } else { 'Green' })

$results = New-Object System.Collections.Generic.List[object]

$files = @(
    Get-ChildItem -LiteralPath $inputResolved -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable gciErrors
)
foreach ($err in $gciErrors) {
    Write-Warning "Enumeration issue: $($err.Exception.Message)"
}

foreach ($file in $files) {
    $sourcePath = $file.FullName
    $destPath = $null
    try {
        if ($file.LastWriteTime -ge $cutoff) {
            continue
        }

        $relative = Get-RelativePathFromRoot -FileFullName $sourcePath -RootFullName $inputResolved
        $initialDest = Join-Path $archiveResolved $relative
        $destPath = Get-UniqueDestinationFilePath -InitialDestPath $initialDest

        if (-not $Commit) {
            $results.Add([pscustomobject]@{
                    SourcePath      = $sourcePath
                    DestinationPath = $destPath
                    LastWriteTime   = $file.LastWriteTime
                    Length          = $file.Length
                    Status          = 'Planned'
                    Message         = ''
                })
            continue
        }

        $destDir = Split-Path -Parent $destPath
        if (-not (Test-Path -LiteralPath $destDir)) {
            $null = New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop
        }

        Move-Item -LiteralPath $sourcePath -Destination $destPath -Force -ErrorAction Stop
        $results.Add([pscustomobject]@{
                SourcePath      = $sourcePath
                DestinationPath = $destPath
                LastWriteTime   = $file.LastWriteTime
                Length          = $file.Length
                Status          = 'Moved'
                Message         = ''
            })
    }
    catch {
        $results.Add([pscustomobject]@{
                SourcePath      = $sourcePath
                DestinationPath = if ($null -ne $destPath -and $destPath -ne '') { $destPath } else { '' }
                LastWriteTime   = $file.LastWriteTime
                Length          = $file.Length
                Status          = 'Failed'
                Message         = $_.Exception.Message
            })
        Write-Host "FAILED: $sourcePath - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Empty-folder cleanup only after a commit run
if ($Commit) {
    $doPrune = $false
    if ($PSBoundParameters.ContainsKey('RemoveEmptyFolders')) {
        $doPrune = ($RemoveEmptyFolders -eq 'Yes')
    }
    else {
        $answer = Read-Host "Remove empty folders under the input tree (never the input root)? [y/N]"
        $doPrune = ($answer -match '^(y|yes)$')
    }

    if ($doPrune) {
        $dirs = @(
            Get-ChildItem -LiteralPath $inputResolved -Recurse -Directory -ErrorAction SilentlyContinue |
            Sort-Object { $_.FullName.Length } -Descending
        )
        foreach ($dir in $dirs) {
            if ($dir.FullName.TrimEnd('\') -eq $inputResolved.TrimEnd('\')) {
                continue
            }
            try {
                $itemCount = (Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction Stop | Measure-Object).Count
                if ($itemCount -eq 0) {
                    Remove-Item -LiteralPath $dir.FullName -Force -ErrorAction Stop
                    Write-Host "Removed empty folder: $($dir.FullName)" -ForegroundColor DarkGray
                }
            }
            catch {
                Write-Warning "Could not remove folder $($dir.FullName): $($_.Exception.Message)"
            }
        }
    }
}

# Report output format
$outFormat = $Output
if (-not $PSBoundParameters.ContainsKey('Output') -or [string]::IsNullOrWhiteSpace($outFormat)) {
    do {
        $outFormat = Read-Host "Output format: type Text or CSV"
        $outFormat = $outFormat.Trim()
        switch -Regex ($outFormat) {
            '^(?i)text$' { $outFormat = 'Text' }
            '^(?i)csv$' { $outFormat = 'CSV' }
        }
    } while ($outFormat -notin 'Text', 'CSV')
}

if ($outFormat -eq 'Text') {
    $results | Format-Table -AutoSize -Property SourcePath, DestinationPath, LastWriteTime, Length, Status, Message
    Write-Host "Rows: $($results.Count)  (Version $script:Version)"
}
else {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvName = "ArchiveReport_$stamp.csv"
    $csvPath = Join-Path $archiveResolved $csvName
    $results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written: $csvPath" -ForegroundColor Cyan
    Write-Host "Rows: $($results.Count)  (Version $script:Version)"
}

exit 0
