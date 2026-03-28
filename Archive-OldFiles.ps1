<#
.SYNOPSIS
    Finds files older than N years under an input folder and plans or moves them to an archive root with mirrored paths.

.DESCRIPTION
    By default runs in preview mode: no files are moved and no folders are created on the archive.
    Use -Commit to perform moves. With -Commit, the archive folder is created if it does not exist.
    The archive path must not be inside the input tree. Preview mode requires the archive folder to already exist.

    Age uses LastWriteTime by default (Explorer "Date modified"). Other -AgeBasis values: LastAccessTime (last
    opened / NTFS last access; may be inaccurate if the volume disables last-access updates), LatestWriteOrAccess
    (newer of modified and last access, so a file is old only if both are before the cutoff), CreationTime, Earliest
    (older of creation vs last write; useful after copies/restores that refresh LastWriteTime).

    After a successful run, settings are saved to Archive-OldFiles.config.json next to the script (unless -NoSaveConfig).
    Each run also appends a short summary to Archive-OldFiles.run.log in the same folder (unless -NoRunLog).
    Optional HTML reports are written under the archive path; the file name includes computer/AD domain, the scanned path, and a timestamp.
    Use -All on a file server to preview every published folder share (HTML per share, never -Commit). Owner column resolves SIDs; unresolved or deleted accounts show as "No active user".

.PARAMETER InputPath
    Root folder to scan recursively. Optional if a valid saved config exists.

.PARAMETER ArchivePath
    Root folder for archived files. Optional if a valid saved config exists.

.PARAMETER Years
    Age threshold in years (number, decimals allowed). Passed as text to avoid binding errors on some hosts. Optional if a valid saved config exists.

.PARAMETER Commit
    Perform Move-Item and create destination directories.

.PARAMETER Output
    Text or HTML. HTML writes a formatted report under the archive folder. If omitted, you are prompted at the end (unless saved in config).

.PARAMETER RemoveEmptyFolders
    Yes or No after -Commit. If omitted, you are prompted (unless saved in config).

.PARAMETER ConfigPath
    Path to the JSON config file. Default: Archive-OldFiles.config.json next to this script.

.PARAMETER NoSaveConfig
    Do not write or update the JSON file after successful completion.

.PARAMETER ArchiveShareName
    On this server, the archive location is assumed to be the local path of this SMB share name (default: Archive).
    If that share does not exist, you are prompted for a full archive path.

.PARAMETER SkipShareMenu
    Do not show the published-share picker for InputPath; use prompts or parameters only.

.PARAMETER AgeBasis
    Timestamp compared to the cutoff: LastWriteTime (default, last modified); LastAccessTime (last accessed / opened);
    LatestWriteOrAccess (newer of last modified and last access); CreationTime; Earliest (older of creation vs modified).
    Aliases: Modified, Opened, LastOpened, ModifiedOrOpened, A/B/C (see Get-Help examples).

.PARAMETER NoRunLog
    Do not append a line to Archive-OldFiles.run.log next to the config file.

.PARAMETER All
    Scan every published disk share on this computer (same rules as the share picker: Type 0, no trailing $ in name).
    Always runs in preview mode: no moves are performed even if -Commit or saved JSON says commit. Implies HTML output:
    one report file per share under ArchivePath. InputPath is ignored. Requires ArchivePath and Years (from parameters or saved config).

.NOTES
    Purpose: Age-based archival of old files to a separate archive location.
    Author: Doug Hesseltine
    Created: 2026-03-27
    Modified: 2026-03-28
    Version: 1.7.2

    Troubleshooting (if the script will not run):
    - After downloading from the web, run: Unblock-File -Path .\Archive-OldFiles.ps1
    - Run explicitly: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Archive-OldFiles.ps1 -InputPath '...' -ArchivePath '...' -Years N
    - Ensure the file is named Archive-OldFiles.ps1 (not .txt). Open in Notepad and Save As, Encoding: UTF-8 if you see odd characters.
    - Settings: Archive-OldFiles.config.json next to the script. Run summary: Archive-OldFiles.run.log (same folder). HTML reports: archive root when you choose HTML.
    - ParserError (Unexpected token '}') after editing: replace the file with a full fresh copy from source; truncated saves break the script mid-block.
    - -All: use -All with no value (switch). Older copies used an untyped -All parameter and required -All:$true; update to v1.7.1+.

.LINK
    https://technologist.services/tools/archive-files/
    https://github.com/hesseltined/ServerCleanUpTools

.EXAMPLE
    .\Archive-OldFiles.ps1

    Uses saved config if present (you confirm or correct fields), otherwise prompts.

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7

    Preview only; skips saved-config prompt for the three core fields.

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7 -Commit -Output HTML

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7 -AgeBasis Earliest

    Uses the older of creation time and last write (common after file copies that refresh modified date).

.EXAMPLE
    .\Archive-OldFiles.ps1 -InputPath 'D:\Data' -ArchivePath '\\nas\archive' -Years 7 -AgeBasis LatestWriteOrAccess

    A file is too new to archive if either last modified or last access is on or after the cutoff.

.EXAMPLE
    .\Archive-OldFiles.ps1 -All -ArchivePath 'D:\Reports\ArchivePreviews' -Years 7

    Preview-only scan of all folder shares; writes one HTML report per share under the archive path (no -Commit).
#>

param(
    $InputPath,
    $ArchivePath,
    $Years,
    $Commit,
    $Output,
    $RemoveEmptyFolders,
    $ConfigPath,
    $NoSaveConfig,
    $ArchiveShareName,
    $SkipShareMenu,
    $AgeBasis,
    $NoRunLog,
    [switch]$All
)

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error 'This script requires Windows PowerShell 5.0 or later.'
    exit 1
}

# With no [Parameter(Position=...)], -File script.ps1 a b c leaves a,b,c in $args - map them if names were not used.
if ($args.Count -ge 1 -and [string]::IsNullOrWhiteSpace("$InputPath")) {
    $InputPath = $args[0]
}
if ($args.Count -ge 2 -and [string]::IsNullOrWhiteSpace("$ArchivePath")) {
    $ArchivePath = $args[1]
}
if ($args.Count -ge 3 -and [string]::IsNullOrWhiteSpace("$Years")) {
    $Years = $args[2]
}

# Years from 3rd positional argument is not recorded in PSBoundParameters; treat as supplied so we do not overwrite $yrs in prompts.
$yearsLockedFromInvocation = $PSBoundParameters.ContainsKey('Years') -or (($args.Count -ge 3) -and -not [string]::IsNullOrWhiteSpace([string]$args[2]))

function ConvertTo-BoundBool {
    param($x)
    if ($null -eq $x) {
        return $false
    }
    if ($x -is [bool]) {
        return [bool]$x
    }
    if ($x -is [switch]) {
        return $x.IsPresent
    }
    if ($x -is [System.Management.Automation.SwitchParameter]) {
        return $x.IsPresent
    }
    $s = "$x".Trim()
    if ($s -match '^(?i)true|yes|1$') {
        return $true
    }
    return $false
}

# Defaults and validation after binding. Param types are minimal except [switch]$All so -All works without a value (PS 5.1).
if (-not $PSBoundParameters.ContainsKey('ArchiveShareName') -or [string]::IsNullOrWhiteSpace("$ArchiveShareName")) {
    $ArchiveShareName = 'Archive'
}
else {
    $ArchiveShareName = "$ArchiveShareName".Trim()
}

if ($PSBoundParameters.ContainsKey('Output') -and -not [string]::IsNullOrWhiteSpace("$Output")) {
    $o = "$Output".Trim()
    if ($o -match '^(?i)text$') {
        $Output = 'Text'
    }
    elseif ($o -match '^(?i)html$') {
        $Output = 'HTML'
    }
    elseif ($o -match '^(?i)csv$') {
        $Output = 'HTML'
    }
    else {
        throw "-Output must be Text or HTML. Got: $Output"
    }
}

if ($PSBoundParameters.ContainsKey('RemoveEmptyFolders') -and -not [string]::IsNullOrWhiteSpace("$RemoveEmptyFolders")) {
    $r = "$RemoveEmptyFolders".Trim()
    if ($r -match '^(?i)yes$') {
        $RemoveEmptyFolders = 'Yes'
    }
    elseif ($r -match '^(?i)no$') {
        $RemoveEmptyFolders = 'No'
    }
    else {
        throw "-RemoveEmptyFolders must be Yes or No. Got: $RemoveEmptyFolders"
    }
}

function Normalize-AgeBasisParameter {
    param([string]$abRaw)
    if ([string]::IsNullOrWhiteSpace($abRaw)) {
        return $null
    }
    $ab = $abRaw.Trim()
    if ($ab -match '^(?i)lastwritetime$') {
        return 'LastWriteTime'
    }
    if ($ab -match '^(?i)(lastmodified|modified)$') {
        return 'LastWriteTime'
    }
    if ($ab -match '^(?i)a$') {
        return 'LastWriteTime'
    }
    if ($ab -match '^(?i)lastaccesstime$') {
        return 'LastAccessTime'
    }
    if ($ab -match '^(?i)(lastaccess|lastopened|opened)$') {
        return 'LastAccessTime'
    }
    if ($ab -match '^(?i)b$') {
        return 'LastAccessTime'
    }
    if ($ab -match '^(?i)latestwriteoraccess$') {
        return 'LatestWriteOrAccess'
    }
    if ($ab -match '^(?i)(modifiedoropened|writeoraccess|latestactivity)$') {
        return 'LatestWriteOrAccess'
    }
    if ($ab -match '^(?i)c$') {
        return 'LatestWriteOrAccess'
    }
    if ($ab -match '^(?i)creationtime$') {
        return 'CreationTime'
    }
    if ($ab -match '^(?i)earliest$') {
        return 'Earliest'
    }
    return $null
}

function Resolve-AgeBasisFromString {
    param([string]$Raw)
    return (Normalize-AgeBasisParameter -abRaw $Raw)
}

$ageBasisEffective = 'LastWriteTime'
if ($PSBoundParameters.ContainsKey('AgeBasis') -and -not [string]::IsNullOrWhiteSpace("$AgeBasis")) {
    $nab = Normalize-AgeBasisParameter -abRaw "$AgeBasis"
    if ($null -eq $nab) {
        throw "-AgeBasis not recognized. Use LastWriteTime, LastAccessTime, LatestWriteOrAccess, CreationTime, Earliest (aliases: Modified, Opened, ModifiedOrOpened, A/B/C). Got: $AgeBasis"
    }
    $ageBasisEffective = $nab
}

# Parse Years as a number (also when set from positional $args, not only when bound by name).
$yrs = 0.0
if (-not [string]::IsNullOrWhiteSpace("$Years")) {
    $yt = 0.0
    if (-not [double]::TryParse(("$Years").Trim(), [ref]$yt)) {
        throw "Cannot convert '-Years' to a number: $Years"
    }
    $yrs = $yt
}

$skipShareMenuFlag = ConvertTo-BoundBool $SkipShareMenu
$noSaveConfigFlag = ConvertTo-BoundBool $NoSaveConfig
$noRunLogFlag = ConvertTo-BoundBool $NoRunLog
$allSharesFlag = $All.IsPresent

$ErrorActionPreference = 'Stop'

$script:Version = '1.7.2'

function Get-DefaultConfigPath {
    if ($PSScriptRoot) {
        return (Join-Path $PSScriptRoot 'Archive-OldFiles.config.json')
    }
    if ($MyInvocation.MyCommand.Path) {
        return (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'Archive-OldFiles.config.json')
    }
    return (Join-Path (Get-Location).Path 'Archive-OldFiles.config.json')
}

function Normalize-UserPath {
    param([string]$Path)
    if ($null -eq $Path) {
        return ''
    }
    $p = $Path.Trim()
    if ($p.Length -ge 2 -and $p.StartsWith('"') -and $p.EndsWith('"')) {
        $p = $p.Trim('"').Trim()
    }
    if ($p.Length -ge 2 -and $p.StartsWith("'") -and $p.EndsWith("'")) {
        $p = $p.Trim("'").Trim()
    }
    return $p
}

function Get-PublishedDiskShares {
    $list = New-Object System.Collections.Generic.List[object]
    $shares = $null
    try {
        $shares = Get-CimInstance -ClassName Win32_Share -ErrorAction Stop
    }
    catch {
        try {
            $shares = Get-WmiObject -Class Win32_Share -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not list published shares (Win32_Share): $($_.Exception.Message)"
            return @()
        }
    }
    $shares = $shares |
        Where-Object {
            $_.Type -eq 0 -and
            $_.Name -notmatch '\$' -and
            -not [string]::IsNullOrWhiteSpace($_.Path)
        } |
        Sort-Object Name
    foreach ($s in $shares) {
        $unc = '\\{0}\{1}' -f $env:COMPUTERNAME, $s.Name
        $list.Add([pscustomobject]@{
                Name      = $s.Name
                LocalPath = $s.Path.TrimEnd('\')
                Unc       = $unc
            })
    }
    # PS 5.1 (e.g. Server 2016): return @($list) on List[object] can throw ArgumentException "Argument types do not match".
    return $list.ToArray()
}

function Get-LocalPathForShareName {
    param([string]$ShareName)
    $name = $ShareName.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        return $null
    }
    try {
        $all = Get-CimInstance -ClassName Win32_Share -ErrorAction Stop
    }
    catch {
        try {
            $all = Get-WmiObject -Class Win32_Share -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not resolve share name '$name': $($_.Exception.Message)"
            return $null
        }
    }
    $s = $all |
        Where-Object { $_.Type -eq 0 -and $_.Name -ieq $name } |
        Select-Object -First 1
    if ($null -ne $s -and -not [string]::IsNullOrWhiteSpace($s.Path)) {
        return $s.Path.TrimEnd('\')
    }
    return $null
}

function Invoke-InputPathShareMenu {
    $shares = @(Get-PublishedDiskShares)
    if ($shares.Count -eq 0) {
        Write-Host 'No published disk shares were returned (check permissions or use -SkipShareMenu).' -ForegroundColor Yellow
        return $null
    }
    Write-Host ''
    Write-Host 'Published folder shares on this server (select INPUT / source to scan):' -ForegroundColor Cyan
    $idx = 1
    foreach ($s in $shares) {
        Write-Host ('  {0}) {1}  ->  {2}  ({3})' -f $idx, $s.Name, $s.LocalPath, $s.Unc)
        $idx++
    }
    $otherNum = $idx
    Write-Host ('  {0}) Other (enter a full local or UNC path manually)' -f $otherNum)
    Write-Host ''
    while ($true) {
        $choice = Read-Host "Enter number (1-$otherNum)"
        $n = 0
        if (-not [int]::TryParse($choice, [ref]$n)) {
            Write-Host 'Enter a number from the list.' -ForegroundColor Red
            continue
        }
        if ($n -ge 1 -and $n -lt $otherNum) {
            return $shares[$n - 1].LocalPath
        }
        if ($n -eq $otherNum) {
            while ($true) {
                $p = Read-Host 'Enter full path for InputPath (local path on this server or UNC)'
                $p = Normalize-UserPath -Path $p
                if (-not [string]::IsNullOrWhiteSpace($p)) {
                    return $p
                }
                Write-Host 'Path cannot be empty.' -ForegroundColor Red
            }
        }
        Write-Host "Choose between 1 and $otherNum." -ForegroundColor Red
    }
}

function Read-SavedConfigFromDisk {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
        return ($raw | ConvertFrom-Json)
    }
    catch {
        Write-Warning "Could not read config file (ignored): $($_.Exception.Message)"
        return $null
    }
}

function Test-SingleConfigValidity {
    param(
        [string]$InputPath,
        [string]$ArchivePath,
        [double]$Years,
        [bool]$DoCommit,
        [bool]$AllSharesMode = $false
    )
    $issues = @()
    if (-not $AllSharesMode) {
        if ([string]::IsNullOrWhiteSpace("$InputPath")) {
            $issues += @{ Field = 'InputPath'; Message = 'InputPath is empty.' }
        }
        elseif (-not (Test-Path -LiteralPath $InputPath)) {
            $issues += @{ Field = 'InputPath'; Message = "InputPath does not exist: $InputPath" }
        }
    }

    if ([string]::IsNullOrWhiteSpace("$ArchivePath")) {
        $issues += @{ Field = 'ArchivePath'; Message = 'ArchivePath is empty.' }
    }
    elseif (-not (Test-Path -LiteralPath $ArchivePath)) {
        if (-not $DoCommit) {
            $issues += @{ Field = 'ArchivePath'; Message = 'ArchivePath does not exist. Create it for preview, or use -Commit to create it.' }
        }
    }

    if ($Years -le 0 -or [double]::IsNaN($Years)) {
        $issues += @{ Field = 'Years'; Message = 'Years must be a number greater than zero.' }
    }
    return [object[]]@($issues)
}

function Show-SavedConfigSummary {
    param(
        [string]$InputPath,
        [string]$ArchivePath,
        [double]$Years,
        [bool]$DoCommit,
        [object]$Output,
        [object]$RemoveEmptyFolders,
        [string]$ArchiveShareNameHint = '',
        [string]$AgeBasis = 'LastWriteTime'
    )
    Write-Host ''
    Write-Host '--- Saved / current configuration ---' -ForegroundColor Cyan
    Write-Host "  InputPath        : $InputPath"
    Write-Host "  ArchivePath      : $ArchivePath"
    if (-not [string]::IsNullOrWhiteSpace($ArchiveShareNameHint)) {
        Write-Host "  ArchiveShareName : $ArchiveShareNameHint"
    }
    Write-Host "  Years            : $Years"
    Write-Host "  AgeBasis         : $AgeBasis"
    Write-Host "  Commit           : $DoCommit"
    Write-Host ('  Output           : {0}' -f ($(if ($null -ne $Output -and $Output -ne '') { $Output } else { '(prompt at end)' })))
    Write-Host ('  RemoveEmptyFldrs : {0}' -f ($(if ($null -ne $RemoveEmptyFolders -and $RemoveEmptyFolders -ne '') { $RemoveEmptyFolders } else { '(prompt if commit)' })))
    Write-Host '-------------------------------------' -ForegroundColor Cyan
    Write-Host ''
}

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

function Get-FileAgeTimestamp {
    param(
        $FileInfo,
        [string]$Basis
    )
    switch ($Basis) {
        'CreationTime' {
            return $FileInfo.CreationTime
        }
        'LastAccessTime' {
            return $FileInfo.LastAccessTime
        }
        'LatestWriteOrAccess' {
            if ($FileInfo.LastWriteTime -gt $FileInfo.LastAccessTime) {
                return $FileInfo.LastWriteTime
            }
            return $FileInfo.LastAccessTime
        }
        'Earliest' {
            if ($FileInfo.LastWriteTime -lt $FileInfo.CreationTime) {
                return $FileInfo.LastWriteTime
            }
            return $FileInfo.CreationTime
        }
        default {
            return $FileInfo.LastWriteTime
        }
    }
}

function Resolve-OwnerDisplayName {
    param([string]$OwnerString)
    if ([string]::IsNullOrWhiteSpace($OwnerString)) {
        return ''
    }
    $s = $OwnerString.Trim()
    if ($s -match '^(?i)O:(S-1-5-[0-9\-]+)$') {
        $s = $Matches[1]
    }
    if ($s -match '^S-1-5-[0-9\-]+$') {
        try {
            $sidObj = New-Object System.Security.Principal.SecurityIdentifier($s)
            $nt = $sidObj.Translate([System.Security.Principal.NTAccount])
            return [string]$nt.Value
        }
        catch {
            return 'No active user'
        }
    }
    try {
        $ntAcc = New-Object System.Security.Principal.NTAccount($s)
        $sidObj2 = $ntAcc.Translate([System.Security.Principal.SecurityIdentifier])
        $ntBack = $sidObj2.Translate([System.Security.Principal.NTAccount])
        return [string]$ntBack.Value
    }
    catch {
        return 'No active user'
    }
}

function Get-FileOwnerForReport {
    param([string]$LiteralPath)
    if ([string]::IsNullOrWhiteSpace($LiteralPath)) {
        return ''
    }
    try {
        $acl = Get-Acl -LiteralPath $LiteralPath -ErrorAction Stop
        if ($null -eq $acl -or [string]::IsNullOrWhiteSpace($acl.Owner)) {
            return ''
        }
        return (Resolve-OwnerDisplayName -OwnerString ([string]$acl.Owner))
    }
    catch {
    }
    return ''
}

function Get-ReportDateColumnHeader {
    param([string]$Basis)
    switch ($Basis) {
        'LastAccessTime' {
            return 'Last opened (access)'
        }
        'LatestWriteOrAccess' {
            return 'Age date (max modified / access)'
        }
        'CreationTime' {
            return 'Created'
        }
        'Earliest' {
            return 'Older of created / modified'
        }
        default {
            return 'Last modified'
        }
    }
}

function Format-DataSize {
    param(
        [long]$Bytes
    )
    if ($Bytes -lt 0) {
        $Bytes = 0
    }
    $KB = [long]1024
    $MB = $KB * 1024
    $GB = $MB * 1024
    $TB = $GB * 1024
    if ($Bytes -ge $TB) {
        return ('{0:N2} TB' -f ($Bytes / [double]$TB))
    }
    if ($Bytes -ge $GB) {
        return ('{0:N2} GB' -f ($Bytes / [double]$GB))
    }
    if ($Bytes -ge $MB) {
        return ('{0:N2} MB' -f ($Bytes / [double]$MB))
    }
    if ($Bytes -ge $KB) {
        return ('{0:N2} KB' -f ($Bytes / [double]$KB))
    }
    return ('{0} bytes' -f $Bytes)
}

function ConvertTo-HtmlEncodedText {
    param([string]$Text)
    if ($null -eq $Text) {
        return ''
    }
    return (((($Text -replace '&', '&amp;') -replace '<', '&lt;') -replace '>', '&gt;') -replace '"', '&quot;') -replace '''', '&#39;'
}

function Normalize-SavedOutputFormat {
    param([object]$Raw)
    if ($null -eq $Raw) {
        return $null
    }
    $s = "$Raw".Trim()
    if ($s -match '^(?i)text$') {
        return 'Text'
    }
    if ($s -match '^(?i)html$') {
        return 'HTML'
    }
    if ($s -match '^(?i)csv$') {
        return 'HTML'
    }
    return $null
}

function Get-ReportDomainLabelForFilename {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($null -ne $cs.Domain) {
            $dom = [string]$cs.Domain
            if (-not [string]::IsNullOrWhiteSpace($dom) -and $dom -notmatch '^(?i)WORKGROUP$') {
                return $dom.Trim()
            }
        }
    }
    catch {
    }
    if (-not [string]::IsNullOrWhiteSpace($env:USERDNSDOMAIN)) {
        return $env:USERDNSDOMAIN.Trim()
    }
    $cn = $env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($cn)) {
        return 'HOST'
    }
    return $cn.Trim()
}

function Get-ArchiveHtmlReportFilename {
    param(
        [string]$DomainPart,
        [string]$InputResolved,
        [string]$ShareNameLabel = ''
    )
    $d = $DomainPart
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $d = $d.Replace([string]$c, '_')
    }
    $d = ($d -replace '\s+', '_').Trim('._')
    if ([string]::IsNullOrWhiteSpace($d)) {
        $d = 'DOMAIN'
    }
    if ($d.Length -gt 48) {
        $d = $d.Substring(0, 48)
    }

    $sharePart = $InputResolved.TrimEnd('\')
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        if ($c -ne [char]'\' -and $c -ne [char]':') {
            $sharePart = $sharePart.Replace([string]$c, '_')
        }
    }
    $sharePart = $sharePart -replace '\\', '-' -replace ':', ''
    while ($sharePart.Contains('--')) {
        $sharePart = $sharePart.Replace('--', '-')
    }
    $sharePart = $sharePart.Trim('-')
    if ([string]::IsNullOrWhiteSpace($sharePart)) {
        $sharePart = 'Source'
    }
    if ($sharePart.Length -gt 100) {
        $sharePart = $sharePart.Substring(0, 100).TrimEnd('-')
    }

    if (-not [string]::IsNullOrWhiteSpace($ShareNameLabel)) {
        $sn = $ShareNameLabel.Trim()
        foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
            $sn = $sn.Replace([string]$c, '_')
        }
        $sn = ($sn -replace '\s+', '_').Trim('._')
        if ([string]::IsNullOrWhiteSpace($sn)) {
            $sn = 'Share'
        }
        if ($sn.Length -gt 48) {
            $sn = $sn.Substring(0, 48).TrimEnd('_')
        }
        $sharePart = $sn + '-' + $sharePart
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return ('ArchiveReport_{0}_{1}_{2}.html' -f $d, $sharePart, $stamp)
}

function Write-ArchiveHtmlReport {
    param(
        [string]$LiteralPath,
        $ResultRows,
        [string]$InputResolved,
        [string]$ArchiveResolved,
        [double]$Years,
        [string]$AgeBasis,
        [datetime]$Cutoff,
        [bool]$Commit,
        [string]$ScriptVersion,
        [int]$FileScanCount,
        [int]$SkippedTooNew,
        [int]$PlannedMovedCount,
        [int]$FailedCount,
        [long]$ReclaimBytes,
        [string]$ReclaimDisplay,
        [string]$DomainLabel
    )

    $genAt = (Get-Date).ToString('F')
    $modeLabel = if ($Commit) { 'Commit (moves performed where successful)' } else { 'Preview (no moves)' }
    $serverName = $env:COMPUTERNAME
    if ([string]::IsNullOrWhiteSpace($serverName)) {
        $serverName = '(unknown)'
    }

    $inputRoot = $InputResolved.TrimEnd('\')
    $htmlParts = New-Object System.Collections.Generic.List[string]

    if ($null -eq $ResultRows -or $ResultRows.Count -eq 0) {
        $null = $htmlParts.Add('<tbody><tr><td colspan="4" class="empty">No files met the age rule for this run.</td></tr></tbody>')
    }
    else {
        $groups = @(
            $ResultRows |
            Group-Object { (Split-Path -Parent $_.SourcePath) } |
            Sort-Object { $_.Name }
        )
        $stripe = 0
        foreach ($g in $groups) {
            $folderFull = $g.Name
            if ([string]::IsNullOrWhiteSpace($folderFull)) {
                $folderFull = $inputRoot
            }
            $folderFull = $folderFull.TrimEnd('\')
            $fc = $g.Group.Count
            $isCollapsible = ($fc -gt 20)
            $hdrEnc = ConvertTo-HtmlEncodedText $folderFull
            $hdrExtra = if ($isCollapsible) {
                ' &mdash; ' + $fc + ' files (collapsed; hover anywhere in this folder block to show all, move pointer away to hide)'
            }
            else {
                ' &mdash; ' + $fc + ' file' + $(if ($fc -ne 1) { 's' })
            }
            $tbodyClass = if ($isCollapsible) { 'folder-group collapsible' } else { 'folder-group' }
            $null = $htmlParts.Add(('<tbody class="{0}">' -f $tbodyClass))
            $null = $htmlParts.Add(('<tr class="folder-hdr"><td colspan="4">{0}{1}</td></tr>' -f $hdrEnc, $hdrExtra))

            foreach ($r in @($g.Group | Sort-Object SourcePath)) {
                $leaf = [System.IO.Path]::GetFileName($r.SourcePath)
                if ([string]::IsNullOrWhiteSpace($leaf)) {
                    $leaf = $r.SourcePath
                }
                $lwt = $r.ComparedForAge.ToString('yyyy-MM-dd HH:mm')
                $len = '{0:N0}' -f $r.Length
                $own = $r.Owner
                if ([string]::IsNullOrWhiteSpace("$own")) {
                    $own = '-'
                }
                $isFail = ($r.Status -eq 'Failed')
                $tip = $r.SourcePath
                if ($isFail -and -not [string]::IsNullOrWhiteSpace($r.Message)) {
                    $tip = $tip + [Environment]::NewLine + 'Error: ' + $r.Message
                }
                $tipEnc = ConvertTo-HtmlEncodedText $tip
                $leafEnc = ConvertTo-HtmlEncodedText $leaf
                $ownEnc = ConvertTo-HtmlEncodedText $own
                if ($isFail) {
                    $rowClass = 'frow failed'
                }
                else {
                    $rowClass = 'frow stripe' + ($stripe % 2)
                    $stripe++
                }
                $hideStyle = if ($isCollapsible) { ' style="display:none"' } else { '' }
                $null = $htmlParts.Add(('<tr class="{0}" title="{1}"{2}><td class="fn">{3}</td><td class="own">{4}</td><td class="dt">{5}</td><td class="num">{6}</td></tr>' -f $rowClass, $tipEnc, $hideStyle, $leafEnc, $ownEnc, $lwt, $len))
            }
            $null = $htmlParts.Add('</tbody>')
        }
    }

    $tableBody = $htmlParts -join [Environment]::NewLine

    $css = @'
body { font-family: Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 0; background: #eef1f5; color: #1a1a1a; font-size: 13px; }
header { background: linear-gradient(135deg, #0d3b66 0%, #1b6ca8 100%); color: #fff; padding: 0.65rem 1rem 0.85rem; }
header h1 { margin: 0 0 0.4rem 0; font-size: 1.1rem; font-weight: 600; }
.head-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem 1.25rem; align-items: start; font-size: 0.8rem; }
.head-grid .col h3 { margin: 0 0 0.35rem 0; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.04em; opacity: 0.85; font-weight: 600; }
.head-grid dl { display: grid; grid-template-columns: 7.2rem 1fr; gap: 0.15rem 0.5rem; margin: 0; }
.head-grid dt { opacity: 0.88; font-weight: 500; }
.head-grid dd { margin: 0; word-break: break-word; line-height: 1.25; }
.wrap { max-width: 1200px; margin: 0 auto; padding: 0.75rem 1rem 1.25rem; }
.tip-top { font-size: 0.78rem; color: #334; background: #e3ecf7; border: 1px solid #c5d4e8; border-radius: 6px; padding: 0.5rem 0.65rem; margin: 0 0 0.65rem 0; line-height: 1.4; }
.card { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 0.75rem; overflow: hidden; }
.card h2 { margin: 0; padding: 0.45rem 0.75rem; font-size: 0.88rem; background: #e8eef4; border-bottom: 1px solid #d0dae6; }
.card .body { padding: 0.5rem 0.65rem; }
table { width: 100%; border-collapse: collapse; font-size: 0.78rem; table-layout: fixed; }
th { text-align: left; background: #0d3b66; color: #fff; padding: 0.35rem 0.45rem; font-weight: 600; }
th:nth-child(1) { width: 36%; }
th:nth-child(2) { width: 22%; }
th:nth-child(3) { width: 18%; }
th:nth-child(4) { width: 24%; }
td { padding: 0.28rem 0.45rem; vertical-align: middle; border-bottom: 1px solid #dde3ea; line-height: 1.2; }
tbody.folder-group { border-bottom: 2px solid #b8c5d4; }
tbody.folder-group.collapsible { cursor: default; background: #fdf6ec; border-left: 4px solid #c9943a; }
tbody.folder-group.collapsible tr.folder-hdr td { background: #e8d4b0; color: #3d3010; border-bottom: 1px solid #d4b78a; }
tbody.folder-group.collapsible tr.frow.stripe0 td { background: #faf3e6; }
tbody.folder-group.collapsible tr.frow.stripe1 td { background: #f3e8d2; }
tr.folder-hdr td { background: #d5dee8; font-weight: 600; color: #0d3b66; padding: 0.4rem 0.45rem; border-bottom: 1px solid #b8c5d4; font-size: 0.76rem; word-break: break-word; }
tr.frow.stripe0 td { background: #f4f5f7; }
tr.frow.stripe1 td { background: #e8f1fb; }
tr.frow.failed td { background: #f5d4d4 !important; color: #6b1212; font-weight: 500; }
td.fn { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
td.own { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #333; font-size: 0.74rem; }
td.dt { white-space: nowrap; color: #333; }
td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.empty { text-align: center; color: #666; padding: 1rem !important; }
'@

    $js = @'
<script>
(function () {
  document.querySelectorAll('tbody.folder-group.collapsible').forEach(function (tb) {
    function showAll() {
      tb.querySelectorAll('tr.frow').forEach(function (tr) { tr.style.display = 'table-row'; });
    }
    function hideFiles() {
      tb.querySelectorAll('tr.frow').forEach(function (tr) { tr.style.display = 'none'; });
    }
    tb.addEventListener('mouseenter', showAll);
    tb.addEventListener('mouseleave', hideFiles);
  });
})();
</script>
'@

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Archive Old Files Report</title>
<style>$css</style>
</head>
<body>
<header>
  <h1>Archive Old Files &mdash; Report</h1>
  <div class="head-grid">
    <div class="col job">
      <h3>Job</h3>
      <dl>
        <dt>Server</dt><dd>$(ConvertTo-HtmlEncodedText $serverName)</dd>
        <dt>Domain</dt><dd>$(ConvertTo-HtmlEncodedText $DomainLabel)</dd>
        <dt>Source</dt><dd>$(ConvertTo-HtmlEncodedText $InputResolved)</dd>
        <dt>Archive</dt><dd>$(ConvertTo-HtmlEncodedText $ArchiveResolved)</dd>
        <dt>Older than</dt><dd>$(ConvertTo-HtmlEncodedText ('{0} yr, basis {1}' -f $Years, $AgeBasis))</dd>
        <dt>Cutoff</dt><dd>$(ConvertTo-HtmlEncodedText $Cutoff.ToString('g'))</dd>
        <dt>Mode</dt><dd>$(ConvertTo-HtmlEncodedText $modeLabel)</dd>
      </dl>
    </div>
    <div class="col results">
      <h3>Results</h3>
      <dl>
        <dt>Generated</dt><dd>$(ConvertTo-HtmlEncodedText $genAt)</dd>
        <dt>Script</dt><dd>v$ScriptVersion</dd>
        <dt>Scanned</dt><dd>$FileScanCount files</dd>
        <dt>Too new</dt><dd>$SkippedTooNew</dd>
        <dt>Met age rule</dt><dd>$($ResultRows.Count) total</dd>
        <dt>Planned / moved</dt><dd>$PlannedMovedCount</dd>
        <dt>Failed</dt><dd>$FailedCount</dd>
        <dt>Size off source</dt><dd><strong>$(ConvertTo-HtmlEncodedText $ReclaimDisplay)</strong> ($ReclaimBytes B)</dd>
      </dl>
    </div>
  </div>
</header>
<div class="wrap">
  <p class="tip-top">Hover a <strong>file name</strong> for the full source path. <strong>Red</strong> rows failed to move or plan; hover the row for the error. Folders with <strong>more than 20</strong> files start with file lines hidden; move the pointer into that folder&rsquo;s block (header or rows) to show them, and move away to collapse again.</p>
  <div class="card">
    <h2>Files (by folder)</h2>
    <div class="body" style="overflow-x:auto;">
      <table>
        <thead>
          <tr>
            <th>File</th>
            <th>Owner</th>
            <th>$(ConvertTo-HtmlEncodedText (Get-ReportDateColumnHeader -Basis $AgeBasis))</th>
            <th>Size (bytes)</th>
          </tr>
        </thead>
$tableBody
      </table>
    </div>
  </div>
</div>
$js
</body>
</html>
"@

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($LiteralPath, $html, $utf8NoBom)
}

# Single-path scan, age filter, optional moves, and report output. Also invoked once per share when -All is used.
function Invoke-SingleArchiveJob {
    param(
        [string]$InputResolved,
        [string]$ArchiveResolved,
        [double]$YearsNum,
        [string]$AgeBasisEff,
        [bool]$DoCommit,
        [string]$OutFormatIn,
        [string]$ShareNameLabel = '',
        [bool]$BoundRemoveEmptyFolders = $false,
        [string]$RemoveEmptyFoldersParam = '',
        [string]$RemoveEmptyPref = ''
    )

    $inputResolved = $InputResolved.TrimEnd('\')
    $archiveResolved = $ArchiveResolved.TrimEnd('\')
    $ageBasisEffective = $AgeBasisEff
    $doCommit = $DoCommit
    $yrs = $YearsNum

    $cutoff = (Get-Date).AddYears(-$yrs)
    Write-Host ''
    Write-Host '========== ARCHIVE RUN ==========' -ForegroundColor Cyan
    Write-Host ('  Source tree (InputPath):  {0}' -f $inputResolved)
    Write-Host ('  Archive root:             {0}' -f $archiveResolved)
    Write-Host ('  Age rule: older than {0} year(s); using {1} (must be strictly before {2}).' -f $yrs, $ageBasisEffective, $cutoff)
    if ($ageBasisEffective -eq 'LastAccessTime') {
        Write-Host '  Note: LastAccessTime depends on NTFS last-access tracking; if the volume disables it, dates may look stale.' -ForegroundColor DarkGray
    }
    Write-Host ('  Mode: {0}' -f $(if ($doCommit) { 'COMMIT - files will be moved to the archive' } else { 'PREVIEW - no moves; plan only' }))
    Write-Host '=================================' -ForegroundColor Cyan
    Write-Host ''

    $topLevelEntryCount = -1
    $depth1FileCount = -1
    $topLevelFolderCount = -1
    try {
        $topChildren = @(Get-ChildItem -LiteralPath $inputResolved -Force -ErrorAction Stop)
        $topLevelEntryCount = $topChildren.Count
        $depth1FileCount = @($topChildren | Where-Object { -not $_.PSIsContainer }).Count
        $topLevelFolderCount = @($topChildren | Where-Object { $_.PSIsContainer }).Count
    }
    catch {
        throw "Could not list the root of InputPath (before recursive scan). Path: $inputResolved  $($_.Exception.Message)"
    }

    Write-Host '--- Step 1: Can we read the source folder? (top level only) ---' -ForegroundColor DarkGray
    Write-Host ("  Entries at root: {0} ({1} files, {2} subfolders here)." -f $topLevelEntryCount, $depth1FileCount, $topLevelFolderCount) -ForegroundColor DarkGray
    Write-Host ''

    $results = New-Object System.Collections.Generic.List[object]

    $files = @(
        Get-ChildItem -LiteralPath $inputResolved -Recurse -File -Force -ErrorAction SilentlyContinue -ErrorVariable gciErrors
    )

    $gciErrCount = 0
    if ($null -ne $gciErrors) {
        $gciErrCount = @($gciErrors).Count
    }
    if ($gciErrCount -gt 0) {
        Write-Host ""
        Write-Host "RECURSIVE SCAN REPORTED $gciErrCount ERROR(S) - results may be incomplete (0 files can mean everything below was blocked). First errors:" -ForegroundColor Red
        $errArr = @($gciErrors)
        $showErr = [Math]::Min(8, $gciErrCount)
        for ($ei = 0; $ei -lt $showErr; $ei++) {
            Write-Host "  [$($ei + 1)] $($errArr[$ei].Exception.Message)" -ForegroundColor Red
        }
        if ($gciErrCount -gt $showErr) {
            Write-Host "  ... and $($gciErrCount - $showErr) more (see Warning stream for full list)." -ForegroundColor Red
        }
        Write-Host ""
        foreach ($err in $errArr) {
            Write-Warning "Enumeration issue: $($err.Exception.Message)"
        }
    }

    $firstGciErrorLog = ''
    if ($gciErrCount -gt 0) {
        $em = [string](@($gciErrors)[0].Exception.Message)
        $em = $em -replace '[\r\n]+', ' '
        if ($em.Length -gt 240) {
            $em = $em.Substring(0, 237) + '...'
        }
        $firstGciErrorLog = $em
    }

    $fileScanCount = $files.Count
    Write-Host '--- Step 2: Recursive file list under source ---' -ForegroundColor DarkGray
    Write-Host ("  Total files found (all ages): {0}" -f $fileScanCount) -ForegroundColor DarkGray
    if ($fileScanCount -eq 0) {
        Write-Host '  No files returned by Get-ChildItem -Recurse -File.' -ForegroundColor Yellow
        Write-Host '  Check: empty tree, permissions on subfolders, or wrong InputPath.' -ForegroundColor Yellow
    }
    Write-Host ''

    Write-Host '--- Step 3: Compare each file to the cutoff (older files are listed below) ---' -ForegroundColor DarkGray
    Write-Host ''

    $skippedTooNew = 0
    foreach ($file in $files) {
        $sourcePath = $file.FullName
        $destPath = $null
        $comparedForAge = Get-FileAgeTimestamp -FileInfo $file -Basis $ageBasisEffective
        try {
            if ($comparedForAge -ge $cutoff) {
                $skippedTooNew++
                continue
            }

            $relative = Get-RelativePathFromRoot -FileFullName $sourcePath -RootFullName $inputResolved
            $initialDest = Join-Path $archiveResolved $relative
            $destPath = Get-UniqueDestinationFilePath -InitialDestPath $initialDest

            if (-not $doCommit) {
                $ownerSnap = Get-FileOwnerForReport -LiteralPath $sourcePath
                $results.Add([pscustomobject]@{
                        SourcePath       = $sourcePath
                        DestinationPath  = $destPath
                        ComparedForAge   = $comparedForAge
                        LastWriteTime    = $file.LastWriteTime
                        CreationTime     = $file.CreationTime
                        Length           = $file.Length
                        Owner            = $ownerSnap
                        Status           = 'Planned'
                        Message          = ''
                    })
                continue
            }

            $destDir = Split-Path -Parent $destPath
            if (-not (Test-Path -LiteralPath $destDir)) {
                $null = New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop
            }

            $ownerSnap = Get-FileOwnerForReport -LiteralPath $sourcePath
            Move-Item -LiteralPath $sourcePath -Destination $destPath -Force -ErrorAction Stop
            $results.Add([pscustomobject]@{
                    SourcePath       = $sourcePath
                    DestinationPath  = $destPath
                    ComparedForAge   = $comparedForAge
                    LastWriteTime    = $file.LastWriteTime
                    CreationTime     = $file.CreationTime
                    Length           = $file.Length
                    Owner            = $ownerSnap
                    Status           = 'Moved'
                    Message          = ''
                })
        }
        catch {
            $ownerSnap = Get-FileOwnerForReport -LiteralPath $sourcePath
            $results.Add([pscustomobject]@{
                    SourcePath       = $sourcePath
                    DestinationPath  = if ($null -ne $destPath -and $destPath -ne '') { $destPath } else { '' }
                    ComparedForAge   = $comparedForAge
                    LastWriteTime    = $file.LastWriteTime
                    CreationTime     = $file.CreationTime
                    Length           = $file.Length
                    Owner            = $ownerSnap
                    Status           = 'Failed'
                    Message          = $_.Exception.Message
                })
            Write-Host "FAILED: $sourcePath - $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $plannedMovedRows = @($results | Where-Object { $_.Status -in 'Planned', 'Moved' })
    $failedRows = @($results | Where-Object { $_.Status -eq 'Failed' })
    $reclaimBytes = [long]0
    $sumObj = $plannedMovedRows | Measure-Object -Property Length -Sum
    if ($null -ne $sumObj -and $null -ne $sumObj.Sum) {
        $reclaimBytes = [long]$sumObj.Sum
    }
    $reclaimDisplay = Format-DataSize -Bytes $reclaimBytes

    Write-Host ''
    Write-Host '--- Step 4: Age filter summary ---' -ForegroundColor DarkGray
    Write-Host ("  Files skipped (too new, on or after cutoff): {0}" -f $skippedTooNew) -ForegroundColor DarkGray
    Write-Host ("  Files listed for archive (met age rule):     {0}" -f $results.Count) -ForegroundColor DarkGray
    Write-Host ("    - Planned or Moved (success path):         {0}" -f $plannedMovedRows.Count) -ForegroundColor DarkGray
    Write-Host ("    - Failed (met age but error on move/plan): {0}" -f $failedRows.Count) -ForegroundColor DarkGray
    Write-Host ("  Total size of Planned+Moved (space no longer on source after a successful move): {0} ({1} bytes)" -f $reclaimDisplay, $reclaimBytes) -ForegroundColor DarkGray
    Write-Host ''

    $removeEmptyChoice = $null
    if ($doCommit) {
        $doPrune = $false
        if ($BoundRemoveEmptyFolders) {
            $doPrune = ($RemoveEmptyFoldersParam -eq 'Yes')
        }
        elseif ($RemoveEmptyPref -in 'Yes', 'No') {
            $doPrune = ($RemoveEmptyPref -eq 'Yes')
        }
        else {
            $answer = Read-Host "Remove empty folders under the input tree (never the input root)? [y/N]"
            $doPrune = ($answer -match '^(y|yes)$')
        }

        if ($doPrune) {
            $removeEmptyChoice = 'Yes'
            $dirs = @(
                Get-ChildItem -LiteralPath $inputResolved -Recurse -Directory -Force -ErrorAction SilentlyContinue |
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
        else {
            $removeEmptyChoice = 'No'
        }
    }

    $outFormat = Normalize-SavedOutputFormat -Raw $OutFormatIn
    if ($null -eq $outFormat) {
        $outFormat = ''
    }
    if ($outFormat -notin 'Text', 'HTML') {
        do {
            $outFormat = Read-Host "Output format: type Text or HTML"
            $outFormat = $outFormat.Trim()
            switch -Regex ($outFormat) {
                '^(?i)text$' { $outFormat = 'Text' }
                '^(?i)html$' { $outFormat = 'HTML' }
                '^(?i)csv$' { $outFormat = 'HTML' }
            }
        } while ($outFormat -notin 'Text', 'HTML')
    }

    $htmlReportOutPath = $null
    if ($outFormat -eq 'Text') {
        Write-Host ''
        Write-Host '========== DETAIL: FILES THAT MET THE AGE RULE ==========' -ForegroundColor Cyan
        if ($results.Count -eq 0) {
            Write-Host '  (none)' -ForegroundColor DarkGray
        }
        else {
            $results | Format-Table -AutoSize -Property SourcePath, Owner, ComparedForAge, LastWriteTime, Length, Status, Message
        }
        Write-Host '=========================================================' -ForegroundColor Cyan
    }
    else {
        $domainForReport = Get-ReportDomainLabelForFilename
        $htmlBaseName = Get-ArchiveHtmlReportFilename -DomainPart $domainForReport -InputResolved $inputResolved -ShareNameLabel $ShareNameLabel
        $htmlReportOutPath = Join-Path $archiveResolved $htmlBaseName
        Write-ArchiveHtmlReport -LiteralPath $htmlReportOutPath -ResultRows $results -InputResolved $inputResolved -ArchiveResolved $archiveResolved `
            -Years $yrs -AgeBasis $ageBasisEffective -Cutoff $cutoff -Commit $doCommit -ScriptVersion $script:Version `
            -FileScanCount $fileScanCount -SkippedTooNew $skippedTooNew -PlannedMovedCount $plannedMovedRows.Count -FailedCount $failedRows.Count `
            -ReclaimBytes $reclaimBytes -ReclaimDisplay $reclaimDisplay -DomainLabel $domainForReport
        Write-Host ''
        Write-Host "HTML report written: $htmlReportOutPath" -ForegroundColor Cyan
        Write-Host ("  Rows (met age rule): {0}" -f $results.Count) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host '==============================================================================' -ForegroundColor Cyan
    Write-Host ' FINAL SUMMARY (this source)' -ForegroundColor Cyan
    Write-Host '==============================================================================' -ForegroundColor Cyan
    Write-Host ('  Script version:          {0}' -f $script:Version)
    Write-Host ('  InputPath (resolved):    {0}' -f $inputResolved)
    Write-Host ('  ArchivePath (resolved):  {0}' -f $archiveResolved)
    Write-Host ('  Years / AgeBasis:        {0} / {1}' -f $yrs, $ageBasisEffective)
    Write-Host ('  Cutoff (exclusive):      {0}' -f $cutoff)
    Write-Host ('  Mode:                    {0}' -f $(if ($doCommit) { 'Commit (moves performed where successful)' } else { 'Preview (no moves)' }))
    Write-Host ('  Output format this run:  {0}' -f $outFormat)
    if ($outFormat -eq 'HTML' -and $null -ne $htmlReportOutPath) {
        Write-Host ('  HTML report path:        {0}' -f $htmlReportOutPath)
    }
    Write-Host ('  Files scanned (all ages): {0}' -f $fileScanCount)
    Write-Host ('  Skipped (too new):        {0}' -f $skippedTooNew)
    Write-Host ('  Listed (met age rule):    {0}' -f $results.Count)
    Write-Host ('    Planned + Moved:        {0}' -f $plannedMovedRows.Count)
    Write-Host ('    Failed:                 {0}' -f $failedRows.Count)
    Write-Host ('  Source space from Planned+Moved: {0} ({1} bytes)' -f $reclaimDisplay, $reclaimBytes)
    Write-Host '    (After commit, this much file data no longer lives under the source tree.)' -ForegroundColor DarkGray
    Write-Host '==============================================================================' -ForegroundColor Cyan
    Write-Host ''

    return [pscustomobject]@{
        InputResolved          = $inputResolved
        ArchiveResolved        = $archiveResolved
        Results                = $results.ToArray()
        FileScanCount          = $fileScanCount
        SkippedTooNew          = $skippedTooNew
        PlannedMovedCount      = $plannedMovedRows.Count
        FailedCount            = $failedRows.Count
        ReclaimBytes           = $reclaimBytes
        ReclaimDisplay         = $reclaimDisplay
        HtmlReportPath         = $htmlReportOutPath
        OutFormat              = $outFormat
        TopLevelEntryCount     = $topLevelEntryCount
        Depth1FileCount        = $depth1FileCount
        TopLevelFolderCount    = $topLevelFolderCount
        GciErrorCount          = $gciErrCount
        FirstGciErrorLog       = $firstGciErrorLog
        DoCommitUsed           = $doCommit
        RemoveEmptyChoice      = $removeEmptyChoice
        Cutoff                 = $cutoff
        AgeBasisEffective      = $ageBasisEffective
        YearsNum               = $yrs
    }
}

# --- Resolve configuration: command line, then Archive-OldFiles.config.json, then interactive prompts ---

$configFile = if ($ConfigPath) { $ConfigPath } else { Get-DefaultConfigPath }

$coreProvidedOnCli = (-not [string]::IsNullOrWhiteSpace("$ArchivePath")) -and
(-not [string]::IsNullOrWhiteSpace("$Years")) -and ($yrs -gt 0) -and (
    $allSharesFlag -or ((-not [string]::IsNullOrWhiteSpace("$InputPath")))
)

$in = $InputPath
$arch = $ArchivePath
$doCommit = ConvertTo-BoundBool $Commit
$outPref = $Output
$removePref = $RemoveEmptyFolders
$effectiveArchiveShareName = $ArchiveShareName

if ($allSharesFlag -and -not [string]::IsNullOrWhiteSpace("$InputPath")) {
    Write-Host 'Note: -All ignores InputPath; each published disk share is scanned in preview-only mode.' -ForegroundColor Yellow
}

if ($coreProvidedOnCli) {
    Write-Host 'Core parameters supplied on command line; skipping saved-config load for InputPath, ArchivePath, and Years.' -ForegroundColor DarkGray
    $savedSide = Read-SavedConfigFromDisk -Path $configFile
    if ($null -ne $savedSide) {
        $sideNames = @($savedSide.PSObject.Properties | ForEach-Object { $_.Name })
        if (-not $PSBoundParameters.ContainsKey('ArchiveShareName') -and ($sideNames -contains 'ArchiveShareName') -and -not [string]::IsNullOrWhiteSpace($savedSide.ArchiveShareName)) {
            $effectiveArchiveShareName = [string]$savedSide.ArchiveShareName.Trim()
        }
    }
    if (-not $PSBoundParameters.ContainsKey('Commit') -and -not $allSharesFlag -and $null -ne $savedSide -and $null -ne $savedSide.PSObject.Properties['Commit']) {
        $doCommit = [bool]$savedSide.Commit
    }
    if (-not $PSBoundParameters.ContainsKey('Output') -and $null -ne $savedSide) {
        $normSideOut = Normalize-SavedOutputFormat -Raw $savedSide.Output
        if ($null -ne $normSideOut) {
            $outPref = $normSideOut
        }
    }
    if (-not $PSBoundParameters.ContainsKey('RemoveEmptyFolders') -and $null -ne $savedSide -and $savedSide.RemoveEmptyFolders -in 'Yes', 'No') {
        $removePref = [string]$savedSide.RemoveEmptyFolders
    }
    if (-not $PSBoundParameters.ContainsKey('AgeBasis') -and $null -ne $savedSide) {
        $pnAge = @($savedSide.PSObject.Properties | ForEach-Object { $_.Name })
        if ($pnAge -contains 'AgeBasis') {
            $resolvedAb = Resolve-AgeBasisFromString -Raw ([string]$savedSide.AgeBasis)
            if ($null -ne $resolvedAb) {
                $ageBasisEffective = $resolvedAb
            }
        }
    }
}
else {
    $saved = Read-SavedConfigFromDisk -Path $configFile
    if ($null -ne $saved) {
        $savedPropNames = @($saved.PSObject.Properties | ForEach-Object { $_.Name })
        if (-not $PSBoundParameters.ContainsKey('ArchiveShareName') -and ($savedPropNames -contains 'ArchiveShareName') -and -not [string]::IsNullOrWhiteSpace($saved.ArchiveShareName)) {
            $effectiveArchiveShareName = [string]$saved.ArchiveShareName.Trim()
        }
    }

    if (-not [string]::IsNullOrWhiteSpace("$InputPath")) {
        $in = $InputPath
    }
    if (-not [string]::IsNullOrWhiteSpace("$ArchivePath")) {
        $arch = $ArchivePath
    }
    if (-not [string]::IsNullOrWhiteSpace("$Years")) {
        $yt = 0.0
        if ([double]::TryParse(("$Years").Trim(), [ref]$yt)) {
            $yrs = $yt
        }
    }
    if ($PSBoundParameters.ContainsKey('Commit')) {
        $doCommit = ConvertTo-BoundBool $Commit
    }
    if ($PSBoundParameters.ContainsKey('Output')) {
        $outPref = $Output
    }
    if ($PSBoundParameters.ContainsKey('RemoveEmptyFolders')) {
        $removePref = $RemoveEmptyFolders
    }

    $useSaved = $false
    if ($null -ne $saved) {
        $archShareHint = ''
        $pnSaved = @($saved.PSObject.Properties | ForEach-Object { $_.Name })
        if (($pnSaved -contains 'ArchiveShareName') -and -not [string]::IsNullOrWhiteSpace($saved.ArchiveShareName)) {
            $archShareHint = [string]$saved.ArchiveShareName
        }
        $savedAgeHint = 'LastWriteTime'
        if (($pnSaved -contains 'AgeBasis') -and -not [string]::IsNullOrWhiteSpace($saved.AgeBasis)) {
            $tryAb = Resolve-AgeBasisFromString -Raw ([string]$saved.AgeBasis)
            if ($null -ne $tryAb) {
                $savedAgeHint = $tryAb
            }
        }
        Show-SavedConfigSummary -InputPath $saved.InputPath -ArchivePath $saved.ArchivePath -Years ([double]$saved.Years) -DoCommit ([bool]$saved.Commit) -Output $saved.Output -RemoveEmptyFolders $saved.RemoveEmptyFolders -ArchiveShareNameHint $archShareHint -AgeBasis $savedAgeHint
        $ans = Read-Host 'Use these saved values (you will fix any that are invalid next)? [Y/n]'
        if ($ans -notmatch '^(n|no)$') {
            $useSaved = $true
        }
    }

    if ($useSaved) {
        if ([string]::IsNullOrWhiteSpace("$InputPath")) {
            $in = [string]$saved.InputPath
        }
        if ([string]::IsNullOrWhiteSpace("$ArchivePath")) {
            $arch = [string]$saved.ArchivePath
        }
        if ([string]::IsNullOrWhiteSpace("$Years")) {
            $yrs = [double]$saved.Years
        }
        if (-not $PSBoundParameters.ContainsKey('Commit') -and -not $allSharesFlag) {
            $doCommit = [bool]$saved.Commit
        }
        if (-not ($PSBoundParameters.ContainsKey('Output') -and $Output -in 'Text', 'HTML')) {
            $normUseOut = Normalize-SavedOutputFormat -Raw $saved.Output
            if ($null -ne $normUseOut -and [string]::IsNullOrWhiteSpace($outPref)) {
                $outPref = $normUseOut
            }
        }
        if (-not ($PSBoundParameters.ContainsKey('RemoveEmptyFolders') -and $RemoveEmptyFolders -in 'Yes', 'No') -and $saved.RemoveEmptyFolders -in 'Yes', 'No') {
            if ([string]::IsNullOrWhiteSpace($removePref)) {
                $removePref = [string]$saved.RemoveEmptyFolders
            }
        }
        if (-not $PSBoundParameters.ContainsKey('AgeBasis')) {
            if (($pnSaved -contains 'AgeBasis') -and -not [string]::IsNullOrWhiteSpace($saved.AgeBasis)) {
                $resolvedUse = Resolve-AgeBasisFromString -Raw ([string]$saved.AgeBasis)
                if ($null -ne $resolvedUse) {
                    $ageBasisEffective = $resolvedUse
                }
            }
        }
    }
    else {
        if (-not $allSharesFlag) {
            if ([string]::IsNullOrWhiteSpace($in)) {
                $pickedIn = $null
                if (-not $skipShareMenuFlag) {
                    $pickedIn = Invoke-InputPathShareMenu
                }
                if (-not [string]::IsNullOrWhiteSpace($pickedIn)) {
                    $in = Normalize-UserPath -Path $pickedIn
                }
                else {
                    $def = if ($null -ne $saved) { $saved.InputPath } else { '' }
                    $prompt = if ($def) { "InputPath [$def]" } else { 'InputPath' }
                    $r = Read-Host $prompt
                    $in = if ([string]::IsNullOrWhiteSpace($r)) { $def } else { $r }
                    $in = Normalize-UserPath -Path $in
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($arch)) {
            $fromArchiveShare = Get-LocalPathForShareName -ShareName $effectiveArchiveShareName
            if (-not [string]::IsNullOrWhiteSpace($fromArchiveShare)) {
                $arch = $fromArchiveShare
                Write-Host ("Using archive share '{0}' -> {1}" -f $effectiveArchiveShareName, $arch) -ForegroundColor Green
            }
            else {
                Write-Host ("No share named '{0}' was found on this computer. Specify the archive folder manually." -f $effectiveArchiveShareName) -ForegroundColor Yellow
                $def = if ($null -ne $saved) { $saved.ArchivePath } else { '' }
                $prompt = if ($def) { "ArchivePath [$def]" } else { 'ArchivePath (full local or UNC path)' }
                $r = Read-Host $prompt
                $arch = if ([string]::IsNullOrWhiteSpace($r)) { $def } else { $r }
                $arch = Normalize-UserPath -Path $arch
            }
        }
        if ((-not $yearsLockedFromInvocation) -or $yrs -le 0) {
            $defY = if ($null -ne $saved) { [double]$saved.Years } else { 0 }
            $prompt = if ($defY -gt 0) { "Years [$defY]" } else { 'Years' }
            $r = Read-Host $prompt
            if ([string]::IsNullOrWhiteSpace($r)) {
                $yrs = $defY
            }
            else {
                $parsedY = 0.0
                if ([double]::TryParse($r, [ref]$parsedY) -and $parsedY -gt 0) {
                    $yrs = $parsedY
                }
                else {
                    $yrs = $defY
                }
            }
        }
        if (-not $PSBoundParameters.ContainsKey('Commit') -and -not $allSharesFlag) {
            $defC = if ($null -ne $saved) { [bool]$saved.Commit } else { $false }
            $r = Read-Host "Run with -Commit (actually move files)? Current default [$defC] (y/N)"
            if ($r -match '^(y|yes)$') {
                $doCommit = $true
            }
            elseif ($r -match '^(n|no)$') {
                $doCommit = $false
            }
            else {
                $doCommit = $defC
            }
        }
        if (-not $PSBoundParameters.ContainsKey('Output') -and -not $allSharesFlag) {
            $defO = ''
            if ($null -ne $saved) {
                $normDef = Normalize-SavedOutputFormat -Raw $saved.Output
                if ($null -ne $normDef) {
                    $defO = $normDef
                }
            }
            $prompt = if ($defO) { "Default report format Text or HTML (blank = prompt at end) [$defO]" } else { 'Default report format Text or HTML (blank = prompt at end)' }
            $r = Read-Host $prompt
            if ([string]::IsNullOrWhiteSpace($r)) {
                $outPref = $defO
            }
            elseif ($r -match '^(?i)text$') {
                $outPref = 'Text'
            }
            elseif ($r -match '^(?i)html$') {
                $outPref = 'HTML'
            }
            elseif ($r -match '^(?i)csv$') {
                $outPref = 'HTML'
            }
            else {
                $outPref = $defO
            }
        }
        if (-not $PSBoundParameters.ContainsKey('RemoveEmptyFolders') -and $doCommit -and -not $allSharesFlag) {
            if ($null -ne $saved -and $saved.RemoveEmptyFolders -in 'Yes', 'No') {
                $defRm = [string]$saved.RemoveEmptyFolders
            }
            else { $defRm = '' }
            $prompt = if ($defRm) { "Remove empty folders after commit Yes/No (blank = prompt later) [$defRm]" } else { 'Remove empty folders after commit Yes/No (blank = prompt later)' }
            $r = Read-Host $prompt
            if ([string]::IsNullOrWhiteSpace($r)) {
                $removePref = $defRm
            }
            elseif ($r -in 'Yes', 'No', 'yes', 'no', 'y', 'n') {
                $removePref = if ($r -match '^(Yes|y|yes)$') { 'Yes' } else { 'No' }
            }
        }
        if (-not $PSBoundParameters.ContainsKey('AgeBasis')) {
            $defAb = 'LastWriteTime'
            if ($null -ne $saved) {
                $pnAb = @($saved.PSObject.Properties | ForEach-Object { $_.Name })
                if (($pnAb -contains 'AgeBasis') -and -not [string]::IsNullOrWhiteSpace($saved.AgeBasis)) {
                    $tryDef = Resolve-AgeBasisFromString -Raw ([string]$saved.AgeBasis)
                    if ($null -ne $tryDef) {
                        $defAb = $tryDef
                    }
                }
            }
            Write-Host ''
            Write-Host 'Age basis (date compared to the cutoff):' -ForegroundColor DarkGray
            Write-Host '  A / LastWriteTime / Modified          = last modified time (default)' -ForegroundColor DarkGray
            Write-Host '  B / LastAccessTime / Opened           = last access (NTFS; can be wrong if last-access updates are disabled)' -ForegroundColor DarkGray
            Write-Host '  C / LatestWriteOrAccess / ModifiedOrOpened = newer of modified and last access' -ForegroundColor DarkGray
            Write-Host '  CreationTime, Earliest                = created, or older of created vs modified' -ForegroundColor DarkGray
            $r = Read-Host "AgeBasis [$defAb]"
            if (-not [string]::IsNullOrWhiteSpace($r)) {
                $nab = Normalize-AgeBasisParameter -abRaw $r
                if ($null -ne $nab) {
                    $ageBasisEffective = $nab
                }
            }
            else {
                $ageBasisEffective = $defAb
            }
        }
    }
}

$in = Normalize-UserPath -Path $in
$arch = Normalize-UserPath -Path $arch

$validationAttempt = 0
$maxValidationAttempts = 25

while ($true) {
    $validationAttempt++
    if ($validationAttempt -gt $maxValidationAttempts) {
        Write-Error "Configuration could not be validated after $maxValidationAttempts attempts. Fix paths (or create missing folders), then re-run. For a missing archive folder only, answer yes when asked to use -Commit, or create the folder first."
        exit 1
    }

    $validateDoCommit = $doCommit -and -not $allSharesFlag
    $issues = @(Test-SingleConfigValidity -InputPath $in -ArchivePath $arch -Years $yrs -DoCommit $validateDoCommit -AllSharesMode:$allSharesFlag)

    if ($issues.Count -eq 0) {
        break
    }

    $onlyArchiveMissing = (-not $validateDoCommit) -and ($issues.Count -eq 1) -and ($issues[0].Field -eq 'ArchivePath') -and
    ($issues[0].Message -match 'does not exist')

    if ($onlyArchiveMissing) {
        Write-Host ''
        if ($allSharesFlag) {
            Write-Host 'The archive folder does not exist yet. Create it manually first; -All only writes HTML reports (preview, no -Commit).' -ForegroundColor Yellow
        }
        else {
            Write-Host 'The archive folder does not exist yet. Preview mode requires that folder to exist (or use -Commit to create it).' -ForegroundColor Yellow
            $offerCommit = Read-Host 'Use -Commit for this run so the archive folder can be created? [y/N]'
            if ($offerCommit -match '^(y|yes)$') {
                $doCommit = $true
                continue
            }
        }
    }

    Write-Host ''
    Write-Host 'Configuration check: one or more values need correction (details below). Press Enter only if you already fixed the path or folder outside this window; otherwise type a new value.' -ForegroundColor Yellow
    foreach ($it in $issues) {
        Write-Host "  - [$($it.Field)] $($it.Message)" -ForegroundColor Yellow
    }
    Write-Host ''

    foreach ($it in $issues) {
        switch ($it.Field) {
            'InputPath' {
                $pickedIn = $null
                if (-not $skipShareMenuFlag) {
                    Write-Host 'Pick a published share, or choose Other for a manual path.' -ForegroundColor DarkGray
                    $pickedIn = Invoke-InputPathShareMenu
                }
                if (-not [string]::IsNullOrWhiteSpace($pickedIn)) {
                    $in = Normalize-UserPath -Path $pickedIn
                }
                else {
                    $r = Read-Host "InputPath [$in]"
                    if (-not [string]::IsNullOrWhiteSpace($r)) {
                        $in = Normalize-UserPath -Path $r
                    }
                }
            }
            'ArchivePath' {
                $fromArchiveShare = Get-LocalPathForShareName -ShareName $effectiveArchiveShareName
                if (-not [string]::IsNullOrWhiteSpace($fromArchiveShare)) {
                    Write-Host ("Use archive share '{0}' -> {1}? [Y/n]" -f $effectiveArchiveShareName, $fromArchiveShare)
                    $useAs = Read-Host
                    if ($useAs -notmatch '^(n|no)$') {
                        $arch = $fromArchiveShare
                    }
                    else {
                        $r = Read-Host "ArchivePath [$arch]"
                        if (-not [string]::IsNullOrWhiteSpace($r)) {
                            $arch = Normalize-UserPath -Path $r
                        }
                    }
                }
                else {
                    $r = Read-Host "ArchivePath [$arch]"
                    if (-not [string]::IsNullOrWhiteSpace($r)) {
                        $arch = Normalize-UserPath -Path $r
                    }
                }
            }
            'Years' {
                $r = Read-Host "Years [$yrs]"
                if (-not [string]::IsNullOrWhiteSpace($r)) {
                    $parsed = 0.0
                    if ([double]::TryParse($r, [ref]$parsed) -and $parsed -gt 0) {
                        $yrs = $parsed
                    }
                }
            }
        }
    }
}

$InputPath = $in
$ArchivePath = $arch
$Years = $yrs
$Output = $outPref
$normCliOut = Normalize-SavedOutputFormat -Raw $Output
if ($null -ne $normCliOut) {
    $Output = $normCliOut
}
$RemoveEmptyFolders = $removePref
$AgeBasis = $ageBasisEffective

if ($allSharesFlag) {
    $doCommit = $false
    $Output = 'HTML'
    $outPref = 'HTML'
}

# --- Main: one Invoke-SingleArchiveJob call, or -All loop (preview-only, HTML per share) ---

if ($allSharesFlag) {
    Write-Host ''
    Write-Host '========== ALL SHARES (preview only, one HTML report per share) ==========' -ForegroundColor Cyan
    Write-Host ('  Report folder (ArchivePath): {0}' -f $ArchivePath)
    Write-Host ('  Years: {0}  AgeBasis: {1}' -f $yrs, $ageBasisEffective)
    Write-Host '  -Commit, saved Commit, and prompts are ignored for moves; no files are moved.'
    Write-Host '==========================================================================' -ForegroundColor Cyan
    Write-Host ''

    try {
        if (-not (Test-Path -LiteralPath $ArchivePath)) {
            throw "ArchivePath does not exist. Create the folder first; -All writes HTML reports there (preview only)."
        }
        $archiveResolved = (Get-Item -LiteralPath $ArchivePath).FullName.TrimEnd('\')
    }
    catch {
        Write-Error $_
        exit 1
    }

    $published = @(Get-PublishedDiskShares)
    if ($published.Count -eq 0) {
        Write-Error 'No published disk shares found (disk shares only; Type 0, names without a trailing $).'
        exit 1
    }

    $outFormat = 'HTML'
    $htmlReportOutPath = $null
    $allHtmlPaths = New-Object System.Collections.Generic.List[string]
    $shareJobs = New-Object System.Collections.Generic.List[object]
    $cutoff = (Get-Date).AddYears(-$yrs)

    foreach ($sh in $published) {
        $shareRoot = $sh.LocalPath.TrimEnd('\')
        if (Test-ArchiveUnderInput -InputResolved $shareRoot -ArchiveResolved $archiveResolved) {
            Write-Host "Skipping share '$($sh.Name)' ($shareRoot): archive folder is this path or inside it." -ForegroundColor Yellow
            continue
        }
        try {
            $j = Invoke-SingleArchiveJob -InputResolved $shareRoot -ArchiveResolved $archiveResolved -YearsNum $yrs `
                -AgeBasisEff $ageBasisEffective -DoCommit $false -OutFormatIn 'HTML' -ShareNameLabel $sh.Name `
                -BoundRemoveEmptyFolders $false -RemoveEmptyFoldersParam '' -RemoveEmptyPref ''
            $null = $shareJobs.Add($j)
            if ($null -ne $j.HtmlReportPath -and $j.HtmlReportPath -ne '') {
                $null = $allHtmlPaths.Add($j.HtmlReportPath)
            }
        }
        catch {
            Write-Warning "Share '$($sh.Name)' ($shareRoot): $($_.Exception.Message)"
        }
    }

    Write-Host ''
    Write-Host '========== ALL-SHARES SUMMARY ==========' -ForegroundColor Cyan
    Write-Host ("  Share jobs completed: {0}" -f $shareJobs.Count)
    Write-Host ("  HTML reports written: {0}" -f $allHtmlPaths.Count)
    foreach ($hp in $allHtmlPaths) {
        Write-Host ("    {0}" -f $hp) -ForegroundColor DarkGray
    }
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host ''

    $InputPath = '(All published disk shares)'
    $inputResolved = $InputPath
    $doCommit = $false
    $removeEmptyChoice = $null
    $fileScanCount = 0
    $skippedTooNew = 0
    $plannedMovedCountAgg = 0
    $failedCountAgg = 0
    $reclaimBytesAgg = [long]0
    $results = @()
    $listedForArchiveAll = 0
    foreach ($sj in $shareJobs) {
        $fileScanCount += $sj.FileScanCount
        $skippedTooNew += $sj.SkippedTooNew
        $plannedMovedCountAgg += $sj.PlannedMovedCount
        $failedCountAgg += $sj.FailedCount
        $reclaimBytesAgg += $sj.ReclaimBytes
        $listedForArchiveAll += $sj.Results.Count
    }
    $reclaimDisplay = Format-DataSize -Bytes $reclaimBytesAgg
    $reclaimBytes = $reclaimBytesAgg
    $topLevelEntryCount = -1
    $depth1FileCount = -1
    $topLevelFolderCount = -1
    $gciErrCount = -1
    $firstGciErrorLog = ''
    $plannedMovedRows = @()
    $failedRows = @()
}
else {
    try {
        $inputResolved = Get-ResolvedPath -Path $InputPath
        $inputRootItem = Get-Item -LiteralPath $inputResolved -Force -ErrorAction Stop
        if (-not $inputRootItem.PSIsContainer) {
            throw "InputPath must be a folder (directory), not a file: $inputResolved"
        }
    }
    catch {
        Write-Error $_
        exit 1
    }

    try {
        if (-not (Test-Path -LiteralPath $ArchivePath)) {
            if ($doCommit) {
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

    $rbRm = $PSBoundParameters.ContainsKey('RemoveEmptyFolders')
    $rmVal = if ($rbRm) { [string]$RemoveEmptyFolders } else { '' }
    $rpStr = if ($null -ne $removePref) { [string]$removePref } else { '' }

    $job = Invoke-SingleArchiveJob -InputResolved $inputResolved -ArchiveResolved $archiveResolved -YearsNum $yrs `
        -AgeBasisEff $ageBasisEffective -DoCommit $doCommit -OutFormatIn "$Output" -ShareNameLabel '' `
        -BoundRemoveEmptyFolders $rbRm -RemoveEmptyFoldersParam $rmVal -RemoveEmptyPref $rpStr

    $inputResolved = $job.InputResolved
    $archiveResolved = $job.ArchiveResolved
    $results = $job.Results
    $fileScanCount = $job.FileScanCount
    $skippedTooNew = $job.SkippedTooNew
    $plannedMovedRows = @($results | Where-Object { $_.Status -in 'Planned', 'Moved' })
    $failedRows = @($results | Where-Object { $_.Status -eq 'Failed' })
    $reclaimBytes = $job.ReclaimBytes
    $reclaimDisplay = $job.ReclaimDisplay
    $outFormat = $job.OutFormat
    $htmlReportOutPath = $job.HtmlReportPath
    $removeEmptyChoice = $job.RemoveEmptyChoice
    $cutoff = $job.Cutoff
    $topLevelEntryCount = $job.TopLevelEntryCount
    $depth1FileCount = $job.Depth1FileCount
    $topLevelFolderCount = $job.TopLevelFolderCount
    $gciErrCount = $job.GciErrorCount
    $firstGciErrorLog = $job.FirstGciErrorLog
}

if (-not $noSaveConfigFlag) {
    try {
        $payload = [ordered]@{
            SchemaVersion      = 1
            ScriptVersion      = $script:Version
            SavedAt            = (Get-Date).ToString('o')
            InputPath          = $InputPath
            ArchivePath        = $ArchivePath
            ArchiveShareName   = $effectiveArchiveShareName
            Years              = $yrs
            AgeBasis           = $ageBasisEffective
            Commit             = [bool]$doCommit
            Output             = $outFormat
            RemoveEmptyFolders = $(if ($null -ne $removeEmptyChoice) { $removeEmptyChoice } else { $null })
            AllShares          = [bool]$allSharesFlag
        }
        $json = $payload | ConvertTo-Json -Depth 4
        Set-Content -LiteralPath $configFile -Value $json -Encoding UTF8 -Force
        Write-Host "Saved settings for next run: $configFile" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not save config file: $($_.Exception.Message)"
    }
}

if (-not $noRunLogFlag) {
    try {
        $runLogDir = Split-Path -Parent $configFile
        if ([string]::IsNullOrWhiteSpace($runLogDir)) {
            $runLogDir = (Get-Location).Path
        }
        $runLogPath = Join-Path $runLogDir 'Archive-OldFiles.run.log'
        $tab = [char]9
        $logBlock = New-Object System.Collections.Generic.List[string]
        $null = $logBlock.Add("----- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -----")
        $null = $logBlock.Add("Version=$script:Version")
        $null = $logBlock.Add("InputPath=$InputPath")
        $null = $logBlock.Add("ResolvedInput=$inputResolved")
        $null = $logBlock.Add("ArchivePath=$ArchivePath")
        $null = $logBlock.Add("ResolvedArchive=$archiveResolved")
        $null = $logBlock.Add("Years=$yrs AgeBasis=$ageBasisEffective Cutoff=$cutoff")
        $null = $logBlock.Add("Probe_TopLevelEntries=$topLevelEntryCount Probe_RootFiles=$depth1FileCount Probe_RootFolders=$topLevelFolderCount")
        $null = $logBlock.Add("GciErrorCount=$gciErrCount FirstGciError=$firstGciErrorLog")
        if ($allSharesFlag) {
            $null = $logBlock.Add('Mode=AllSharesPreview Commit=false (forced)')
            $pmcLog = $plannedMovedCountAgg
            $fmcLog = $failedCountAgg
            $listedLog = $listedForArchiveAll
        }
        else {
            $pmcLog = $plannedMovedRows.Count
            $fmcLog = $failedRows.Count
            $listedLog = $results.Count
        }
        $null = $logBlock.Add("FilesScanned=$fileScanCount SkippedNewerThanCutoff=$skippedTooNew ListedForArchive=$listedLog Output=$outFormat Commit=$doCommit")
        $null = $logBlock.Add("PlannedMovedCount=$pmcLog FailedCount=$fmcLog ReclaimBytes_PlannedMoved=$reclaimBytes ReclaimHuman=$reclaimDisplay")
        $null = $logBlock.Add("FILES_MEETING_AGE_RULE_TAB_SEPARATED_SourcePath_Bytes_Status_ComparedForAge_ISO")
        if ($allSharesFlag) {
            foreach ($sj in $shareJobs) {
                $null = $logBlock.Add("SHARE_BEGIN=$($sj.InputResolved)")
                foreach ($row in $sj.Results) {
                    $cmp = $row.ComparedForAge.ToString('o')
                    $null = $logBlock.Add(('{0}{1}{2}{1}{3}{1}{4}' -f $row.SourcePath, $tab, $row.Length, $row.Status, $cmp))
                }
                $null = $logBlock.Add('SHARE_END')
            }
        }
        else {
            foreach ($row in $results) {
                $cmp = $row.ComparedForAge.ToString('o')
                $null = $logBlock.Add(('{0}{1}{2}{1}{3}{1}{4}' -f $row.SourcePath, $tab, $row.Length, $row.Status, $cmp))
            }
        }
        $null = $logBlock.Add('END_FILES_MEETING_AGE_RULE')
        $null = $logBlock.Add('-----')
        Add-Content -LiteralPath $runLogPath -Value ($logBlock.ToArray() -join [Environment]::NewLine) -Encoding UTF8
        Write-Host "Run log (includes per-file list): $runLogPath" -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "Could not write run log: $($_.Exception.Message)"
    }
}

exit 0
