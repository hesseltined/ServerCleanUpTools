<#
Purpose: Mirror NTFS permissions from a source path to a destination path, with validation sampling and reporting.
Date Created: 2026-03-31
Date Modified: 2026-03-31
Author: Doug Hesseltine
Version: 1.0.0
Version Notes: Runtime metadata updates are automatically tracked in the JSON config (LastRunVersion/RunCount).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path -Path $PSScriptRoot -ChildPath "Sync-NTFSPermissions.config.json"),

    [Parameter(Mandatory = $false)]
    [int]$RandomSampleCount = 20,

    [Parameter(Mandatory = $false)]
    [int]$SequentialFolderCheckCount = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptVersion = "1.0.0"
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Now = Get-Date
$Timestamp = $Now.ToString("yyyyMMdd-HHmmss")

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[ OK ] $Message" -ForegroundColor Green
}

function Load-Config {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warn "Config exists but could not be parsed: $Path"
        return $null
    }
}

function Save-Config {
    param(
        [string]$Path,
        [string]$Source,
        [string]$Destination,
        [int]$RandomCount,
        [int]$SeqFolderCount,
        [int]$PriorRunCount
    )

    $configObject = [ordered]@{
        SourcePath                  = $Source
        DestinationPath             = $Destination
        RandomSampleCount           = $RandomCount
        SequentialFolderCheckCount  = $SeqFolderCount
        LastRunBy                   = $CurrentUser
        LastRunOn                   = (Get-Date).ToString("s")
        LastRunVersion              = $ScriptVersion
        RunCount                    = ($PriorRunCount + 1)
    }

    $configObject | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Resolve-InputValue {
    param(
        [string]$Label,
        [string]$ProvidedValue,
        [string]$SavedValue
    )

    if (-not [string]::IsNullOrWhiteSpace($ProvidedValue)) {
        return $ProvidedValue.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($SavedValue)) {
        $useSaved = Read-Host "$Label saved as '$SavedValue'. Press Enter to keep, or type a new path"
        if ([string]::IsNullOrWhiteSpace($useSaved)) {
            return $SavedValue.Trim()
        }
        return $useSaved.Trim()
    }

    while ($true) {
        $manualValue = Read-Host "Enter $Label"
        if (-not [string]::IsNullOrWhiteSpace($manualValue)) {
            return $manualValue.Trim()
        }
    }
}

function Get-RelativePath {
    param(
        [string]$BasePath,
        [string]$FullPath
    )

    if ($FullPath.TrimEnd('\') -eq $BasePath.TrimEnd('\')) {
        return "."
    }

    $baseUri = New-Object System.Uri(($BasePath.TrimEnd('\') + '\'))
    $fullUri = New-Object System.Uri($FullPath)
    $relativeUri = $baseUri.MakeRelativeUri($fullUri).ToString()
    $relativePath = [System.Uri]::UnescapeDataString($relativeUri).Replace('/', '\')
    return $relativePath
}

function Get-AccessSddl {
    param([string]$Path)
    $acl = Get-Acl -LiteralPath $Path
    return $acl.GetSecurityDescriptorSddlForm([System.Security.AccessControl.AccessControlSections]::Access)
}

function Normalize-PathForTraversal {
    param([string]$Path)

    $full = (Get-Item -LiteralPath $Path).FullName
    $root = [System.IO.Path]::GetPathRoot($full)

    # Keep trailing slash for filesystem roots (e.g., C:\, \\server\share\).
    if ($full -eq $root) {
        return $full
    }

    return $full.TrimEnd('\')
}

function Get-TreeItems {
    param([string]$Path)

    $rootItem = Get-Item -LiteralPath $Path -Force
    if (-not $rootItem.PSIsContainer) {
        return @($rootItem)
    }

    # Some trees may contain protected subfolders; continue and return what can be enumerated.
    $children = @(Get-ChildItem -LiteralPath ($rootItem.FullName + '\*') -Recurse -Force -ErrorAction SilentlyContinue)
    return @($rootItem) + $children
}

function Get-SafeCount {
    param([object]$InputObject)

    if ($null -eq $InputObject) {
        return 0
    }

    if ($InputObject -is [System.Collections.ICollection]) {
        return $InputObject.Count
    }

    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        return @($InputObject).Count
    }

    return 1
}

function Try-SetOwnerThenAcl {
    param(
        [string]$TargetPath,
        [System.Security.AccessControl.FileSystemSecurity]$SourceAcl,
        [string]$DomainUser
    )

    $changedOwner = $false
    $ownerReason = $null

    try {
        $targetAcl = Get-Acl -LiteralPath $TargetPath
        $ownerIdentity = New-Object System.Security.Principal.NTAccount($DomainUser)
        $targetAcl.SetOwner($ownerIdentity)
        Set-Acl -LiteralPath $TargetPath -AclObject $targetAcl
        $changedOwner = $true
        $ownerReason = "Ownership changed to $DomainUser to permit ACL update."
    }
    catch {
        return [pscustomobject]@{
            Success      = $false
            OwnerChanged = $false
            Reason       = "Ownership fallback failed: $($_.Exception.Message)"
        }
    }

    try {
        Set-Acl -LiteralPath $TargetPath -AclObject $SourceAcl
        return [pscustomobject]@{
            Success      = $true
            OwnerChanged = $changedOwner
            Reason       = $ownerReason
        }
    }
    catch {
        return [pscustomobject]@{
            Success      = $false
            OwnerChanged = $changedOwner
            Reason       = "Ownership changed, but ACL apply still failed: $($_.Exception.Message)"
        }
    }
}

function New-ItemResult {
    param(
        [string]$RelativePath,
        [string]$ItemType,
        [bool]$Success,
        [bool]$OwnerChanged,
        [string]$Message
    )

    return [pscustomobject]@{
        RelativePath = $RelativePath
        ItemType     = $ItemType
        Success      = $Success
        OwnerChanged = $OwnerChanged
        Message      = $Message
    }
}

Write-Info "Loading config (if present): $ConfigPath"
$config = Load-Config -Path $ConfigPath
$priorRunCount = 0
if ($null -ne $config -and $null -ne $config.RunCount) {
    $priorRunCount = [int]$config.RunCount
}

$savedSource = if ($null -ne $config) { [string]$config.SourcePath } else { $null }
$savedDestination = if ($null -ne $config) { [string]$config.DestinationPath } else { $null }
$savedRandom = if ($null -ne $config -and $null -ne $config.RandomSampleCount) { [int]$config.RandomSampleCount } else { $RandomSampleCount }
$savedSeq = if ($null -ne $config -and $null -ne $config.SequentialFolderCheckCount) { [int]$config.SequentialFolderCheckCount } else { $SequentialFolderCheckCount }

$SourcePath = Resolve-InputValue -Label "SourcePath" -ProvidedValue $SourcePath -SavedValue $savedSource
$DestinationPath = Resolve-InputValue -Label "DestinationPath" -ProvidedValue $DestinationPath -SavedValue $savedDestination

if ($RandomSampleCount -le 0) { $RandomSampleCount = $savedRandom }
if ($SequentialFolderCheckCount -le 0) { $SequentialFolderCheckCount = $savedSeq }
if ($RandomSampleCount -le 0) { $RandomSampleCount = 20 }
if ($SequentialFolderCheckCount -le 0) { $SequentialFolderCheckCount = 20 }

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "SourcePath does not exist: $SourcePath"
}
if (-not (Test-Path -LiteralPath $DestinationPath)) {
    throw "DestinationPath does not exist: $DestinationPath"
}

$sourceRoot = Normalize-PathForTraversal -Path $SourcePath
$destinationRoot = Normalize-PathForTraversal -Path $DestinationPath

Save-Config -Path $ConfigPath -Source $sourceRoot -Destination $destinationRoot -RandomCount $RandomSampleCount -SeqFolderCount $SequentialFolderCheckCount -PriorRunCount $priorRunCount
Write-Ok "Config saved: $ConfigPath"

Write-Info "Collecting source and destination trees..."
$sourceItems = @(Get-TreeItems -Path $sourceRoot)
$targetItemsByRelative = @{}

$destinationItems = @(Get-TreeItems -Path $destinationRoot)
foreach ($destItem in $destinationItems) {
    $relative = Get-RelativePath -BasePath $destinationRoot -FullPath $destItem.FullName
    $targetItemsByRelative[$relative.ToLowerInvariant()] = $destItem
}

$results = New-Object System.Collections.Generic.List[object]
$missingTargets = New-Object System.Collections.Generic.List[object]
$ownershipChanges = New-Object System.Collections.Generic.List[object]

Write-Info "Applying ACLs from source to destination..."
foreach ($sourceItem in $sourceItems) {
    $relativePath = Get-RelativePath -BasePath $sourceRoot -FullPath $sourceItem.FullName
    $key = $relativePath.ToLowerInvariant()

    if (-not $targetItemsByRelative.ContainsKey($key)) {
        $missingTargets.Add([pscustomobject]@{
            RelativePath = $relativePath
            ItemType     = if ($sourceItem.PSIsContainer) { "Directory" } else { "File" }
            Message      = "Target path not found under destination."
        })
        continue
    }

    $targetItem = $targetItemsByRelative[$key]
    $itemType = if ($sourceItem.PSIsContainer) { "Directory" } else { "File" }

    try {
        $sourceAcl = Get-Acl -LiteralPath $sourceItem.FullName
        Set-Acl -LiteralPath $targetItem.FullName -AclObject $sourceAcl
        $results.Add((New-ItemResult -RelativePath $relativePath -ItemType $itemType -Success $true -OwnerChanged $false -Message "ACL applied successfully."))
    }
    catch {
        $firstError = $_.Exception.Message
        $fallback = Try-SetOwnerThenAcl -TargetPath $targetItem.FullName -SourceAcl $sourceAcl -DomainUser $CurrentUser

        if ($fallback.Success) {
            if ($fallback.OwnerChanged) {
                $ownershipChanges.Add([pscustomobject]@{
                    RelativePath = $relativePath
                    ItemType     = $itemType
                    Message      = $fallback.Reason
                })
            }

            $results.Add((New-ItemResult -RelativePath $relativePath -ItemType $itemType -Success $true -OwnerChanged $fallback.OwnerChanged -Message "Initial ACL apply failed: $firstError; fallback succeeded.")) 
        }
        else {
            $results.Add((New-ItemResult -RelativePath $relativePath -ItemType $itemType -Success $false -OwnerChanged $fallback.OwnerChanged -Message "Initial ACL apply failed: $firstError; fallback failed: $($fallback.Reason)"))
        }
    }
}

Write-Info "Running validation checks..."
$sourceFolders = $sourceItems | Where-Object { $_.PSIsContainer }
$sourceFiles = $sourceItems | Where-Object { -not $_.PSIsContainer }

$sequentialFolderChecks = $sourceFolders | Select-Object -First $SequentialFolderCheckCount
$samplePool = @()
$samplePool += $sourceFolders
$samplePool += $sourceFiles

$randomChecks = @()
if ((Get-SafeCount -InputObject $samplePool) -gt 0) {
    $randomChecks = $samplePool | Get-Random -Count ([Math]::Min($RandomSampleCount, (Get-SafeCount -InputObject $samplePool)))
}

$validationTargets = @()
$validationTargets += $sequentialFolderChecks
$validationTargets += $randomChecks | Where-Object { $sequentialFolderChecks.FullName -notcontains $_.FullName }

$validationResults = New-Object System.Collections.Generic.List[object]
foreach ($sourceItem in $validationTargets) {
    $relativePath = Get-RelativePath -BasePath $sourceRoot -FullPath $sourceItem.FullName
    $destPath = if ($relativePath -eq ".") { $destinationRoot } else { Join-Path -Path $destinationRoot -ChildPath $relativePath }

    if (-not (Test-Path -LiteralPath $destPath)) {
        $validationResults.Add([pscustomobject]@{
            RelativePath = $relativePath
            ItemType     = if ($sourceItem.PSIsContainer) { "Directory" } else { "File" }
            Match        = $false
            Reason       = "Destination path missing."
        })
        continue
    }

    try {
        $sourceSddl = Get-AccessSddl -Path $sourceItem.FullName
        $destSddl = Get-AccessSddl -Path $destPath
        $isMatch = ($sourceSddl -eq $destSddl)
        $validationResults.Add([pscustomobject]@{
            RelativePath = $relativePath
            ItemType     = if ($sourceItem.PSIsContainer) { "Directory" } else { "File" }
            Match        = $isMatch
            Reason       = if ($isMatch) { "DACL matches." } else { "DACL mismatch." }
        })
    }
    catch {
        $validationResults.Add([pscustomobject]@{
            RelativePath = $relativePath
            ItemType     = if ($sourceItem.PSIsContainer) { "Directory" } else { "File" }
            Match        = $false
            Reason       = "Validation error: $($_.Exception.Message)"
        })
    }
}

$applyFailures = @($results | Where-Object { -not $_.Success })
$validationFailures = @($validationResults | Where-Object { -not $_.Match })

$reportObject = [ordered]@{
    Script                 = "Sync-NTFSPermissions.ps1"
    ScriptVersion          = $ScriptVersion
    RunAt                  = $Now.ToString("s")
    RunBy                  = $CurrentUser
    SourcePath             = $sourceRoot
    DestinationPath        = $destinationRoot
    Summary                = [ordered]@{
        SourceItemsSeen         = (Get-SafeCount -InputObject $sourceItems)
        DestinationItemsSeen    = (Get-SafeCount -InputObject $destinationItems)
        ACLApplySuccessCount    = (@($results | Where-Object { $_.Success })).Count
        ACLApplyFailureCount    = (Get-SafeCount -InputObject $applyFailures)
        MissingTargetCount      = (Get-SafeCount -InputObject $missingTargets)
        OwnershipChangeCount    = (Get-SafeCount -InputObject $ownershipChanges)
        ValidationSampleCount   = (Get-SafeCount -InputObject $validationResults)
        ValidationFailureCount  = (Get-SafeCount -InputObject $validationFailures)
        OverallSuccess          = (((Get-SafeCount -InputObject $applyFailures) -eq 0) -and ((Get-SafeCount -InputObject $missingTargets) -eq 0) -and ((Get-SafeCount -InputObject $validationFailures) -eq 0))
    }
    RandomSampleTestedPaths = @($validationResults | Select-Object -ExpandProperty RelativePath)
    Problems               = [ordered]@{
        MissingTargets      = $missingTargets
        ACLApplyFailures    = $applyFailures
        ValidationFailures  = $validationFailures
    }
    OwnershipChanges       = $ownershipChanges
}

$reportFileJson = Join-Path -Path $PSScriptRoot -ChildPath ("Sync-NTFSPermissions.report.$Timestamp.json")
$reportFileTxt = Join-Path -Path $PSScriptRoot -ChildPath ("Sync-NTFSPermissions.report.$Timestamp.txt")

$reportObject | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportFileJson -Encoding UTF8

$textLines = New-Object System.Collections.Generic.List[string]
$textLines.Add("Sync-NTFSPermissions Report")
$textLines.Add("RunAt: $($reportObject.RunAt)")
$textLines.Add("RunBy: $($reportObject.RunBy)")
$textLines.Add("SourcePath: $sourceRoot")
$textLines.Add("DestinationPath: $destinationRoot")
$textLines.Add("")
$textLines.Add("Summary")
$textLines.Add("-------")
foreach ($entry in $reportObject.Summary.GetEnumerator()) {
    $textLines.Add(("{0}: {1}" -f $entry.Key, $entry.Value))
}
$textLines.Add("")
$textLines.Add("Random/Sequential Sample Tested Paths")
$textLines.Add("-----------------------------------")
foreach ($p in $reportObject.RandomSampleTestedPaths) {
    $textLines.Add(" - $p")
}
$textLines.Add("")
$textLines.Add("Problems")
$textLines.Add("--------")
if (((Get-SafeCount -InputObject $applyFailures) -eq 0) -and ((Get-SafeCount -InputObject $missingTargets) -eq 0) -and ((Get-SafeCount -InputObject $validationFailures) -eq 0)) {
    $textLines.Add("No problems found.")
}
else {
    if ((Get-SafeCount -InputObject $missingTargets) -gt 0) {
        $textLines.Add("MissingTargets:")
        foreach ($m in $missingTargets) {
            $textLines.Add(" - [$($m.ItemType)] $($m.RelativePath): $($m.Message)")
        }
    }
    if ((Get-SafeCount -InputObject $applyFailures) -gt 0) {
        $textLines.Add("ACLApplyFailures:")
        foreach ($f in $applyFailures) {
            $textLines.Add(" - [$($f.ItemType)] $($f.RelativePath): $($f.Message)")
        }
    }
    if ((Get-SafeCount -InputObject $validationFailures) -gt 0) {
        $textLines.Add("ValidationFailures:")
        foreach ($v in $validationFailures) {
            $textLines.Add(" - [$($v.ItemType)] $($v.RelativePath): $($v.Reason)")
        }
    }
}
$textLines.Add("")
$textLines.Add("Ownership Changes (only when required)")
$textLines.Add("--------------------------------------")
if ((Get-SafeCount -InputObject $ownershipChanges) -eq 0) {
    $textLines.Add("No ownership changes were required.")
}
else {
    foreach ($o in $ownershipChanges) {
        $textLines.Add(" - [$($o.ItemType)] $($o.RelativePath): $($o.Message)")
    }
}

$textLines | Set-Content -LiteralPath $reportFileTxt -Encoding UTF8

if ($reportObject.Summary.OverallSuccess) {
    Write-Ok "Permissions sync completed successfully."
}
else {
    Write-Warn "Permissions sync completed with issues. Review report files."
}

Write-Host ""
Write-Host "Report JSON: $reportFileJson"
Write-Host "Report TXT : $reportFileTxt"
Write-Host "Config     : $ConfigPath"
