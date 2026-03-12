<#
.SYNOPSIS
    Hard drive usage + per-user profile sizing + large folder/file reporting (RMM-ready).
    Includes specific alerting for User Downloads folders > 5GB, accurate cloud-file sizing, 
    and read-only SharePoint sync detection.

.NOTES
    Name:       HDDUsageCheck_v5.4.ps1
    Author:     Stu Villanti
    Modified:   Refactored for cleaner RMM console output and strict JSON handling.
    Version:    5.4

.DESCRIPTION
    - Reports used/free space for all fixed drives.
    - Reports per-drive Recycle Bin usage.
    - Lists each user profile under C:\Users with total size (ignoring cloud-only files).
    - Checks User "Downloads" folder specifically against a threshold.
    - Identifies synced SharePoint locations via HKEY_USERS (Read-Only).
    - Console output pushes alerts to the top for quick technician review. Returns exit codes.

.EXIT CODES
    0 = OK            (no issues)
    1 = Low space     (drive free space < threshold)
    2 = Large items   (Downloads > 5GB OR other large folders/files found)
    3 = Both issues   (low space AND large items detected)
    4 = Script error  (unexpected failure)
#>

[CmdletBinding()]
param(
    [Parameter()] [int]$LowSpaceThreshold = 10,
    [Parameter()] [string]$UserRoot = 'C:\Users',
    [Parameter()] [int]$LargeThresholdGB = 5,
    [Parameter()] [int]$DownloadsThresholdGB = 5,
    [Parameter()] [object]$IncludeHidden = $true,
    [Parameter()] [object]$AsJson,
    [Parameter()] [object]$IncludeRecycleBin = $true,
    [Parameter()] [int]$RecycleBinThresholdGB = 5,
    [Parameter()] [object]$IncludeLargeFiles = $true
)

function Convert-ToBool {
    param([Parameter(ValueFromPipeline)][AllowNull()][object]$Value)
    process {
        if ($null -eq $Value) { return $false }
        if ($Value -is [bool]) { return $Value }
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return [bool]([int]$Value) }
        switch ("$Value".Trim().ToLowerInvariant()) {
            'true'  { return $true }
            '1'     { return $true }
            default { return $false }
        }
    }
}

$IncludeHidden      = Convert-ToBool $IncludeHidden
$AsJson             = Convert-ToBool $AsJson
$IncludeRecycleBin  = Convert-ToBool $IncludeRecycleBin
$IncludeLargeFiles  = Convert-ToBool $IncludeLargeFiles

# ---------------- STATE & RESULTS ARRAYS ----------------
$LowSpaceFound    = $false
$LargeItemsFound  = $false

$Alerts             = New-Object System.Collections.Generic.List[string]
$driveResults       = @()
$recycleBinResults  = @()
$profileResults     = @()
$largeItemsResults  = @()
$downloadsResults   = @()
$sharePointResults  = @()

# ---------------- HELPERS ----------------
function To-GB {
    param([long]$Bytes, [int]$Round=2)
    if ($Bytes -le 0) { return 0 }
    return [math]::Round($Bytes / 1GB, $Round)
}

function Get-FolderSizeBytes {
    param([Parameter(Mandatory)][string]$Path)
    try {
        return (Get-ChildItem -Path $Path -Recurse -File -Force:$IncludeHidden -ErrorAction SilentlyContinue |
                Where-Object { -not ($_.Attributes -match 'RecallOnDataAccess' -or $_.Attributes -match 'Offline') } |
                Measure-Object -Property Length -Sum).Sum
    } catch { return 0 }
}

function Get-TopLevelNameFromProfilePath {
    param([Parameter(Mandatory)][string]$ProfilePath, [Parameter(Mandatory)][string]$FileFullPath)
    try {
        $rel = $FileFullPath.Substring($ProfilePath.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { return '(root)' }
        $parts = $rel.Split('\')
        if ($parts.Count -ge 2) { return $parts[0] }
        return '(root)'
    } catch { return '(unknown)' }
}

function Get-SharePointSyncFolders {
    $syncPaths = @()
    try {
        $userSIDs = Get-ChildItem -Path "Registry::HKEY_USERS" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Name -match 'S-1-5-21-[\d\-]+$' }
        
        foreach ($sid in $userSIDs) {
            $regPath = "$($sid.PSPath)\Software\SyncEngines\Providers\OneDrive"
            if (Test-Path $regPath) {
                $syncKeys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
                foreach ($key in $syncKeys) {
                    $mountPoint = (Get-ItemProperty -Path $key.PSPath -Name "MountPoint" -ErrorAction SilentlyContinue).MountPoint
                    if ($mountPoint -and ($mountPoint -notmatch "\\OneDrive$") -and ($mountPoint -notmatch "\\OneDrive - ")) {
                        if ($syncPaths -notcontains $mountPoint) { $syncPaths += $mountPoint }
                    }
                }
            }
        }
    } catch { }
    return $syncPaths
}

function Get-ProfileStatsSinglePass {
    param([Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$ProfilePath, [Parameter(Mandatory)][int64]$ThresholdBytes)
    $total = [int64]0
    $bytesByTop = @{}
    $largeFiles = New-Object System.Collections.Generic.List[object]

    try {
        Get-ChildItem -Path $ProfilePath -Recurse -File -Force:$IncludeHidden -ErrorAction SilentlyContinue | ForEach-Object {
            $f = $_
            $isCloudOnly = ($f.Attributes -match 'RecallOnDataAccess') -or ($f.Attributes -match 'Offline')
            
            if (-not $isCloudOnly) {
                $len = [int64]$f.Length
                $total += $len

                $top = Get-TopLevelNameFromProfilePath -ProfilePath $ProfilePath -FileFullPath $f.FullName
                if (-not $bytesByTop.ContainsKey($top)) { $bytesByTop[$top] = [int64]0 }
                $bytesByTop[$top] += $len

                if ($IncludeLargeFiles -and ($len -ge $ThresholdBytes)) {
                    $largeFiles.Add([pscustomobject]@{ ItemType = 'File'; User = $User; Name = $f.Name; SizeGB = To-GB $len; Path = $f.FullName }) | Out-Null
                }
            }
        }
    } catch { }

    return [pscustomobject]@{ User = $User; Path = $ProfilePath; TotalBytes = $total; BytesByTopFolder = $bytesByTop; LargeFiles = $largeFiles.ToArray() }
}

# ==============================================================================
# DATA COLLECTION PHASE
# ==============================================================================

try {
    # --- DRIVE USAGE ---
    $drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    foreach ($d in $drives) {
        $totalGB = To-GB $d.Size
        $freeGB  = To-GB $d.FreeSpace
        $usedGB  = [math]::Round($totalGB - $freeGB, 2)
        $pctFree = if ($d.Size) { [math]::Round(($d.FreeSpace / $d.Size) * 100, 1) } else { 0 }

        $driveResults += [pscustomobject]@{ Drive = $d.DeviceID; TotalGB = $totalGB; UsedGB = $usedGB; FreeGB = $freeGB; PercentFree = $pctFree }

        if ($pctFree -lt $LowSpaceThreshold) {
            $LowSpaceFound = $true
            $Alerts.Add("Drive $($d.DeviceID) is low on space ($pctFree% free. Limit: $LowSpaceThreshold%)")
        }

        # --- RECYCLE BIN ---
        if ($IncludeRecycleBin) {
            $rbPath = Join-Path ($d.DeviceID + '\') '$Recycle.Bin'
            if (Test-Path $rbPath) {
                $rbBytes = Get-FolderSizeBytes -Path $rbPath
                $rbGB = To-GB $rbBytes
                if ($rbGB -ge $RecycleBinThresholdGB) {
                    $LargeItemsFound = $true
                    $Alerts.Add("Drive $($d.DeviceID) Recycle Bin is $rbGB GB (Limit: $RecycleBinThresholdGB GB)")
                }
                $recycleBinResults += [pscustomobject]@{ Drive = $d.DeviceID; RecycleBinGB = $rbGB; RecycleBinPath = $rbPath }
            }
        }
    }

    # --- SHAREPOINT SYNC ---
    $sharePointResults = Get-SharePointSyncFolders

    # --- USER PROFILES ---
    if (Test-Path $UserRoot) {
        $skip = @('Default','Default User','Public','All Users')
        $profiles = Get-ChildItem -Path $UserRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $skip -notcontains $_.Name }
        $generalThresholdBytes = [int64]($LargeThresholdGB * 1GB)
        $downloadsThresholdBytes = [int64]($DownloadsThresholdGB * 1GB)

        foreach ($p in $profiles) {
            $stats = Get-ProfileStatsSinglePass -User $p.Name -ProfilePath $p.FullName -ThresholdBytes $generalThresholdBytes
            $profileResults += [pscustomobject]@{ User = $p.Name; Path = $p.FullName; SizeGB = To-GB $stats.TotalBytes }

            # Check Downloads
            if ($stats.BytesByTopFolder.ContainsKey('Downloads')) {
                $dlGB = To-GB $stats.BytesByTopFolder['Downloads']
                if ($dlGB -ge $DownloadsThresholdGB) {
                    $LargeItemsFound = $true
                    $Alerts.Add("User '$($stats.User)' Downloads folder is $dlGB GB (Limit: $DownloadsThresholdGB GB)")
                    $downloadsResults += [pscustomobject]@{ User = $stats.User; SizeGB = $dlGB }
                }
            }

            # Check General Large Folders
            foreach ($k in $stats.BytesByTopFolder.Keys) {
                $b = [int64]$stats.BytesByTopFolder[$k]
                if ($b -ge $generalThresholdBytes -and $k -ne 'Downloads') {
                    $LargeItemsFound = $true
                    $largeItemsResults += [pscustomobject]@{ ItemType = 'Folder'; User = $stats.User; Name = $k; SizeGB = To-GB $b; Path = if ($k -eq '(root)') { $stats.Path } else { Join-Path $stats.Path $k } }
                }
            }

            # Add Large Files
            if ($stats.LargeFiles.Count -gt 0) {
                $LargeItemsFound = $true
                $largeItemsResults += $stats.LargeFiles
            }
        }
    }

} catch {
    if (-not $AsJson) { Write-Host "CRITICAL SCRIPT ERROR: $($_.Exception.Message)" -ForegroundColor Red }
    exit 4
}

# ==============================================================================
# OUTPUT PHASE
# ==============================================================================

if ($AsJson) {
    # Output strictly JSON on a single compressed line (ideal for RMM custom fields)
    [pscustomobject]@{
        Version              = '5.4'
        LowSpaceFound        = $LowSpaceFound
        LargeItemsFound      = $LargeItemsFound
        Alerts               = $Alerts
        Drives               = $driveResults
        RecycleBins          = $recycleBinResults
        DownloadsAlerts      = $downloadsResults
        LargeItems           = $largeItemsResults
        Profiles             = $profileResults
        SharePointPaths      = $sharePointResults
    } | ConvertTo-Json -Depth 6 -Compress | Write-Output
} 
else {
    # Standard Console Output formatted for easy reading
    if ($Alerts.Count -gt 0) {
        Write-Host "=== ACTION REQUIRED ===" -ForegroundColor Yellow
        foreach ($alert in $Alerts) { Write-Host "[WARN] $alert" -ForegroundColor Yellow }
        Write-Host ""
    } else {
        Write-Host "[ OK ] No storage thresholds exceeded.`n" -ForegroundColor Green
    }

    Write-Host "=== DRIVE STATUS ===" -ForegroundColor Cyan
    foreach ($d in $driveResults) {
        Write-Host ("{0,-5} | {1,6}% Free | {2,7} GB Used / {3,7} GB Total" -f $d.Drive, $d.PercentFree, $d.UsedGB, $d.TotalGB)
    }
    Write-Host ""

    if ($largeItemsResults.Count -gt 0) {
        Write-Host ("=== LARGE ITEMS FOUND (> {0}GB) ===" -f $LargeThresholdGB) -ForegroundColor Cyan
        $largeItemsResults | Sort-Object SizeGB -Descending | ForEach-Object {
            Write-Host ("{0,-15} | {1,-6} | {2,7} GB | {3}" -f $_.User, $_.ItemType, $_.SizeGB, $_.Name)
        }
        Write-Host ""
    }

    if ($sharePointResults.Count -gt 0) {
        Write-Host "=== ACTIVE SHAREPOINT SYNCS ===" -ForegroundColor Cyan
        foreach ($sp in $sharePointResults) { Write-Host $sp }
        Write-Host ""
    }
}

# ---------------- EXIT CODES ----------------
if ($LowSpaceFound -and $LargeItemsFound) { exit 3 }
elseif ($LowSpaceFound)                   { exit 1 }
elseif ($LargeItemsFound)                 { exit 2 }
else                                      { exit 0 }