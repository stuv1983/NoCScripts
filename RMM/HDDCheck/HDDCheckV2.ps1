<#
.SYNOPSIS
    Hard drive usage + per-user profile sizing + large folder/file reporting (RMM-ready).
    v6.0: fast .NET enumeration, cloud-only vs local sizing, per-user Recycle Bin,
    classified OneDrive/SharePoint sync roots with [SYNC] tagging.

.NOTES
    Name:       HDDUsageCheck_v6.2.ps1
    Author:     Stu Villanti
    Modified:   v6.2 - Large-item console output shows full paths; large top-level
                folders now include a drill-down of their biggest second-level
                subfolders (top 5, >= 1 GB) collected in the same single pass.
                v6.1 - Extended-length ("\\?\" prefix) path support fixes >260-char
                failures (CloudStore tile caches etc.) and silent under-counting;
                SkippedDirs counter per profile surfaces unreadable directories.
                v6.0 - Stack-based .NET file enumeration (3-10x faster than GCI -Recurse),
                bitwise cloud-attribute checks, junction/reparse-point skipping (fixes
                potential loops and double counting), per-profile Local/CloudOnly/Synced
                byte breakdown, per-user Recycle Bin sizing with SID->name translation,
                sync roots classified as OneDrive vs SharePoint/Teams and mapped to
                owning user, [SYNC] tagging on large folders under a sync root.
    Version:    6.2
    Requires:   Windows PowerShell 5.1+ (also runs on PowerShell 7).

.DESCRIPTION
    - Reports used/free space for all fixed drives.
    - Reports Recycle Bin usage per drive AND per user (SID folders translated to names).
    - Lists each user profile under C:\Users with Local, Cloud-only (dehydrated) and
      Synced (under a OneDrive/SharePoint root) sizes.
    - Checks User "Downloads" folder against its own threshold (local + cloud split).
    - Identifies and classifies synced OneDrive/SharePoint locations via HKEY_USERS
      (read-only) and maps each root to its owning user.
    - Large folders that sit under a sync root are tagged [SYNC] so technicians can
      distinguish cloud-backed data (candidate for dehydration) from local-only data.
    - Console output pushes alerts to the top for quick technician review.
    - Returns structured exit codes for N-able task condition matching.

.EXIT CODES
    0 = OK            (no issues)
    1 = Low space     (drive free space below threshold)
    2 = Large items   (Downloads > threshold OR other large folders/files found)
    3 = Both issues   (low space AND large items detected)
    4 = Script error  (unexpected failure during data collection)

.PARAMETER LowSpaceThreshold
    Percentage of free space below which a drive is considered low. Default: 10

.PARAMETER UserRoot
    Root path containing user profile folders. Default: C:\Users

.PARAMETER LargeThresholdGB
    Size in GB above which a top-level profile folder or file triggers a large-item alert.
    NOTE: The Downloads folder is excluded from this check - it has its own dedicated
    threshold ($DownloadsThresholdGB) so the two can be tuned independently without
    double-alerting. Default: 5

.PARAMETER DownloadsThresholdGB
    Size in GB (local bytes only) above which a user's Downloads folder alerts. Default: 5

.PARAMETER IncludeHidden
    Include hidden files/folders in sizing. [object] type to accept N-able string inputs
    ("true"/"false"/"1"/"0") as well as native booleans; normalised via Convert-ToBool.
    Default: true

.PARAMETER AsJson
    When set, outputs a single compressed JSON line (ideal for N-able custom field
    ingestion). All Write-Host colour output is suppressed. [object] type - same
    N-able string compatibility reason as IncludeHidden. Default: false

.PARAMETER IncludeRecycleBin
    Whether to scan and report Recycle Bin sizes (per drive and per user).
    [object] type - same N-able string compatibility reason as IncludeHidden. Default: true

.PARAMETER RecycleBinThresholdGB
    Size in GB above which a drive's Recycle Bin triggers a large-item alert. Default: 5

.PARAMETER IncludeLargeFiles
    Whether to report individual large files within profiles.
    [object] type - same N-able string compatibility reason as IncludeHidden. Default: true
#>

[CmdletBinding()]
param(
    [Parameter()] [int]$LowSpaceThreshold      = 10,
    [Parameter()] [string]$UserRoot            = 'C:\Users',
    [Parameter()] [int]$LargeThresholdGB       = 5,
    [Parameter()] [int]$DownloadsThresholdGB   = 5,
    # [object] used (not [bool]) so N-able can pass "true"/"false" strings via task params.
    [Parameter()] [object]$IncludeHidden       = $true,
    [Parameter()] [object]$AsJson,
    [Parameter()] [object]$IncludeRecycleBin   = $true,
    [Parameter()] [int]$RecycleBinThresholdGB  = 5,
    [Parameter()] [object]$IncludeLargeFiles   = $true
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

# ---------------- ATTRIBUTE MASKS (bitwise - far faster than regex -match per file) ----
# FILE_ATTRIBUTE_OFFLINE                = 0x00001000
# FILE_ATTRIBUTE_RECALL_ON_OPEN         = 0x00040000
# FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS  = 0x00400000  (OneDrive Files-On-Demand placeholder)
$CLOUD_MASK   = 0x00441000
$HIDDEN_FLAG  = 0x00000002   # FILE_ATTRIBUTE_HIDDEN
$REPARSE_FLAG = 0x00000400   # FILE_ATTRIBUTE_REPARSE_POINT (junctions/symlinks - skip to
                             # avoid loops and double counting, e.g. legacy 'Application Data')

# ---------------- LONG PATH SUPPORT ----------------
# Windows paths > 260 chars (MAX_PATH) make DirectoryInfo throw "Could not find a part
# of the path" - e.g. deeply nested CloudStore tile-cache folders. Prefixing the scan
# root with '\\?\' switches the Win32 layer to extended-length mode (32k chars) and
# also tolerates names with trailing spaces/dots. All descendants enumerated from a
# prefixed root inherit the prefix, so applying it once at the root covers the tree.
# Requires .NET Framework 4.6.2+ (standard on Win10/11).
function Add-LongPathPrefix {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.StartsWith('\\'))   { return '\\?\UNC\' + $Path.Substring(2) }  # UNC form
    return '\\?\' + $Path
}
function Remove-LongPathPrefix {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path.StartsWith('\\?\UNC\')) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))     { return $Path.Substring(4) }
    return $Path
}

# ---------------- STATE & RESULTS COLLECTIONS ----------------
$LowSpaceFound   = $false
$LargeItemsFound = $false

# Generic List[T] throughout to avoid the O(n) array copy that += causes on @() arrays.
$Alerts             = New-Object System.Collections.Generic.List[string]
$driveResults       = New-Object System.Collections.Generic.List[object]
$recycleBinResults  = New-Object System.Collections.Generic.List[object]
$recycleBinPerUser  = New-Object System.Collections.Generic.List[object]
$profileResults     = New-Object System.Collections.Generic.List[object]
$largeItemsResults  = New-Object System.Collections.Generic.List[object]
$downloadsResults   = New-Object System.Collections.Generic.List[object]
$sharePointResults  = New-Object System.Collections.Generic.List[object]

# ---------------- HELPERS ----------------
function To-GB {
    param([long]$Bytes, [int]$Round = 2)
    if ($Bytes -le 0) { return 0 }
    return [math]::Round($Bytes / 1GB, $Round)
}

# SID -> friendly name, cached. Falls back to ProfileList registry (handles orphaned
# SIDs whose accounts no longer resolve), then to the raw SID string.
$script:SidNameCache = @{}
function Resolve-SidName {
    param([Parameter(Mandatory)][string]$Sid)
    if ($script:SidNameCache.ContainsKey($Sid)) { return $script:SidNameCache[$Sid] }
    $name = $null
    try {
        $name = ([System.Security.Principal.SecurityIdentifier]$Sid).
                    Translate([System.Security.Principal.NTAccount]).Value
    } catch { }
    if (-not $name) {
        try {
            $pl = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid" -ErrorAction Stop
            if ($pl.ProfileImagePath) { $name = Split-Path $pl.ProfileImagePath -Leaf }
        } catch { }
    }
    if (-not $name) { $name = $Sid }
    $script:SidNameCache[$Sid] = $name
    return $name
}

# ---------------- FAST FOLDER SIZING (stack-based .NET enumeration) ----------------
# Replaces Get-ChildItem -Recurse: DirectoryInfo.EnumerateFiles/EnumerateDirectories is
# typically 3-10x faster and lets us skip reparse points and inherit per-directory state
# (top-level folder name, sync membership) instead of recomputing it per file.
function Measure-FolderFast {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$SplitCloud   # when set, returns Local/CloudOnly split; otherwise Local only
    )
    $local = [int64]0; $cloud = [int64]0; $skipped = 0
    $stack = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
    try { $stack.Push([System.IO.DirectoryInfo]::new((Add-LongPathPrefix $Path))) } catch { return [pscustomobject]@{ LocalBytes = 0; CloudOnlyBytes = 0; SkippedDirs = 1 } }

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            foreach ($f in $dir.EnumerateFiles()) {
                $attrs = [int]$f.Attributes
                if (-not $IncludeHidden -and ($attrs -band $HIDDEN_FLAG)) { continue }
                if ($attrs -band $CLOUD_MASK) { if ($SplitCloud) { $cloud += $f.Length }; continue }
                $local += $f.Length
            }
            foreach ($sub in $dir.EnumerateDirectories()) {
                $attrs = [int]$sub.Attributes
                if ($attrs -band $REPARSE_FLAG) { continue }
                if (-not $IncludeHidden -and ($attrs -band $HIDDEN_FLAG)) { continue }
                $stack.Push($sub)
            }
        } catch {
            # Inaccessible directory (permissions/locks) - skip and keep going.
            $skipped++
            Write-Verbose "Measure-FolderFast: skipped '$(Remove-LongPathPrefix $dir.FullName)': $($_.Exception.Message)"
        }
    }
    return [pscustomobject]@{ LocalBytes = $local; CloudOnlyBytes = $cloud; SkippedDirs = $skipped }
}

# ---------------- SYNC ROOT DISCOVERY & CLASSIFICATION ----------------
# Reads HKEY_USERS\<SID>\Software\SyncEngines\Providers\* (read-only) and classifies
# each mount point as OneDrive (personal/business) or SharePoint/Teams library, mapped
# to the owning user via SID translation.
function Get-SyncRoots {
    $roots = New-Object System.Collections.Generic.List[object]
    try {
        $userSIDs = Get-ChildItem -Path "Registry::HKEY_USERS" -ErrorAction SilentlyContinue |
                    Where-Object { $_.PSChildName -match '^S-1-5-21-[\d\-]+$' }

        foreach ($sidKey in $userSIDs) {
            $sid      = $sidKey.PSChildName
            $userName = Resolve-SidName $sid
            $provRoot = "$($sidKey.PSPath)\Software\SyncEngines\Providers"
            if (-not (Test-Path $provRoot)) { continue }

            foreach ($provider in (Get-ChildItem -Path $provRoot -ErrorAction SilentlyContinue)) {
                foreach ($key in (Get-ChildItem -Path $provider.PSPath -ErrorAction SilentlyContinue)) {
                    $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                    $mount = $props.MountPoint
                    if (-not $mount) { continue }
                    if ($roots | Where-Object { $_.Path -eq $mount }) { continue }  # dedupe

                    # Classification:
                    #  - '...\OneDrive' or '...\OneDrive - <Org>' mounts = OneDrive
                    #  - UrlNamespace containing '/personal/' (my.sharepoint.com) = OneDrive
                    #  - everything else (team sites, Teams channel libraries)   = SharePoint
                    $url  = "$($props.UrlNamespace)"
                    $type = if ($mount -match '\\OneDrive( - .+)?$' -or $url -match '(?i)-my\.sharepoint\.com|/personal/') {
                                'OneDrive'
                            } else {
                                'SharePoint/Teams'
                            }

                    $roots.Add([pscustomobject]@{
                        Path = $mount
                        Type = $type
                        User = $userName
                        Url  = $url
                    })
                }
            }
        }
    } catch {
        # Non-fatal: registry access may fail on locked-down endpoints.
        Write-Verbose "Get-SyncRoots failed: $($_.Exception.Message)"
    }
    return $roots
}

# ---------------- SINGLE-PASS PROFILE SCAN ----------------
# One traversal per profile collects: total local + cloud-only bytes, per-top-level-folder
# local bytes, per-top-level cloud bytes, synced bytes (files under a sync root), and
# large files. Top-level name and sync membership are inherited per-directory when pushed
# onto the stack - no per-file string maths.
function Get-ProfileStatsSinglePass {
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][string]$ProfilePath,
        [Parameter(Mandatory)][int64]$ThresholdBytes,
        [Parameter()][object[]]$SyncRoots = @()
    )
    $localTotal = [int64]0
    $cloudTotal = [int64]0
    $syncTotal  = [int64]0
    $skipped    = 0
    $localByTop = @{}
    $localBySub = @{}   # "Top\Sub" -> bytes; enables drill-down inside large top folders
    $cloudByTop = @{}
    $syncedTops = @{}   # top-level folder name -> $true if any part of it is under a sync root
    $largeFiles = New-Object System.Collections.Generic.List[object]

    # Pre-normalise sync roots (trailing slash, lowercase) for fast prefix checks.
    $normRoots = @($SyncRoots | ForEach-Object { ($_.Path.TrimEnd('\') + '\').ToLowerInvariant() })

    function Test-UnderSyncRoot([string]$fullPath) {
        # FullName carries the \\?\ prefix (long-path mode) - strip before comparing.
        $p = ((Remove-LongPathPrefix $fullPath).TrimEnd('\') + '\').ToLowerInvariant()
        foreach ($r in $normRoots) { if ($p.StartsWith($r)) { return $true } }
        return $false
    }

    # Stack entries: DirectoryInfo, its top-level folder name, and inherited sync flag.
    $stack = New-Object System.Collections.Generic.Stack[object]
    try {
        # \\?\ prefix = extended-length path mode; survives >260-char paths and odd
        # names (trailing spaces/dots) that otherwise throw during re-enumeration.
        $rootDir = [System.IO.DirectoryInfo]::new((Add-LongPathPrefix $ProfilePath))

        # Seed: files directly in the profile root count as '(root)'.
        # Sub = "Top\SecondLevel" key inherited from depth 2 downward; IsTop marks a
        # top-level directory itself (its child dirs define the Sub keys).
        $stack.Push([pscustomobject]@{ Dir = $rootDir; Top = '(root)'; Sub = $null; Synced = $false; IsRoot = $true; IsTop = $false })

        while ($stack.Count -gt 0) {
            $entry = $stack.Pop()
            $dir   = $entry.Dir
            try {
                foreach ($f in $dir.EnumerateFiles()) {
                    $attrs = [int]$f.Attributes
                    if (-not $IncludeHidden -and ($attrs -band $HIDDEN_FLAG)) { continue }

                    $top = $entry.Top
                    if ($attrs -band $CLOUD_MASK) {
                        # Dehydrated placeholder: contributes no local disk usage but is
                        # tracked so reports show how much data lives cloud-only.
                        $cloudTotal += $f.Length
                        if (-not $cloudByTop.ContainsKey($top)) { $cloudByTop[$top] = [int64]0 }
                        $cloudByTop[$top] += $f.Length
                        continue
                    }

                    $len = [int64]$f.Length
                    $localTotal += $len
                    if (-not $localByTop.ContainsKey($top)) { $localByTop[$top] = [int64]0 }
                    $localByTop[$top] += $len
                    # Second-level attribution for drill-down. Files sitting directly in
                    # a top-level folder are grouped under "(loose files)".
                    $subKey = if ($entry.Sub) { $entry.Sub } elseif ($entry.IsTop) { "$top\(loose files)" } else { $null }
                    if ($subKey) {
                        if (-not $localBySub.ContainsKey($subKey)) { $localBySub[$subKey] = [int64]0 }
                        $localBySub[$subKey] += $len
                    }
                    if ($entry.Synced) { $syncTotal += $len }

                    if ($IncludeLargeFiles -and ($len -ge $ThresholdBytes)) {
                        $largeFiles.Add([pscustomobject]@{
                            ItemType = 'File'; User = $User; Name = $f.Name
                            SizeGB   = To-GB $len; Path = Remove-LongPathPrefix $f.FullName
                            Synced   = [bool]$entry.Synced
                        })
                    }
                }

                foreach ($sub in $dir.EnumerateDirectories()) {
                    $attrs = [int]$sub.Attributes
                    if ($attrs -band $REPARSE_FLAG) { continue }   # skip junctions/symlinks
                    if (-not $IncludeHidden -and ($attrs -band $HIDDEN_FLAG)) { continue }

                    # Children of the profile root define the top-level folder name;
                    # children of a top-level dir define the "Top\Sub" drill-down key;
                    # deeper directories inherit both. Sync membership is inherited too,
                    # only re-tested while still false.
                    $subTop = if ($entry.IsRoot) { $sub.Name } else { $entry.Top }
                    $subKey = if ($entry.IsRoot)  { $null }
                              elseif ($entry.IsTop) { "$($entry.Top)\$($sub.Name)" }
                              else                  { $entry.Sub }
                    $subSynced = $entry.Synced
                    if (-not $subSynced -and $normRoots.Count -gt 0) {
                        $subSynced = Test-UnderSyncRoot $sub.FullName
                    }
                    if ($subSynced) { $syncedTops[$subTop] = $true }

                    $stack.Push([pscustomobject]@{ Dir = $sub; Top = $subTop; Sub = $subKey; Synced = $subSynced; IsRoot = $false; IsTop = [bool]$entry.IsRoot })
                }
            } catch {
                # Inaccessible directory (locked, permission-denied). Partial data is
                # better than none; SkippedDirs surfaces how much was missed.
                $skipped++
                Write-Verbose "Profile scan skipped '$(Remove-LongPathPrefix $dir.FullName)': $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Verbose "Get-ProfileStatsSinglePass partial failure for user '$User': $($_.Exception.Message)"
    }

    return [pscustomobject]@{
        User             = $User
        Path             = $ProfilePath
        LocalBytes       = $localTotal
        CloudOnlyBytes   = $cloudTotal
        SyncedBytes      = $syncTotal
        SkippedDirs      = $skipped
        LocalByTopFolder = $localByTop
        LocalBySubFolder = $localBySub
        CloudByTopFolder = $cloudByTop
        SyncedTopFolders = $syncedTops
        LargeFiles       = $largeFiles.ToArray()
    }
}

# ==============================================================================
# DATA COLLECTION PHASE
# ==============================================================================

try {
    # --- SHAREPOINT / ONEDRIVE SYNC ROOTS (needed before profile scans for tagging) ---
    $syncRoots = @(Get-SyncRoots)
    foreach ($sr in $syncRoots) { $sharePointResults.Add($sr) }

    # --- DRIVE USAGE ---
    $drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    foreach ($d in $drives) {
        $totalGB = To-GB $d.Size
        $freeGB  = To-GB $d.FreeSpace
        $usedGB  = [math]::Round($totalGB - $freeGB, 2)
        $pctFree = if ($d.Size) { [math]::Round(($d.FreeSpace / $d.Size) * 100, 1) } else { 0 }

        $driveResults.Add([pscustomobject]@{
            Drive       = $d.DeviceID; TotalGB = $totalGB
            UsedGB      = $usedGB;     FreeGB  = $freeGB
            PercentFree = $pctFree
        })

        if ($pctFree -lt $LowSpaceThreshold) {
            $LowSpaceFound = $true
            $Alerts.Add("Drive $($d.DeviceID) is low on space ($pctFree% free. Limit: $LowSpaceThreshold%)")
        }

        # --- RECYCLE BIN (per drive + per user SID folder) ---
        if ($IncludeRecycleBin) {
            # Join-Path "C:" '$Recycle.Bin' produces "C:$Recycle.Bin" (missing slash);
            # appending '\' to DeviceID first is the standard workaround.
            $rbPath = Join-Path ($d.DeviceID + '\') '$Recycle.Bin'
            if (Test-Path $rbPath) {
                $rbTotal = [int64]0
                # Each subfolder of $Recycle.Bin is one user's bin, named by SID.
                $sidDirs = Get-ChildItem -Path $rbPath -Directory -Force -ErrorAction SilentlyContinue
                foreach ($sd in $sidDirs) {
                    $m = Measure-FolderFast -Path $sd.FullName
                    if ($m.LocalBytes -le 0) { continue }
                    $rbTotal += $m.LocalBytes
                    $recycleBinPerUser.Add([pscustomobject]@{
                        Drive  = $d.DeviceID
                        User   = Resolve-SidName $sd.Name
                        Sid    = $sd.Name
                        SizeGB = To-GB $m.LocalBytes
                    })
                }
                $rbGB = To-GB $rbTotal
                if ($rbGB -ge $RecycleBinThresholdGB) {
                    $LargeItemsFound = $true
                    $Alerts.Add("Drive $($d.DeviceID) Recycle Bin is $rbGB GB (Limit: $RecycleBinThresholdGB GB)")
                }
                $recycleBinResults.Add([pscustomobject]@{
                    Drive          = $d.DeviceID
                    RecycleBinGB   = $rbGB
                    RecycleBinPath = $rbPath
                })
            }
        }
    }

    # --- USER PROFILES ---
    if (Test-Path $UserRoot) {
        # Skip list covers system/built-in profiles that should never be scanned.
        $skip = @('Default', 'Default User', 'Public', 'All Users', 'WDAGUtilityAccount', 'DefaultAppPool')
        $profiles = Get-ChildItem -Path $UserRoot -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $skip -notcontains $_.Name }

        $generalThresholdBytes = [int64]($LargeThresholdGB * 1GB)

        foreach ($p in $profiles) {
            # Only pass sync roots that live under this profile - keeps prefix checks tiny.
            $profileSyncRoots = @($syncRoots | Where-Object {
                $_.Path -and $_.Path.StartsWith($p.FullName, [System.StringComparison]::OrdinalIgnoreCase)
            })

            $stats = Get-ProfileStatsSinglePass -User $p.Name -ProfilePath $p.FullName `
                        -ThresholdBytes $generalThresholdBytes -SyncRoots $profileSyncRoots

            $profileResults.Add([pscustomobject]@{
                User        = $p.Name
                Path        = $p.FullName
                LocalGB     = To-GB $stats.LocalBytes     # actual bytes on disk
                CloudOnlyGB = To-GB $stats.CloudOnlyBytes # dehydrated placeholders (no disk cost)
                SyncedGB    = To-GB $stats.SyncedBytes    # local bytes under a OneDrive/SP root
                SkippedDirs = $stats.SkippedDirs          # dirs that could not be read (permissions)
                SizeGB      = To-GB $stats.LocalBytes     # kept for backwards JSON compatibility
            })

            # --- DOWNLOADS CHECK ---
            # Downloads is evaluated against $DownloadsThresholdGB independently and is
            # excluded from the general large-folder check so the two thresholds can be
            # tuned separately without double-alerting. Always recorded (with Alerted
            # flag) so JSON consumers see the size even when below threshold.
            $dlLocalGB = 0; $dlCloudGB = 0; $dlAlerted = $false
            if ($stats.LocalByTopFolder.ContainsKey('Downloads')) { $dlLocalGB = To-GB $stats.LocalByTopFolder['Downloads'] }
            if ($stats.CloudByTopFolder.ContainsKey('Downloads')) { $dlCloudGB = To-GB $stats.CloudByTopFolder['Downloads'] }
            if ($dlLocalGB -ge $DownloadsThresholdGB) {
                $LargeItemsFound = $true
                $dlAlerted       = $true
                $Alerts.Add("User '$($stats.User)' Downloads folder is $dlLocalGB GB (Limit: $DownloadsThresholdGB GB)")
            }
            $downloadsResults.Add([pscustomobject]@{
                User        = $stats.User
                SizeGB      = $dlLocalGB
                CloudOnlyGB = $dlCloudGB
                Alerted     = $dlAlerted
            })

            # --- GENERAL LARGE FOLDERS (Downloads excluded - see note above) ---
            foreach ($k in $stats.LocalByTopFolder.Keys) {
                $b = [int64]$stats.LocalByTopFolder[$k]
                if ($b -ge $generalThresholdBytes -and $k -ne 'Downloads') {
                    $LargeItemsFound = $true
                    $folderPath = if ($k -eq '(root)') { $stats.Path } else { Join-Path $stats.Path $k }

                    # Drill-down: largest second-level folders inside this top folder,
                    # so "AppData is 15 GB" becomes actionable (top 5, >= 1 GB each).
                    $prefixKey = "$k\"
                    $subs = @($stats.LocalBySubFolder.GetEnumerator() |
                              Where-Object { $_.Key.StartsWith($prefixKey) -and $_.Value -ge 1GB } |
                              Sort-Object Value -Descending |
                              Select-Object -First 5 |
                              ForEach-Object {
                                  [pscustomobject]@{
                                      Name   = $_.Key.Substring($prefixKey.Length)
                                      SizeGB = To-GB ([int64]$_.Value)
                                  }
                              })

                    $largeItemsResults.Add([pscustomobject]@{
                        ItemType   = 'Folder'
                        User       = $stats.User
                        Name       = $k
                        SizeGB     = To-GB $b
                        Path       = $folderPath
                        Synced     = [bool]$stats.SyncedTopFolders.ContainsKey($k)
                        SubFolders = $subs
                    })
                }
            }

            # --- LARGE FILES ---
            if ($stats.LargeFiles.Count -gt 0) {
                $LargeItemsFound = $true
                foreach ($lf in $stats.LargeFiles) { $largeItemsResults.Add($lf) }
            }
        }
    }

} catch {
    # Write-Output (not Write-Host) so N-able captures the error in task output.
    Write-Output "CRITICAL SCRIPT ERROR: $($_.Exception.Message)"
    exit 4
}

# ==============================================================================
# OUTPUT PHASE
# ==============================================================================

if ($AsJson) {
    # Single compressed JSON line - ideal for N-able custom field ingestion.
    [pscustomobject]@{
        Version           = '6.2'
        LowSpaceFound     = $LowSpaceFound
        LargeItemsFound   = $LargeItemsFound
        Alerts            = $Alerts
        Drives            = $driveResults
        RecycleBins       = $recycleBinResults
        RecycleBinByUser  = $recycleBinPerUser
        Downloads         = $downloadsResults
        LargeItems        = $largeItemsResults
        Profiles          = $profileResults
        SyncRoots         = $sharePointResults
        # Back-compat: flat path list as consumed by v5.x integrations.
        SharePointPaths   = @($sharePointResults | ForEach-Object { $_.Path })
    } | ConvertTo-Json -Depth 6 -Compress | Write-Output
}
else {
    # --- ALERTS BANNER ---
    if ($Alerts.Count -gt 0) {
        Write-Host "=== ACTION REQUIRED ===" -ForegroundColor Yellow
        foreach ($alert in $Alerts) { Write-Host "[WARN] $alert" -ForegroundColor Yellow }
        Write-Host ""
    } else {
        Write-Host "[ OK ] No storage thresholds exceeded.`n" -ForegroundColor Green
    }

    # --- DRIVE STATUS ---
    Write-Host "=== DRIVE STATUS ===" -ForegroundColor Cyan
    foreach ($d in $driveResults) {
        Write-Host ("{0,-5} | {1,6}% Free | {2,7} GB Used / {3,7} GB Total" -f $d.Drive, $d.PercentFree, $d.UsedGB, $d.TotalGB)
    }
    Write-Host ""

    # --- RECYCLE BIN STATUS ---
    if ($IncludeRecycleBin -and $recycleBinResults.Count -gt 0) {
        Write-Host "=== RECYCLE BIN (per drive) ===" -ForegroundColor Cyan
        foreach ($rb in $recycleBinResults) {
            $marker = if ($rb.RecycleBinGB -ge $RecycleBinThresholdGB) { "[WARN]" } else { "[ OK ]" }
            Write-Host ("{0} {1,-5} | {2,7} GB" -f $marker, $rb.Drive, $rb.RecycleBinGB)
        }
        if ($recycleBinPerUser.Count -gt 0) {
            Write-Host "--- Per user ---" -ForegroundColor DarkCyan
            $recycleBinPerUser | Sort-Object SizeGB -Descending | ForEach-Object {
                Write-Host ("       {0,-25} | {1,7} GB | {2}" -f $_.User, $_.SizeGB, $_.Drive)
            }
        }
        Write-Host ""
    }

    # --- USER PROFILES ---
    if ($profileResults.Count -gt 0) {
        Write-Host "=== USER PROFILES (Local = on disk | Cloud = dehydrated | Synced = under OneDrive/SP root) ===" -ForegroundColor Cyan
        Write-Host ("{0,-20} | {1,9} | {2,9} | {3,9}" -f 'User', 'Local GB', 'Cloud GB', 'Synced GB') -ForegroundColor DarkCyan
        $profileResults | Sort-Object LocalGB -Descending | ForEach-Object {
            $note = if ($_.SkippedDirs -gt 0) { " ($($_.SkippedDirs) dirs unreadable)" } else { '' }
            Write-Host ("{0,-20} | {1,9} | {2,9} | {3,9}{4}" -f $_.User, $_.LocalGB, $_.CloudOnlyGB, $_.SyncedGB, $note)
        }
        Write-Host ""
    }

    # --- DOWNLOADS SUMMARY ---
    $downloadsWithData = $downloadsResults | Where-Object { $_.SizeGB -gt 0 -or $_.CloudOnlyGB -gt 0 }
    if ($downloadsWithData) {
        Write-Host ("=== DOWNLOADS (Alert threshold: {0} GB local) ===" -f $DownloadsThresholdGB) -ForegroundColor Cyan
        $downloadsWithData | Sort-Object SizeGB -Descending | ForEach-Object {
            $marker = if ($_.Alerted) { "[WARN]" } else { "[ OK ]" }
            Write-Host ("{0} {1,-20} | {2,7} GB local | {3,7} GB cloud-only" -f $marker, $_.User, $_.SizeGB, $_.CloudOnlyGB)
        }
        Write-Host ""
    }

    # --- LARGE ITEMS ---
    if ($largeItemsResults.Count -gt 0) {
        Write-Host ("=== LARGE ITEMS FOUND (> {0} GB) === ([SYNC] = under a OneDrive/SharePoint root)" -f $LargeThresholdGB) -ForegroundColor Cyan
        $largeItemsResults | Sort-Object SizeGB -Descending | ForEach-Object {
            $tag = if ($_.Synced) { '[SYNC]' } else { '      ' }
            Write-Host ("{0} {1,-15} | {2,-6} | {3,7} GB | {4}" -f $tag, $_.User, $_.ItemType, $_.SizeGB, $_.Path)
            # Drill-down: biggest second-level folders inside a flagged top folder.
            if ($_.ItemType -eq 'Folder' -and $_.SubFolders -and $_.SubFolders.Count -gt 0) {
                foreach ($sf in $_.SubFolders) {
                    Write-Host ("       {0,-15} |        | {1,7} GB |   +-- {2}" -f '', $sf.SizeGB, $sf.Name) -ForegroundColor DarkGray
                }
            }
        }
        Write-Host ""
    }

    # --- SYNC ROOTS ---
    if ($sharePointResults.Count -gt 0) {
        Write-Host "=== ONEDRIVE / SHAREPOINT SYNC ROOTS ===" -ForegroundColor Cyan
        $sharePointResults | Sort-Object Type, User | ForEach-Object {
            Write-Host ("[{0,-16}] {1,-20} | {2}" -f $_.Type, $_.User, $_.Path)
        }
        Write-Host ""
    }
}

# ---------------- EXIT CODES ----------------
if ($LowSpaceFound -and $LargeItemsFound) { exit 3 }
elseif ($LowSpaceFound)                   { exit 1 }
elseif ($LargeItemsFound)                 { exit 2 }
else                                      { exit 0 }