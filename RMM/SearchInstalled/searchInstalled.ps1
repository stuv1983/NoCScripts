<#
.SYNOPSIS
Search installed applications and common install folders.

.DESCRIPTION
Designed for N-able N-sight/N-central scripted tasks.

Checks:
- HKLM 64-bit uninstall registry
- HKLM 32-bit uninstall registry
- HKU per-user uninstall registry where available
- Program Files
- Program Files (x86)
- User AppData locations

.PARAMETER AppKeywords
Comma-separated keywords to search for.
Example:
Zoom,Teams,Chrome,Firefox

.PARAMETER SearchAll
Set to $true to return all detected apps/folders.
Set to $false to only return keyword matches.

.PARAMETER IncludeFileSystem
Set to $true to also check Program Files and AppData folders.

.EXAMPLE
.\Search-InstalledApps.ps1 -AppKeywords "Zoom,Teams" -SearchAll $false

.EXAMPLE
.\Search-InstalledApps.ps1 -SearchAll $true
#>

[CmdletBinding()]
param(
    [string[]]$AppKeywords = @(),
    [bool]$SearchAll = $false,
    [bool]$IncludeFileSystem = $true
)

$ErrorActionPreference = "SilentlyContinue"

$DeviceName = $env:COMPUTERNAME
$Results = New-Object System.Collections.Generic.List[object]

$Keywords = @()

if (-not $SearchAll) {
    # Accepts either a quoted comma-string ("Outlook,Teams") or an unquoted
    # array from the command line (Outlook,Teams -> @("Outlook","Teams")).
    # Splitting each element on "," again handles both cases safely.
    $Keywords = $AppKeywords |
        ForEach-Object { $_ -split "," } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne "" }

    if ($Keywords.Count -eq 0) {
        Write-Output "ERROR: No keywords provided. Enter app names separated by commas, or set SearchAll to True."
        exit 1
    }
}

function Get-MatchedKeyword {
    param(
        [string]$SearchText
    )

    if ($SearchAll) {
        return "ALL"
    }

    foreach ($Keyword in $Keywords) {
        if ($SearchText -like "*$Keyword*") {
            return $Keyword
        }
    }

    return $null
}

function Add-AppResult {
    param(
        [string]$Source,
        [string]$Scope,
        [string]$UserProfile,
        [string]$DisplayName,
        [string]$DisplayVersion,
        [string]$Publisher,
        [string]$InstallLocation,
        [string]$UninstallString,
        [string]$MatchedKeyword
    )

    if ([string]::IsNullOrWhiteSpace($DisplayName) -and [string]::IsNullOrWhiteSpace($InstallLocation)) {
        return
    }

    $Results.Add([PSCustomObject]@{
        DeviceName      = $DeviceName
        Source          = $Source
        Scope           = $Scope
        UserProfile     = $UserProfile
        DisplayName     = $DisplayName
        DisplayVersion  = $DisplayVersion
        Publisher       = $Publisher
        InstallLocation = $InstallLocation
        UninstallString = $UninstallString
        MatchedKeyword  = $MatchedKeyword
    })
}

function Search-RegistryApps {
    param(
        [string]$RegistryPath,
        [string]$Scope,
        [string]$UserProfile = ""
    )

    $Apps = Get-ItemProperty -Path $RegistryPath

    foreach ($App in $Apps) {
        $DisplayName = $App.DisplayName

        if ([string]::IsNullOrWhiteSpace($DisplayName)) {
            continue
        }

        $SearchText = @(
            $App.DisplayName
            $App.DisplayVersion
            $App.Publisher
            $App.InstallLocation
            $App.UninstallString
        ) -join " "

        $MatchedKeyword = Get-MatchedKeyword -SearchText $SearchText

        if ($SearchAll -or $MatchedKeyword) {
            Add-AppResult `
                -Source "Registry" `
                -Scope $Scope `
                -UserProfile $UserProfile `
                -DisplayName $App.DisplayName `
                -DisplayVersion $App.DisplayVersion `
                -Publisher $App.Publisher `
                -InstallLocation $App.InstallLocation `
                -UninstallString $App.UninstallString `
                -MatchedKeyword $MatchedKeyword
        }
    }
}

function Get-UserProfiles {
    $ProfileList = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" |
        Where-Object {
            $_.ProfileImagePath -like "C:\Users\*" -and
            $_.ProfileImagePath -notmatch "\\(Default|Default User|Public|All Users)$" -and
            (Test-Path $_.ProfileImagePath)
        }

    foreach ($Profile in $ProfileList) {
        [PSCustomObject]@{
            SID         = Split-Path $Profile.PSPath -Leaf
            ProfilePath = $Profile.ProfileImagePath
            UserName    = Split-Path $Profile.ProfileImagePath -Leaf
        }
    }
}

function Search-PerUserRegistryApps {
    $Profiles = $Script:CachedUserProfiles

    foreach ($Profile in $Profiles) {
        $Sid = $Profile.SID
        $UserName = $Profile.UserName
        $NtUserDat = Join-Path $Profile.ProfilePath "NTUSER.DAT"

        $LoadedHivePath = "Registry::HKEY_USERS\$Sid"
        $TempHiveName = "Temp_AppSearch_$($Sid -replace '[^A-Za-z0-9]', '_')"
        $TempHivePath = "Registry::HKEY_USERS\$TempHiveName"
        $HiveLoadedByScript = $false

        if (Test-Path $LoadedHivePath) {
            $HiveRoot = $LoadedHivePath
        }
        else {
            & reg.exe load "HKU\$TempHiveName" "$NtUserDat" | Out-Null

            if ($LASTEXITCODE -eq 0) {
                $HiveRoot = $TempHivePath
                $HiveLoadedByScript = $true
            }
            else {
                continue
            }
        }

        $UserRegistryPaths = @(
            "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "$HiveRoot\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($Path in $UserRegistryPaths) {
            Search-RegistryApps `
                -RegistryPath $Path `
                -Scope "PerUserRegistry" `
                -UserProfile $UserName
        }

        if ($HiveLoadedByScript) {
            [GC]::Collect()
            Start-Sleep -Milliseconds 300
            & reg.exe unload "HKU\$TempHiveName" | Out-Null
        }
    }
}

function Get-DirectoriesByDepth {
    param(
        [string]$RootPath,
        [int]$MaxDepth = 2
    )

    if (-not (Test-Path $RootPath)) {
        return
    }

    $SkipFolders = @(
        "\Temp",
        "\Cache",
        "\Caches",
        "\Packages",
        "\Microsoft\Windows",
        "\Microsoft\Edge",
        "\Mozilla\Firefox\Profiles",
        "\Google\Chrome\User Data"
    )

    $Queue = New-Object System.Collections.Queue
    $Queue.Enqueue([PSCustomObject]@{
        Path  = $RootPath
        Depth = 0
    })

    while ($Queue.Count -gt 0) {
        $Current = $Queue.Dequeue()

        if ($Current.Depth -ge $MaxDepth) {
            continue
        }

        $Children = Get-ChildItem -Path $Current.Path -Directory -Force

        foreach ($Child in $Children) {
            $ShouldSkip = $false

            foreach ($Skip in $SkipFolders) {
                if ($Child.FullName -like "*$Skip*") {
                    $ShouldSkip = $true
                    break
                }
            }

            if ($ShouldSkip) {
                continue
            }

            $Child

            $Queue.Enqueue([PSCustomObject]@{
                Path  = $Child.FullName
                Depth = $Current.Depth + 1
            })
        }
    }
}

function Search-FileSystemApps {
    $MachinePaths = @(
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)}
    ) | Where-Object {
        $_ -and (Test-Path $_)
    } | Select-Object -Unique

    foreach ($Path in $MachinePaths) {
        $Folders = Get-DirectoriesByDepth -RootPath $Path -MaxDepth 2

        foreach ($Folder in $Folders) {
            $SearchText = $Folder.Name
            $MatchedKeyword = Get-MatchedKeyword -SearchText $SearchText

            if ($SearchAll -or $MatchedKeyword) {
                Add-AppResult `
                    -Source "FileSystem" `
                    -Scope "ProgramFiles" `
                    -UserProfile "" `
                    -DisplayName $Folder.Name `
                    -DisplayVersion "" `
                    -Publisher "" `
                    -InstallLocation $Folder.FullName `
                    -UninstallString "" `
                    -MatchedKeyword $MatchedKeyword
            }
        }
    }

    $Profiles = $Script:CachedUserProfiles

    foreach ($Profile in $Profiles) {
        $UserName = $Profile.UserName

        $AppDataPaths = @(
            "$($Profile.ProfilePath)\AppData\Local\Programs",
            "$($Profile.ProfilePath)\AppData\Local",
            "$($Profile.ProfilePath)\AppData\Roaming"
        ) | Where-Object {
            Test-Path $_
        }

        foreach ($Path in $AppDataPaths) {
            $Depth = 2

            if ($Path -like "*\AppData\Local\Programs") {
                $Depth = 3
            }

            $Folders = Get-DirectoriesByDepth -RootPath $Path -MaxDepth $Depth

            foreach ($Folder in $Folders) {
                $SearchText = $Folder.Name
                $MatchedKeyword = Get-MatchedKeyword -SearchText $SearchText

                if ($SearchAll -or $MatchedKeyword) {
                    Add-AppResult `
                        -Source "FileSystem" `
                        -Scope "AppData" `
                        -UserProfile $UserName `
                        -DisplayName $Folder.Name `
                        -DisplayVersion "" `
                        -Publisher "" `
                        -InstallLocation $Folder.FullName `
                        -UninstallString "" `
                        -MatchedKeyword $MatchedKeyword
                }
            }
        }
    }
}

# Cache user profiles once — used by both registry and filesystem searches
$Script:CachedUserProfiles = Get-UserProfiles

# Registry checks - machine-wide
Search-RegistryApps `
    -RegistryPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    -Scope "Machine64"

Search-RegistryApps `
    -RegistryPath "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" `
    -Scope "Machine32"

# Registry checks - per-user
Search-PerUserRegistryApps

# File system checks
if ($IncludeFileSystem) {
    Search-FileSystemApps
}

# Remove duplicates
$UniqueResults = $Results |
    Sort-Object DeviceName, Source, Scope, UserProfile, DisplayName, DisplayVersion, InstallLocation -Unique

if (-not $UniqueResults -or $UniqueResults.Count -eq 0) {
    Write-Output "No matching applications found on $DeviceName."
    exit 0
}

$UniqueResults |
    Sort-Object DisplayName, UserProfile, Source |
    Format-Table `
        DeviceName,
        Source,
        Scope,
        UserProfile,
        DisplayName,
        DisplayVersion,
        Publisher,
        InstallLocation,
        MatchedKeyword `
        -AutoSize