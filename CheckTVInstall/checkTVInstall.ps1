# Audit-TeamViewer12and15.ps1  (PS 5.1 compatible)
$ErrorActionPreference = 'SilentlyContinue'
$Computer = $env:COMPUTERNAME

$RegPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$Results = @()

foreach ($Path in $RegPaths) {

    if ($Path -like '*WOW6432Node*') { $Hive = 'WOW6432Node (32-bit view)' }
    else { $Hive = 'Native (64-bit view)' }

    Get-ItemProperty -Path $Path | Where-Object {
        $_.DisplayName -like 'TeamViewer*'
    } | ForEach-Object {

        $Major = $null
        if ($_.DisplayVersion) { $Major = $_.DisplayVersion.Split('.')[0] }

        $IsTV12 = if ($Major -eq '12') { 'YES' } else { 'NO' }
        $IsTV15 = if ($Major -eq '15') { 'YES' } else { 'NO' }

        $Results += [pscustomobject]@{
            ComputerName   = $Computer
            DisplayName    = $_.DisplayName
            DisplayVersion = $_.DisplayVersion
            TV12           = $IsTV12
            TV15           = $IsTV15
            HivePath       = $Hive
        }
    }
}

if (-not $Results) {
    Write-Output "$Computer | TeamViewer not found"
    exit 0
}

$Results | Sort-Object DisplayName, DisplayVersion | ForEach-Object {
    "$($_.ComputerName) | $($_.DisplayName) | $($_.DisplayVersion) | TV12=$($_.TV12) | TV15=$($_.TV15) | $($_.HivePath)"
}

exit 0
