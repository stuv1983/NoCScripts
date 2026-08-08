Write-Host "=== N-sight Patch Scan / Windows Update Check ===" -ForegroundColor Cyan
Write-Host "Device: $env:COMPUTERNAME"
Write-Host "Date: $(Get-Date)"
Write-Host ""

Write-Host "=== Windows Update services ===" -ForegroundColor Cyan
Get-CimInstance Win32_Service |
    Where-Object Name -in @('wuauserv','BITS','CryptSvc','msiserver') |
    Select-Object Name, State, StartMode, Status |
    Format-Table -AutoSize

Write-Host "=== Pending reboot indicators ===" -ForegroundColor Cyan
$rebootChecks = [ordered]@{
    CBSRebootPending = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    WindowsUpdateRebootRequired = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    PendingFileRenameOperations = $null -ne (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue)
}

[pscustomobject]$rebootChecks | Format-List

Write-Host "=== System drive space ===" -ForegroundColor Cyan
Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" |
    Select-Object DeviceID,
        @{Name='FreeGB';Expression={[math]::Round($_.FreeSpace / 1GB,2)}},
        @{Name='TotalGB';Expression={[math]::Round($_.Size / 1GB,2)}} |
    Format-Table -AutoSize

Write-Host "=== Recent Windows Update errors ===" -ForegroundColor Cyan
Get-WinEvent -FilterHashtable @{
    LogName='System'
    ProviderName='Microsoft-Windows-WindowsUpdateClient'
    StartTime=(Get-Date).AddDays(-7)
    Level=2,3
} -ErrorAction SilentlyContinue |
    Select-Object -First 20 TimeCreated, Id, LevelDisplayName, Message |
    Format-List

Write-Host "=== Recent update history ===" -ForegroundColor Cyan
try {
    (New-Object -ComObject Microsoft.Update.Session).QueryHistory('',0,20) |
        Select-Object Date, Title,
            @{Name='Result';Expression={
                switch ($_.ResultCode) {
                    2 {'Succeeded'}
                    3 {'Succeeded with errors'}
                    4 {'Failed'}
                    5 {'Aborted'}
                    default {$_.ResultCode}
                }
            }} |
        Format-Table -AutoSize -Wrap
}
catch {
    Write-Warning "Unable to retrieve update history: $($_.Exception.Message)"
}

Write-Host "=== Windows Update search test - no downloads or installations ===" -ForegroundColor Cyan
$job = Start-Job -ScriptBlock {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result = $searcher.Search("IsInstalled=0 and IsHidden=0")

    [pscustomobject]@{
        ResultCode = $result.ResultCode
        MissingUpdates = $result.Updates.Count
    }
}

$completed = Wait-Job $job -Timeout 300

if ($completed) {
    Receive-Job $job
}
else {
    Write-Warning "Windows Update search did not complete within 5 minutes. This supports the N-sight WUA timeout error."
    Stop-Job $job -ErrorAction SilentlyContinue
}

Remove-Job $job -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Check completed - no updates installed and no reboot performed ===" -ForegroundColor Green