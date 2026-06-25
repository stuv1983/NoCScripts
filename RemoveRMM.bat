@echo off
:: --- Check for admin rights ---
>nul 2>&1 net session || (
    echo Requesting administrator access...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: --- Variables ---
set "AGENTDIR=C:\Program Files (x86)\Advanced Monitoring Agent"
set "OLDFILE=%AGENTDIR%\settings.ini"
set "NEWFILE=%AGENTDIR%\settings.old.ini"

:: --- Stop related services ---
echo Stopping related services...
for %%S in (
    "Advanced Monitoring Agent"
    "Advanced Monitoring Agent Network"
    "N-able Windows Agent"
    "Take Control Agent"
    "Take Control Agent (x86)"
) do (
    sc stop %%~S >nul 2>&1
)

:: --- Kill any leftover processes ---
taskkill /F /IM "AdvancedMonitoringAgent.exe" >nul 2>&1
taskkill /F /IM "WindowsAgent.exe" >nul 2>&1
taskkill /F /IM "TakeControlAgent.exe" >nul 2>&1

:: --- Take ownership and grant full control ---
takeown /F "%OLDFILE%" >nul 2>&1
icacls "%OLDFILE%" /grant administrators:F >nul 2>&1

:: --- Rename (move) the file ---
if exist "%OLDFILE%" (
    echo Renaming settings.ini to settings.old.ini...
    rename "%OLDFILE%" "settings.old.ini"
)

:: --- Verify ---
if exist "%NEWFILE%" (
    echo [✓] settings.ini successfully renamed to settings.old.ini.
) else (
    echo [!] Rename failed — file may be locked or protected.
)

pause