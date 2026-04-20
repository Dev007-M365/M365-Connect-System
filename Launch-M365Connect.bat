REM What This BAT File Does (In Plain English)
REM Auto‑detects your PowerShell script
REM Forces PowerShell 7
REM Checks required modules
REM Launches your script cleanly
REM Works on ANY machine - As long as PowerShell 7 is installed
REM Kills any stuck PowerShell 7 processes
REM Detects missing modules
REM Installs missing modules automatically
REM Launches your WPF GUI cleanly


@echo off
setlocal

echo ==========================================
echo   M365-Connect SYSTEM Launcher (pwsh 7)
echo ==========================================
echo.

REM ---------------------------------------------------------
REM 1) Kill any stuck PowerShell 7 processes
REM ---------------------------------------------------------
echo Checking for stuck PowerShell 7 processes...
tasklist | findstr /I "pwsh.exe" >nul 2>&1
if %errorlevel%==0 (
    echo Found running pwsh.exe processes. Terminating...
    taskkill /IM pwsh.exe /F >nul 2>&1
    echo Processes terminated.
) else (
    echo No stuck pwsh.exe processes found.
)
echo.

REM ---------------------------------------------------------
REM 2) Define script directory
REM ---------------------------------------------------------
set SCRIPT_DIR=%~dp0

REM ---------------------------------------------------------
REM 3) Auto-detect ANY .ps1 script in the folder
REM ---------------------------------------------------------
set SCRIPT_PATH=

for %%F in ("%SCRIPT_DIR%*.ps1") do (
    set SCRIPT_PATH=%%F
    goto found
)

echo ERROR: No PowerShell script (.ps1) found in:
echo   %SCRIPT_DIR%
echo.
pause
exit /b 1

:found
echo Found script:
echo   %SCRIPT_PATH%
echo.

REM ---------------------------------------------------------
REM 4) Check for PowerShell 7
REM ---------------------------------------------------------
where pwsh >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: PowerShell 7 not found.
    echo Install from: https://aka.ms/powershell
    pause
    exit /b 1
)

REM ---------------------------------------------------------
REM 5) Run updater.ps1 (self-update system)
REM ---------------------------------------------------------
if exist "%SCRIPT_DIR%updater.ps1" (
    echo Checking for updates...
    pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%updater.ps1"
    echo.
)

REM ---------------------------------------------------------
REM 6) Check required modules
REM ---------------------------------------------------------
echo Checking required PowerShell modules...
pwsh -NoLogo -NoProfile -Command ^
    " $modules = 'Microsoft.Graph','ExchangeOnlineManagement','PnP.PowerShell','MicrosoftTeams' ;" ^
    " $missing = @() ;" ^
    " foreach ($m in $modules) {" ^
    "   if (-not (Get-Module -ListAvailable -Name $m)) {" ^
    "       Write-Host 'Missing module:' $m -ForegroundColor Red ;" ^
    "       $missing += $m ;" ^
    "   } else {" ^
    "       Write-Host 'OK:' $m -ForegroundColor Green" ^
    "   }" ^
    " } ;" ^
    " if ($missing.Count -gt 0) { exit 2 } else { exit 0 }"

REM ---------------------------------------------------------
REM 7) Auto-install missing modules
REM ---------------------------------------------------------
if %errorlevel%==2 (
    echo.
    echo Missing modules detected. Auto-installing...
    echo.

    pwsh -NoLogo -NoProfile -Command ^
        " Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber ;" ^
        " Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber ;" ^
        " Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber ;" ^
        " Install-Module MicrosoftTeams -Scope CurrentUser -Force -AllowClobber ;"

    echo.
    echo Module installation complete.
    echo.
)

REM ---------------------------------------------------------
REM 8) Launch the GUI script
REM ---------------------------------------------------------
echo Launching M365-Connect SYSTEM...
echo.

pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
endlocal
