@echo off
setlocal EnableExtensions

rem ============================================================================
rem  Remove pre-installed Office / Copilot / OneDrive  -  launcher
rem  Version : 1.2
rem  Credit  : Thibaut VNC
rem
rem  Double-click this file, or right-click > Run as administrator.
rem  It elevates itself, bypasses the execution policy for this run only,
rem  and starts uninstall_preinstalled_office.ps1 from the same folder.
rem ============================================================================

set "PS1=%~dp0uninstall_preinstalled_office.ps1"

title Remove pre-installed Office / Copilot / OneDrive  -  Thibaut VNC

rem ---- Is the PowerShell script next to this launcher? ---------------------
if not exist "%PS1%" (
    echo.
    echo   Could not find:
    echo     %PS1%
    echo.
    echo   Keep remove_preinstalled.bat and uninstall_preinstalled_office.ps1
    echo   together in the same folder.
    echo.
    pause
    exit /b 1
)

rem ---- Elevate if we are not running as administrator -----------------------
net session >nul 2>&1
if errorlevel 1 (
    echo.
    echo   Requesting administrator rights...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo   Elevation was cancelled or failed.
        echo   Right-click this file and choose "Run as administrator".
        echo.
        pause
    )
    exit /b
)

rem ---- Unblock and run in a single PowerShell start -------------------------
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%PS1%' -ErrorAction SilentlyContinue; & '%PS1%'"

set "RC=%errorlevel%"
if not "%RC%"=="0" (
    echo.
    echo   The script exited with code %RC%.
    echo.
    pause
)

endlocal
exit /b 0
