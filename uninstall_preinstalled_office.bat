@echo off
setlocal EnableExtensions

rem ============================================================================
rem  Remove pre-installed Office / Copilot / OneDrive / HP  -  launcher
rem  Version : 1.7
rem  Credit  : Thibaut VNC
rem ============================================================================

set "PS1=%~dp0uninstall_preinstalled_office.ps1"

title Remove pre-installed bloatware  -  Thibaut VNC

echo.
echo   Launcher starting...
echo   Folder : %~dp0
echo   Script : %PS1%
echo.

rem ---- Is the PowerShell script next to this launcher? ---------------------
if not exist "%PS1%" (
    echo   [X] Could not find the PowerShell script.
    echo.
    echo   Files present in this folder:
    dir /b "%~dp0*.ps1" 2>nul
    echo.
    echo   Keep the .bat and the .ps1 together, with matching names.
    echo   A browser download named "... (1).ps1" will not be found.
    echo.
    pause
    exit /b 1
)

rem ---- Elevate if we are not running as administrator -----------------------
net session >nul 2>&1
if errorlevel 1 (
    echo   Not elevated - requesting administrator rights...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    if errorlevel 1 (
        echo.
        echo   [X] Elevation was cancelled or failed.
        echo   Right-click this file and choose "Run as administrator".
        echo.
        pause
    )
    exit /b
)

echo   Running as administrator - starting the tool.
echo.

rem ---- Unblock and run ------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%PS1%' -ErrorAction SilentlyContinue; & '%PS1%'"
set "RC=%errorlevel%"

echo.
if not "%RC%"=="0" (
    echo   [X] PowerShell exited with code %RC%.
    echo   The error above explains why. Take a screenshot if you need help.
) else (
    echo   Finished.
)
echo.
pause

endlocal
exit /b 0
