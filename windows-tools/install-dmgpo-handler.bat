@echo off
REM ============================================================
REM  DMG PO URL handler installer
REM  Registers the dmgpo:// scheme so clicking PO links in the
REM  Production Control web app opens the right PDF/folder.
REM
REM  Run ONCE per user. Approve the UAC prompt.
REM  Re-running is safe (idempotent — overwrites existing keys).
REM ============================================================

setlocal

set "INSTALL_DIR=%LOCALAPPDATA%\DMG"
set "LAUNCHER=%INSTALL_DIR%\dmgpo-launcher.ps1"
set "SOURCE=%~dp0dmgpo-launcher.ps1"

if not exist "%SOURCE%" (
    echo [X] Cannot find dmgpo-launcher.ps1 next to this installer.
    echo     Expected: %SOURCE%
    pause
    exit /b 1
)

echo Installing DMG PO handler for %USERNAME% ...

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
copy /Y "%SOURCE%" "%LAUNCHER%" >nul
if errorlevel 1 (
    echo [X] Failed to copy launcher to %LAUNCHER%
    pause
    exit /b 1
)

REM Register dmgpo:// in HKCU (no admin needed for current-user install)
reg add "HKCU\Software\Classes\dmgpo" /ve /d "URL:DMG PO Opener" /f >nul
reg add "HKCU\Software\Classes\dmgpo" /v "URL Protocol" /d "" /f >nul
reg add "HKCU\Software\Classes\dmgpo\DefaultIcon" /ve /d "%SystemRoot%\System32\shell32.dll,1" /f >nul
reg add "HKCU\Software\Classes\dmgpo\shell\open\command" /ve /d "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%LAUNCHER%\" \"%%1\"" /f >nul
if errorlevel 1 (
    echo [X] Failed to register dmgpo:// scheme.
    pause
    exit /b 1
)

echo.
echo [OK] dmgpo:// handler installed.
echo      Click any PO button in the Production Control app to open the PO.
echo      First click in each browser may show a one-time confirm prompt.
echo.
pause
endlocal
