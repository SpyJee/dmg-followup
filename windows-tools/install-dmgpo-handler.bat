@echo off
REM ============================================================
REM  DMG PO URL handler installer
REM  Registers the dmgpo:// scheme so clicking PO links in the
REM  Production Control web app opens the right PDF/folder.
REM
REM  Run ONCE per user. No admin required (writes to HKCU only).
REM  Re-running is safe (idempotent — overwrites existing keys).
REM ============================================================

setlocal EnableDelayedExpansion

set "INSTALL_DIR=%LOCALAPPDATA%\DMG"
set "LAUNCHER=%INSTALL_DIR%\dmgpo-launcher.ps1"
set "SOURCE=%~dp0dmgpo-launcher.ps1"
set "FAILED="

if not exist "%SOURCE%" (
    echo [X] Cannot find dmgpo-launcher.ps1 next to this installer.
    echo     Expected: %SOURCE%
    pause
    exit /b 1
)

echo Installing DMG PO handler for %USERNAME% ...
echo.

if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
    if errorlevel 1 (
        echo [X] Failed to create %INSTALL_DIR%
        pause
        exit /b 1
    )
)

copy /Y "%SOURCE%" "%LAUNCHER%" >nul
if errorlevel 1 (
    echo [X] Failed to copy launcher to %LAUNCHER%
    pause
    exit /b 1
)
echo  [+] Copied launcher to %LAUNCHER%

REM -- Register dmgpo:// in HKCU. Each step checked individually --

reg add "HKCU\Software\Classes\dmgpo" /ve /d "URL:DMG PO Opener" /f >nul
if errorlevel 1 set "FAILED=1" & echo  [X] reg add: scheme root

reg add "HKCU\Software\Classes\dmgpo" /v "URL Protocol" /d "" /f >nul
if errorlevel 1 set "FAILED=1" & echo  [X] reg add: URL Protocol

reg add "HKCU\Software\Classes\dmgpo\DefaultIcon" /ve /d "%SystemRoot%\System32\shell32.dll,1" /f >nul
if errorlevel 1 set "FAILED=1" & echo  [X] reg add: DefaultIcon

set "CMD=powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%LAUNCHER%\" \"%%1\""
reg add "HKCU\Software\Classes\dmgpo\shell\open\command" /ve /d "%CMD%" /f >nul
if errorlevel 1 set "FAILED=1" & echo  [X] reg add: shell\open\command

if defined FAILED (
    echo.
    echo [X] One or more registry steps failed. dmgpo:// will not work.
    pause
    exit /b 1
)

echo  [+] Registered dmgpo:// scheme in HKCU

REM -- Verify the install actually took --
reg query "HKCU\Software\Classes\dmgpo\shell\open\command" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [X] Verification failed — dmgpo:// not found after install.
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
