@echo off
REM ===========================================
REM  Kisbye Consulting – Lav exe
REM  Bygger EXE ud fra businesscasegpt_v9_0_web.py
REM  (tilpas SCRIPT= hvis du ændrer filnavn)
REM ===========================================

cd /d "%~dp0"

set SCRIPT=businesscasegpt_v9_0_web.py
set EXENAME=BusinessCaseGPT.exe

echo.
echo === Bygger %EXENAME% fra %SCRIPT% ===
echo Arbejdsmappe: %cd%
echo.

REM 1) tjek PyInstaller
pyinstaller --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo PyInstaller mangler - installerer...
    python -m pip install --upgrade pip
    python -m pip install pyinstaller
)

REM 2) ryd op fra sidst
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist %EXENAME% del %EXENAME%

REM 3) saml --add-data
set ADDDATA=

REM hvis der findes en static-mappe, så tag den med
if exist static (
    set ADDDATA=--add-data "static;static"
)

REM hvis logo ligger ved siden af .py, så tag det med
if exist kisbye_logo.png (
    set ADDDATA=%ADDDATA% --add-data "kisbye_logo.png;."
)

if exist kisbye_logo.ico (
    set ADDDATA=%ADDDATA% --add-data "kisbye_logo.ico;."
)

echo Kører PyInstaller...
pyinstaller --noconfirm --onefile --windowed ^
 --name "%EXENAME%" ^
 %ADDDATA% ^
 "%SCRIPT%"

echo.
echo ===========================================
echo Færdig. Se dist\%EXENAME%
echo ===========================================
pause
