@echo off
setlocal enabledelayedexpansion
title PDF-XChange Stamp Importer

:: ============================================================
:: PDF-XChange JavaScript Stamp Importer
::
:: Guides the user through importing JavaScript-enabled PDF
:: stamps into PDF-XChange Editor step by step.
::
:: Exit Codes: 0=Success, 1=Failure
:: ============================================================

:: --- Configuration ---
set "STAMPS_DIR=%APPDATA%\Tracker Software\PDFXEditor\3.0\Stamps"
set "PDFXCHANGE_EXE=C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe"
set "SCRIPT_DIR=%~dp0"
set "STAGING_DIR=%TEMP%\StampImport"

:: ============================================================
::   PRE-FLIGHT CHECKS
:: ============================================================

:: Check PDF-XChange is installed
if not exist "%PDFXCHANGE_EXE%" (
    echo.
    echo   [ERROR] PDF-XChange Editor is not installed.
    echo   Expected: %PDFXCHANGE_EXE%
    echo.
    echo   Install PDF-XChange Editor and try again.
    echo.
    pause
    exit /b 1
)

:: Check Stamps directory exists
if not exist "%STAMPS_DIR%" (
    echo.
    echo   [ERROR] The PDF-XChange Stamps directory does not exist.
    echo   Expected: %STAMPS_DIR%
    echo.
    echo   Open PDF-XChange Editor at least once first, then try again.
    echo.
    pause
    exit /b 1
)

:: Check PDF-XChange is NOT already running
powershell -NoProfile -Command "if (Get-Process -Name PDFXEdit -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
if !errorlevel! neq 0 (
    echo.
    echo   [ERROR] PDF-XChange Editor is currently running.
    echo   Close all PDF-XChange windows before running this script.
    echo.
    pause
    exit /b 1
)

:: Discover available stamp PDFs in the script's directory
set "pdfCount=0"
for %%F in ("%SCRIPT_DIR%*.pdf") do (
    set /a pdfCount+=1
)
if !pdfCount! equ 0 (
    echo.
    echo   [ERROR] No stamp PDF files found in:
    echo   %SCRIPT_DIR%
    echo.
    echo   The stamp PDFs should be in the same folder as this script.
    echo.
    pause
    exit /b 1
)

:: ============================================================
::   STAMP SELECTION
:: ============================================================

echo.
echo   ============================================
echo      PDF-XChange Stamp Importer
echo   ============================================
echo.
echo   Available stamps:
echo.

set "idx=0"
for %%F in ("%SCRIPT_DIR%*.pdf") do (
    set /a idx+=1
    set "stamp_!idx!=%%~nxF"
    set "stampName_!idx!=%%~nF"
    echo     !idx!. %%~nF
)

echo.
set /p "choice=  Select a stamp to import (1-!idx!): "

:: Validate selection
set "selectedStamp=!stamp_%choice%!"
if not defined selectedStamp (
    echo.
    echo   [ERROR] Invalid selection. Enter a number between 1 and !idx!.
    echo.
    pause
    exit /b 1
)

set "selectedName=!stampName_%choice%!"

:: ============================================================
::   PREPARE STAGING AREA
:: ============================================================

:: Clean and create staging directory
if exist "%STAGING_DIR%" rd /s /q "%STAGING_DIR%" >nul 2>&1
mkdir "%STAGING_DIR%"

:: Copy selected stamp to staging
copy "%SCRIPT_DIR%!selectedStamp!" "%STAGING_DIR%\!selectedStamp!" >nul
if !errorlevel! neq 0 (
    echo.
    echo   [ERROR] Failed to copy stamp file to staging area.
    echo.
    pause
    exit /b 1
)

echo.
echo   Selected: !selectedName!
echo.

:: ============================================================
::   BASELINE SNAPSHOT
:: ============================================================

:: Capture current state of stamps directory before user does anything
dir "%STAMPS_DIR%" /b /a-d > "%STAGING_DIR%\baseline.txt" 2>nul
:: Create empty file if stamps dir was empty
if not exist "%STAGING_DIR%\baseline.txt" type nul > "%STAGING_DIR%\baseline.txt"

:: ============================================================
::   STEP 1: OPEN PDF-XCHANGE
:: ============================================================

echo   ============================================
echo      STEP 1 of 3: Opening PDF-XChange
echo   ============================================
echo.
echo   Opening the stamp file in PDF-XChange Editor...

start "" "%PDFXCHANGE_EXE%" "%STAGING_DIR%\!selectedStamp!"

echo   Done. PDF-XChange should now be open with your stamp.
echo.

:: ============================================================
::   STEP 2: INSTRUCTIONS FOR CREATING THE STAMP
:: ============================================================

echo   ============================================
echo      STEP 2 of 3: Create the Stamp
echo   ============================================
echo.
echo   In PDF-XChange, follow these steps:
echo.
echo     a) Click the STAMPS menu in the top ribbon
echo     b) Click "Stamps Palette"
echo     c) At the bottom of the palette, click "Add New"
echo     d) Choose "New Stamp from Active Document"
echo     e) Give your stamp any unique name
echo.
echo     *** IMPORTANT ***
echo     f) Click the "Add New" button next to Target Collection
echo        to create a NEW collection with any unique name
echo        (e.g. "My Shop Stamps" or "Custom Stamps")
echo.
echo     g) Click OK to save, then OK again to close
echo.

:: ============================================================
::   STEP 3: CLOSE AND CONFIRM
:: ============================================================

echo   ============================================
echo      STEP 3 of 3: Close PDF-XChange
echo   ============================================
echo.
echo   After creating the stamp:
echo     - Close ALL PDF-XChange windows
echo     - You can close without saving the document
echo     - Then come back here and press 'y'
echo.

:waitForClose
set "confirm="
set /p "confirm=  Done? Press 'y' when you have closed PDF-XChange: "
if /i not "!confirm!"=="y" (
    echo.
    echo   No problem. Take your time, then press 'y' when ready.
    echo.
    goto waitForClose
)

:: Verify PDF-XChange is actually closed
powershell -NoProfile -Command "if (Get-Process -Name PDFXEdit -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
if !errorlevel! neq 0 (
    echo.
    echo   [WARNING] PDF-XChange is still running!
    echo   Please close ALL PDF-XChange windows first.
    echo   Check the taskbar for any remaining windows.
    echo.
    goto waitForClose
)

:: ============================================================
::   DETECT AND REPLACE STAMP FILE
:: ============================================================

echo.
echo   Checking for new stamp...

:: Use PowerShell to compare baseline with current state and replace the file
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$stampsDir = $env:APPDATA + '\Tracker Software\PDFXEditor\3.0\Stamps';" ^
    "$stagingDir = '%STAGING_DIR%';" ^
    "$selectedStamp = '!selectedStamp!';" ^
    "$baselinePath = Join-Path $stagingDir 'baseline.txt';" ^
    "$baseline = @(Get-Content $baselinePath -ErrorAction SilentlyContinue);" ^
    "$current = @(Get-ChildItem $stampsDir -File -ErrorAction Stop | Select-Object -ExpandProperty Name);" ^
    "$newFiles = @($current | Where-Object { $_ -notin $baseline });" ^
    "if ($newFiles.Count -eq 0) {" ^
    "  Write-Host '';" ^
    "  Write-Host '  [ERROR] No new stamp file was detected.' -ForegroundColor Red;" ^
    "  Write-Host '';" ^
    "  Write-Host '  This usually means one of these things:';" ^
    "  Write-Host '    - You added the stamp to an EXISTING collection';" ^
    "  Write-Host '      (you must create a NEW collection each time)';" ^
    "  Write-Host '    - You closed PDF-XChange without saving the stamp';" ^
    "  Write-Host '    - The stamp was not created from the active document';" ^
    "  Write-Host '';" ^
    "  Write-Host '  Please try again from the beginning.';" ^
    "  exit 1;" ^
    "}" ^
    "$newFile = $newFiles[0];" ^
    "if ($newFiles.Count -gt 1) {" ^
    "  Write-Host '  [NOTE] Multiple new files detected. Using the most recent one.' -ForegroundColor Yellow;" ^
    "  $newest = Get-ChildItem $stampsDir -File | Where-Object { $_.Name -in $newFiles } | Sort-Object LastWriteTime -Descending | Select-Object -First 1;" ^
    "  $newFile = $newest.Name;" ^
    "}" ^
    "$targetPath = Join-Path $stampsDir $newFile;" ^
    "$sourcePath = Join-Path $stagingDir $selectedStamp;" ^
    "try {" ^
    "  Copy-Item -Path $sourcePath -Destination $targetPath -Force -ErrorAction Stop;" ^
    "  Write-Host '';" ^
    "  Write-Host '  Stamp file replaced successfully.' -ForegroundColor Green;" ^
    "  exit 0;" ^
    "} catch {" ^
    "  Write-Host '';" ^
    "  Write-Host ('  [ERROR] Failed to replace stamp file: ' + $_.Exception.Message) -ForegroundColor Red;" ^
    "  exit 1;" ^
    "}"

if !errorlevel! neq 0 (
    echo.
    pause
    :: Cleanup before exit
    if exist "%STAGING_DIR%" rd /s /q "%STAGING_DIR%" >nul 2>&1
    exit /b 1
)

:: ============================================================
::   CLEANUP AND SUCCESS
:: ============================================================

:: Remove staging directory
if exist "%STAGING_DIR%" rd /s /q "%STAGING_DIR%" >nul 2>&1

echo.
echo   ============================================
echo      SUCCESS!
echo   ============================================
echo.
echo   Your stamp "!selectedName!" has been imported.
echo.
echo   To verify it works:
echo     1. Open any PDF in PDF-XChange
echo     2. Go to Stamps ^> Stamps Palette
echo     3. Find your stamp in the collection you created
echo     4. Place it on the document -- it should prompt
echo        for your name and auto-fill the date
echo.
echo   To import another stamp, run this script again.
echo.
pause
exit /b 0
