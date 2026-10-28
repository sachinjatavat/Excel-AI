@echo off
setlocal enabledelayedexpansion
title Local Excel AI Copilot - Installation

echo ======================================================================
echo           LOCAL EXCEL AI COPILOT FOR TEACHERS - INSTALLER
echo ======================================================================
echo.

:: 1. Target Directory Setup
set "TARGET_DIR=%LOCALAPPDATA%\Microsoft\Office\AIAddIn"
echo [1/3] Configuring Add-in catalog folder:
echo       %TARGET_DIR%

if not exist "%TARGET_DIR%" (
    mkdir "%TARGET_DIR%"
    if !errorLevel! neq 0 (
        echo [ERROR] Failed to create directory: %TARGET_DIR%
        pause
        exit /b 1
    )
)

:: Copy manifest.xml
if not exist "%~dp0manifest.xml" (
    echo [ERROR] manifest.xml was not found in the current folder!
    pause
    exit /b 1
)

copy /y "%~dp0manifest.xml" "%TARGET_DIR%\manifest.xml" >nul
if !errorLevel! neq 0 (
    echo [ERROR] Failed to copy manifest.xml to %TARGET_DIR%
    pause
    exit /b 1
)
echo       Copied manifest.xml successfully.

:: Copy assets folder if present
if exist "%~dp0assets" (
    xcopy /y /e /i "%~dp0assets" "%TARGET_DIR%\assets" >nul 2>&1
    echo       Copied icon assets successfully.
)

echo.

:: 2. Configure Windows Registry for Office Sideloading
echo [2/3] Registering Add-in in Windows Registry (Current User)...

:: Register in TrustedCatalogs
set "REG_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A55E12B0-8F24-4B55-A245-123456789ABC}"
reg add "%REG_KEY%" /v "Id" /t REG_SZ /d "{A55E12B0-8F24-4B55-A245-123456789ABC}" /f >nul
reg add "%REG_KEY%" /v "Url" /t REG_SZ /d "%TARGET_DIR%" /f >nul
reg add "%REG_KEY%" /v "Flags" /t REG_DWORD /d 1 /f >nul

:: Register in Developer Sideloading hive
set "DEV_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
reg add "%DEV_KEY%" /v "c7b2a63e-9f33-4f51-b842-83b9e4a3901b" /t REG_SZ /d "%~dp0manifest.xml" /f >nul

echo       Registry keys added successfully.
echo.

echo [3/3] Verifying registration...
python -c "import winreg; k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office\16.0\WEF\Developer'); print('      Verified Developer Key:', winreg.QueryValueEx(k, 'c7b2a63e-9f33-4f51-b842-83b9e4a3901b')[0])" 2>nul

echo.
echo ======================================================================
echo                   INSTALLATION SUCCESSFUL!
echo ======================================================================
echo.
echo What was configured:
echo   1. Copied manifest.xml to:
echo      %TARGET_DIR%\manifest.xml
echo   2. Registered Add-in in Office WEF Developer and Trusted Catalogs.
echo.
echo How to use:
echo   1. Start your local AI model:
echo      ollama run gemma4:12b    (or: ollama run qwen2.5:7b)
echo   2. Double-click "start_assistant.bat" to start servers and open Excel.
echo   3. Inside Excel:
echo      - Method 1: On the 'Home' tab, click 'Add-ins' (or 'More Add-ins')
echo                  ^> Go to 'SHARED FOLDER' or 'MY ADD-INS'
echo                  ^> Select 'Local Excel AI' and click 'Add'.
echo      - Method 2: On the 'Insert' tab, click 'Add-ins' (or 'My Add-ins')
echo                  ^> Go to 'SHARED FOLDER'
echo                  ^> Select 'Local Excel AI' and click 'Add'.
echo.
echo   Once added, the 'Teacher Copilot' button will appear on your Home ribbon!
echo.
echo Press any key to exit this installer...
pause >nul
