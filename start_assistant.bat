@echo off
setlocal enabledelayedexpansion
title Local Excel AI Copilot - Service Runner

cd /d "%~dp0"

echo ======================================================================
echo              LOCAL EXCEL AI COPILOT FOR TEACHERS
echo ======================================================================
echo.

:: Check Python installation
python --version >nul 2>&1
if !errorLevel! neq 0 (
    echo [ERROR] Python was not found in your system PATH!
    echo Please install Python 3.10+ and make sure "Add Python to PATH" is checked.
    echo.
    pause
    exit /b 1
)

:: Ensure Add-in Registration in Office WEF Developer and Catalogs
echo [*] Verifying Office Add-in Registration...
set "TARGET_DIR=%LOCALAPPDATA%\Microsoft\Office\AIAddIn"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%" >nul 2>&1
copy /y "%~dp0manifest.xml" "%TARGET_DIR%\manifest.xml" >nul 2>&1
if exist "%~dp0assets" xcopy /y /e /i "%~dp0assets" "%TARGET_DIR%\assets" >nul 2>&1

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "c7b2a63e-9f33-4f51-b842-83b9e4a3901b" /t REG_SZ /d "%~dp0manifest.xml" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A55E12B0-8F24-4B55-A245-123456789ABC}" /v "Id" /t REG_SZ /d "{A55E12B0-8F24-4B55-A245-123456789ABC}" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A55E12B0-8F24-4B55-A245-123456789ABC}" /v "Url" /t REG_SZ /d "%TARGET_DIR%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A55E12B0-8F24-4B55-A245-123456789ABC}" /v "Flags" /t REG_DWORD /d 1 /f >nul 2>&1

:: Check Ollama
tasklist /fi "IMAGENAME eq ollama.exe" 2>nul | find /i "ollama.exe" >nul
if !errorLevel! neq 0 (
    echo [*] Starting Ollama in background...
    start "Ollama Engine" /min ollama serve
    timeout /t 2 /nobreak >nul
)

:: 1. Start FastAPI Backend on Port 8000
echo [1/3] Starting AI Backend (FastAPI on http://127.0.0.1:8000)...
start "Excel AI Backend (Port 8000)" /min python backend.py

:: 2. Start HTTP Server for Taskpane on Port 3000
echo [2/3] Starting Add-in Web Server (http://localhost:3000)...
start "Excel AI Web Server (Port 3000)" /min python -m http.server 3000 --bind 127.0.0.1

:: 3. Wait 2 seconds for services to initialize
echo [3/3] Waiting 2 seconds for services to initialize...
timeout /t 2 /nobreak >nul

:: 4. Launch Microsoft Excel
echo.
echo [*] Launching Microsoft Excel...
start excel.exe

echo.
echo ======================================================================
echo                    SERVICES RUNNING SUCCESSFULLY
echo ======================================================================
echo.
echo  * Backend API:       http://127.0.0.1:8000
echo  * Taskpane Web:      http://localhost:3000/taskpane.html
echo  * LLM Engine:        Ollama (http://localhost:11434)
echo.
echo  ------------------------------------------------------------------
echo  WHERE TO FIND TEACHER COPILOT IN EXCEL:
echo  ------------------------------------------------------------------
echo  1. In Excel, open a blank workbook.
echo.
echo  2. Look on the Home ribbon (on the right) for the 'Teacher Copilot'
echo     button.
echo.
echo  3. IF YOU DO NOT SEE THE BUTTON YET:
echo     - In Excel 365: Click 'Home' tab ^> 'Add-ins' button (far right)
echo                     ^> 'More Add-ins' ^> 'SHARED FOLDER' or 'MY ADD-INS'
echo                     ^> Click 'Local Excel AI' ^> 'Add'.
echo.
echo     - In Excel 2016/2019/2021: Click 'Insert' tab ^> 'Add-ins' (or 'My Add-ins')
echo                                ^> 'SHARED FOLDER' tab
echo                                ^> Click 'Local Excel AI' ^> 'Add'.
echo.
echo  4. The Teacher AI Copilot sidebar will open on the right of your sheet!
echo  ------------------------------------------------------------------
echo.
echo  NOTE:
echo    Keep this window open while using the add-in.
echo    To stop all servers when finished, close this window or press any key.
echo ======================================================================
echo.
pause

:: Clean up processes on exit
echo Stopping background servers...
taskkill /fi "WINDOWTITLE eq Excel AI Backend (Port 8000)*" /f >nul 2>&1
taskkill /fi "WINDOWTITLE eq Excel AI Web Server (Port 3000)*" /f >nul 2>&1
echo Done.
exit /b 0
