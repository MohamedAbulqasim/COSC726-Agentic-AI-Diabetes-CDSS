@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title [Al-Neelain University] Multi-Agent Clinical AI System (Streamlit Dashboard)

echo ======================================================================
echo    AL-NEELAIN UNIVERSITY - MASTER'S THESIS PROJECT
echo    Multi-Agent Clinical AI Decision Support System for Diabetes
echo    Author: Mohamed Abuelgasim - Supervisor: Dr. Fakhreldeen Saeed
echo ======================================================================
echo.

:: 1. Detect Python command (try 'python', then 'py')
set "PY_CMD="
python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PY_CMD=python"
) else (
    py --version >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set "PY_CMD=py"
    )
)

if "%PY_CMD%"=="" goto :NO_PYTHON

echo [INFO] Using Python: %PY_CMD%
%PY_CMD% --version
echo.

:: 2. Check if required libraries are installed
echo [CHECK] Verifying required Python dependencies...
%PY_CMD% -c "import streamlit, sklearn, pandas, pydantic, docx" >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto :INSTALL_DEPS

goto :LAUNCH_APP

:INSTALL_DEPS
echo [SETUP] Required libraries not found. Installing dependencies from requirements.txt...
echo [INFO] This initial setup happens only once. Please wait a moment...
echo.
%PY_CMD% -m pip install --upgrade pip
if exist "%~dp0requirements.txt" (
    %PY_CMD% -m pip install -r "%~dp0requirements.txt"
) else if exist "%~dp0New_Project_After_Agents\requirements.txt" (
    %PY_CMD% -m pip install -r "%~dp0New_Project_After_Agents\requirements.txt"
)

if %ERRORLEVEL% NEQ 0 goto :INSTALL_FAILED
echo.
echo [SUCCESS] Dependencies installed successfully!
echo.
goto :LAUNCH_APP

:LAUNCH_APP
cd /d "%~dp0New_Project_After_Agents"
echo [STARTING] Launching Streamlit Web Application...
echo [INFO] Once started, your web browser will automatically open at:
echo        http://localhost:8501
echo.
%PY_CMD% -m streamlit run app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Streamlit encountered an issue starting the application.
)
pause
exit /b 0

:NO_PYTHON
echo ======================================================================
echo [ERROR] Python is not installed or not added to your system PATH!
echo.
echo Solution:
echo 1. Download and install Python 3.10, 3.11, or 3.12 from:
echo    https://www.python.org/downloads/
echo 2. During installation, make sure to CHECK the box:
echo    [x] "Add python.exe to PATH"
echo ======================================================================
echo.
pause
exit /b 1

:INSTALL_FAILED
echo ======================================================================
echo [ERROR] Failed to install dependencies automatically.
echo Please open PowerShell and run:
echo    pip install -r requirements.txt
echo ======================================================================
echo.
pause
exit /b 1
