@echo off
chcp 65001 >nul
title [Al-Neelain University] Multi-Agent Clinical AI System (Streamlit Dashboard)
echo ======================================================================
echo    AL-NEELAIN UNIVERSITY - MASTER'S THESIS PROJECT
echo    Multi-Agent Clinical AI Decision Support System for Diabetes
echo    Author: Mohamed Abuelgasim - Supervisor: Dr. Fakhreldeen Saeed
echo ======================================================================
echo.

:: 1. Check Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not added to your system PATH.
    echo Please install Python (3.10+) from https://www.python.org/
    echo and ensure "Add Python to PATH" is checked during setup.
    echo.
    pause
    exit /b 1
)

:: 2. Check and install dependencies
echo [CHECK] Verifying required Python dependencies...
python -c "import streamlit, sklearn, pandas, pydantic, docx" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [SETUP] Required libraries not found. Installing dependencies from requirements.txt...
    echo [INFO] This initial setup happens only once. Please wait...
    echo.
    python -m pip install --upgrade pip
    if exist "%~dp0requirements.txt" (
        python -m pip install -r "%~dp0requirements.txt"
    ) else if exist "%~dp0..\requirements.txt" (
        python -m pip install -r "%~dp0..\requirements.txt"
    )
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Failed to install dependencies. Please run 'pip install -r requirements.txt' manually.
        pause
        exit /b 1
    )
    echo.
    echo [SUCCESS] Dependencies installed successfully!
    echo.
)

:: 3. Launch App
echo [STARTING] Launching Interactive Streamlit Dashboard...
echo [INFO] Once started, your web browser will automatically open at:
echo        http://localhost:8501
echo.
python -m streamlit run app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Streamlit encountered an issue starting the application.
)
pause
