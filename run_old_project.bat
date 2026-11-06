@echo off
chcp 65001 >nul
title [Al-Neelain University] Old Diabetes Project (Tkinter Desktop)
echo ======================================================================
echo    AL-NEELAIN UNIVERSITY - MASTER'S PROJECT
echo    Original Diabetes Prediction App (Before Intelligent Agents)
echo ======================================================================
echo.
cd /d "%~dp0Old_Project_Before_Agents"
echo Launching Old Project from: %CD%
python diabetes_predict.py
pause
