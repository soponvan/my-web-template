@echo off
title Sheep Farm Management System
cd /d "%~dp0"
echo ===================================================
echo     Sheep Farm Management System (Web Application)
echo ===================================================
echo.
echo Starting server...
echo URL: http://localhost:5000
echo URL: http://127.0.0.1:5000
echo.
echo Press Ctrl+C to stop server.
echo.
venv\Scripts\python.exe app.py
pause
