@echo off
title Rentora - Rental Management System
echo ========================================================
echo Starting Rentora Development Server...
echo ========================================================
cd /d "%~dp0rental1"
"%~dp0.venv\Scripts\python.exe" manage.py runserver 127.0.0.1:8000
pause
