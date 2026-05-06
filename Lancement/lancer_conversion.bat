@echo off
chcp 65001 >nul
title Conversion Cegid → Mon-Planning
cd /d "%~dp0"
echo.
echo ============================================
echo   Conversion Cegid ^> Mon-Planning CSV
echo ============================================
echo.
python cegid_to_monplanning.py
echo.
pause
