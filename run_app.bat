@echo off
title Abstract of Canvass Generator
cd /d "%~dp0"

python launcher.py

if errorlevel 1 (
    echo.
    echo --------------------------------------------------------
    echo An error occurred while running the application.
    echo --------------------------------------------------------
    pause
)
