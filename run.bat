@echo off
setlocal
cd /d "%~dp0"
set "TablyfernPython=.venv\Scripts\python.exe"
if not exist "%TablyfernPython%" set "TablyfernPython=..\.venv-standard\Scripts\python.exe"
if not exist "%TablyfernPython%" set "TablyfernPython=python"
if exist "..\.cache\tablyfern-deps\xlsxwriter" set "PYTHONPATH=%CD%\..\.cache\tablyfern-deps;%PYTHONPATH%"
"%TablyfernPython%" app.py
if errorlevel 1 (
    echo See README.md or README.ko.md for setup.
    pause
    exit /b 1
)
