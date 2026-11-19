@echo off
setlocal
call "%~dp0_python.cmd" "%~dp0catalog-update\launch-receipt-e8.py" %*
if "%~1"=="--check" exit /b %errorlevel%
if errorlevel 1 pause
