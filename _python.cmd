@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONNOUSERSITE=1"
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONPATH=%~dp0.packages"
set "PYCMD=%SCANER_PYTHON%"
if not defined PYCMD if exist "%~dp0.python-path" set /p "PYCMD="<"%~dp0.python-path"
if not defined PYCMD goto try_launcher
if not exist "%PYCMD%" goto try_launcher
"%PYCMD%" -c "import struct,sys; assert sys.version_info[:2]==(3,12) and struct.calcsize('P')==8" >nul 2>&1
if errorlevel 1 goto try_launcher
if exist "%~dp0.tcl\tcl8.6\init.tcl" set "TCL_LIBRARY=%~dp0.tcl\tcl8.6"
if exist "%~dp0.tcl\tk8.6\tk.tcl" set "TK_LIBRARY=%~dp0.tcl\tk8.6"
"%PYCMD%" "%~1" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:try_launcher
py -3.12-64 -c "import sys; assert sys.version_info[:2]==(3,12)" >nul 2>&1
if errorlevel 1 goto try_path
py -3.12-64 "%~1" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:try_path
python -c "import struct,sys; assert sys.version_info[:2]==(3,12) and struct.calcsize('P')==8" >nul 2>&1
if errorlevel 1 goto missing
python "%~1" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:missing
echo Scaner needs 64-bit Python 3.12. See README.md.
exit /b 1
