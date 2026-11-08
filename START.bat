@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 "%~dp0folder_organizer.py"
if errorlevel 1 pause
exit /b
:use_python
python "%~dp0folder_organizer.py"
if errorlevel 1 pause
exit /b
