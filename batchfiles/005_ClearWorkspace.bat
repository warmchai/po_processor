@echo off
REM filepath: c:\Users\Work\Documents\FOURTHTRY\RUN_ClearWorkspace.bat

REM Run the clear_workspace.py script using py from the Excel folder
cd /d "%~dp0Excel"
py clear_workspace.py
pause