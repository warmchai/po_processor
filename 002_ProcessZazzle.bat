@echo off
cd /d "%~dp0"

echo === Running zazzle_processor.py === > ZazzleDebugLog.txt 2>&1
py "Excel\zazzle_processor.py" >> ZazzleDebugLog.txt 2>&1

echo All done. See ZazzleDebugLog.txt for details.
pause