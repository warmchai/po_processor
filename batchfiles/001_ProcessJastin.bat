@echo off
cd /d "%~dp0"

REM Run exceltocsv.py and log output
echo === Running exceltocsv.py === >> JastinDebugLog.txt 2>&1
cd /d "%~dp0Excel"
py exceltocsv.py >> "%~dp0JastinDebugLog.txt" 2>&1
cd /d "%~dp0"
echo. >> JastinDebugLog.txt 2>&1

REM Run process_jastin.py interactively (no redirection)
echo === Running process_jastin.py ===
py "Excel\process_jastin.py"
echo.

REM Run the rest and log output
echo === Running non_golf.py === >> JastinDebugLog.txt 2>&1
py "Excel\non_golf.py" >> JastinDebugLog.txt 2>&1
echo. >> JastinDebugLog.txt 2>&1

echo === Running organize_packs.py === >> JastinDebugLog.txt 2>&1
py "Excel\organize_packs.py" >> JastinDebugLog.txt 2>&1
echo. >> JastinDebugLog.txt 2>&1

echo === Running label_items.py === >> JastinDebugLog.txt 2>&1
py "Excel\label_items.py" >> JastinDebugLog.txt 2>&1
echo. >> JastinDebugLog.txt 2>&1

echo === Running Delete_A.py === >> JastinDebugLog.txt 2>&1
py "Excel\Delete_A.py" >> JastinDebugLog.txt 2>&1
echo. >> JastinDebugLog.txt 2>&1

echo All scripts finished. See JastinDebugLog.txt for details.
pause