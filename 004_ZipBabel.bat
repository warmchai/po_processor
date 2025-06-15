@echo off
setlocal enabledelayedexpansion

REM Get the directory where the script is located
set "ROOTDIR=%~dp0"
set "UPLOADDIR=%ROOTDIR%Upload to Babel"

REM Create the Upload to Babel folder if it doesn't exist
if not exist "%UPLOADDIR%" (
    mkdir "%UPLOADDIR%"
)

REM Go to the Artwork folder
cd /d "%ROOTDIR%Artwork"

REM Loop through all folders except JastinTech
for /d %%F in (*) do (
    if /i not "%%F"=="JastinTech" (
        echo Zipping %%F...
        "C:\Program Files\7-Zip\7z.exe" a "%%F.zip" "%%F\*"
    )
)

REM Move all zip files to Upload to Babel
move "*.zip" "%UPLOADDIR%"

echo Done.
pause