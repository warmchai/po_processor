@echo off
setlocal enabledelayedexpansion

REM Set up paths
set "ROOT=%~dp0"
set "EXCEL=%ROOT%Excel"
set "ARTWORK=%ROOT%Artwork"
set "JASTINTECH=%ARTWORK%\JastinTech"

REM Temp files
set "JASTIN_TMP=%ROOT%_jastin_check.tmp"
set "ZAZZLE_TMP=%ROOT%_zazzle_check.tmp"
set "FINAL_GOOD=%ROOT%ALL ORDERS READY TO GO.txt"
set "FINAL_ERR=%ROOT%ERROR CHECK TXT FILE FOR DETAILS.txt"

REM Run Jastin checker and capture output
py "%EXCEL%\Jastin_CHECKER.py" > "%JASTIN_TMP%" 2>&1

REM Run Zazzle checker and capture output
py "%EXCEL%\Zazzle_CHECKER.py" > "%ZAZZLE_TMP%" 2>&1

REM Check for errors/warnings in Jastin checker
C:\Windows\System32\findstr.exe /I "WARNING ERROR" "%JASTIN_TMP%" >nul
set JASTIN_ERR=%ERRORLEVEL%

REM Check for errors in Zazzle checker
C:\Windows\System32\findstr.exe /I "ERROR" "%ZAZZLE_TMP%" >nul
set ZAZZLE_ERR=%ERRORLEVEL%

REM Decide output file name
set "FINAL=%FINAL_GOOD%"
if %JASTIN_ERR%==0 (
    set "FINAL=%FINAL_ERR%"
)
if %ZAZZLE_ERR%==0 (
    set "FINAL=%FINAL_ERR%"
)

REM Write errors at the top if any
if "%FINAL%"=="%FINAL_ERR%" (
    >"%FINAL%" (
        echo ERROR: CHECK TXT FILE FOR DETAILS
        echo.
        if %JASTIN_ERR%==0 (
            echo --- JastinTech Issues ---
            C:\Windows\System32\findstr.exe /I "WARNING ERROR" "%JASTIN_TMP%"
            echo.
        )
        if %ZAZZLE_ERR%==0 (
            echo --- Zazzle Issues ---
            C:\Windows\System32\findstr.exe /I "ERROR" "%ZAZZLE_TMP%"
            echo.
        )
    )
) else (
    REM Create/clear the good file
    >"%FINAL%" (
        rem nothing, just clear
    )
)

REM Write the main report (append)
>>"%FINAL%" (
    echo ALL JASTINTECH ORDERS READY TO GO
    echo.
    type "%JASTIN_TMP%"
    echo.
    type "%ZAZZLE_TMP%"
    echo.
)

REM Clean up temp files
del "%JASTIN_TMP%" 2>nul
del "%ZAZZLE_TMP%" 2>nul

echo Done! See %FINAL%
pause