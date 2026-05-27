@echo off
setlocal EnableDelayedExpansion

title CDISC CORE Validation Launcher

echo ============================================================
echo              CDISC CORE Validation Launcher
echo ============================================================
echo.

set "CORE_DIR=P:\BSP_LocalDev\Manivannan.Mathialag\zzzz_My_SAS_Files\My GitHub\cdisc-rules-engine"
set "CORE_PY=%CORE_DIR%\core.py"

set /p STANDARD=Enter CDISC Standard (Example: sdtmig): 
echo.

set /p VERSION=Enter Standard Version (Example: 3.4): 
echo.

set /p XPT_PATH=Enter XPT Folder Location: 
echo.

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set DTS=%%i

set "OUTPUT_FILE=%XPT_PATH%\CORE_Validation_Report_%DTS%.xlsx"

echo ============================================================
echo VALIDATION DETAILS
echo ============================================================
echo Standard      : %STANDARD%
echo Version       : %VERSION%
echo XPT Folder    : %XPT_PATH%
echo Output File   : %OUTPUT_FILE%
echo Output Format : XLSX
echo File Type     : XPT only
echo ============================================================
echo.

pause

if not exist "%CORE_PY%" (
    echo.
    echo ERROR: core.py not found:
    echo %CORE_PY%
    pause
    exit /b 1
)

if not exist "%XPT_PATH%" (
    echo.
    echo ERROR: XPT folder not found:
    echo %XPT_PATH%
    pause
    exit /b 1
)

cd /d "%CORE_DIR%"

echo.
echo Running CDISC CORE Validation...
echo.

python "%CORE_PY%" validate ^
-s %STANDARD% ^
-v %VERSION% ^
-d "%XPT_PATH%" ^
-ft xpt ^
-of XLSX ^
-o "%OUTPUT_FILE%"

if errorlevel 1 (
    echo.
    echo ============================================================
    echo Validation Failed
    echo ============================================================
    echo Report was NOT generated successfully.
    echo Please check the Python error above.
    echo ============================================================
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Validation Completed Successfully
echo ============================================================
echo Report Generated:
echo %OUTPUT_FILE%
echo ============================================================
echo.

pause