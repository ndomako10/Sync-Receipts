@echo off
:: Kill-Excel.bat
:: Force-closes all running EXCEL.EXE processes.
:: Use this when Excel has crashed or hung and is holding a workbook file locked.

:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

echo Checking for running EXCEL.EXE processes...
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I "EXCEL.EXE" >NUL
if errorlevel 1 (
    echo No EXCEL.EXE processes found.
) else (
    taskkill /F /IM EXCEL.EXE
    echo Done.
)

pause
popd
