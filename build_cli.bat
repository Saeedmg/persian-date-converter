@echo off
REM =====================================================================
REM  Builds PersianDateConverterCLI  (onedir mode — fast startup)
REM =====================================================================
cd /d "%~dp0"

echo.
echo === Cleaning ===
if exist build rmdir /s /q build
if exist "dist\PersianDateConverterCLI" rmdir /s /q "dist\PersianDateConverterCLI"

echo.
echo === Building CLI (onedir) ===
python -m PyInstaller --onedir --noconsole --clean --name PersianDateConverterCLI ^
    --icon="%~dp0icon.ico" ^
    --version-file="%~dp0version_cli.txt" ^
    --hidden-import=jdatetime ^
    --collect-all=jdatetime ^
    PersianDateConverter.py

echo.
echo Output folder:
echo   %~dp0dist\PersianDateConverterCLI\
echo Output exe:
echo   %~dp0dist\PersianDateConverterCLI\PersianDateConverterCLI.exe
pause