@echo off
REM =====================================================================
REM  Builds PersianDateConverter.exe  (GUI)
REM  Uses --windowed so no console appears.
REM =====================================================================
cd /d "%~dp0"

echo.
echo === Environment check ===
where python
python --version
python -c "import jdatetime; print('jdatetime OK')"
python -c "import docx; print('docx OK')"
python -m PyInstaller --version

echo.
echo === Cleaning previous build ===
if exist build rmdir /s /q build
if exist PersianDateConverter.spec del /q PersianDateConverter.spec
if exist "dist\PersianDateConverter.exe" del /q "dist\PersianDateConverter.exe"

echo.
echo === Building GUI executable ===
python -m PyInstaller --onefile --windowed --clean --name PersianDateConverter ^
    --icon="%~dp0icon.ico" ^
    --version-file="%~dp0version_gui.txt" ^
    --hidden-import=jdatetime ^
    --collect-all=jdatetime ^
    PersianDateConverter.py

echo.
echo === Done ===
echo Output: %~dp0dist\PersianDateConverter.exe
pause