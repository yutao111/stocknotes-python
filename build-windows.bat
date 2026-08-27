@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python launcher "py" was not found.
    echo Install 64-bit Python 3.11 on this build computer first.
    pause
    exit /b 1
)

if not exist ".build-venv\Scripts\python.exe" (
    py -3.11 -m venv .build-venv
    if errorlevel 1 goto :failed
)

call ".build-venv\Scripts\activate.bat"
python -m pip install --upgrade pip
if errorlevel 1 goto :failed
python -m pip install -r requirements-build.txt
if errorlevel 1 goto :failed

python -m PyInstaller --noconfirm --clean StockNotes.spec
if errorlevel 1 goto :failed

if exist "stocknotes.db" copy /Y "stocknotes.db" "dist\StockNotes\stocknotes.db" >nul
copy /Y "README-WINDOWS.txt" "dist\StockNotes\README-WINDOWS.txt" >nul

if exist "dist\StockNotes-Windows.zip" del /Q "dist\StockNotes-Windows.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\StockNotes' -DestinationPath 'dist\StockNotes-Windows.zip' -CompressionLevel Optimal"
if errorlevel 1 goto :failed

echo.
echo Build completed: dist\StockNotes-Windows.zip
echo Extract the ZIP on Windows, then double-click StockNotes.exe.
pause
exit /b 0

:failed
echo.
echo [ERROR] Build failed. Review the messages above.
pause
exit /b 1
