@echo off
REM Script to build Python executable for bundling with Electron app (Windows)

echo Building Python executable for Excel to CSV converter...

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

REM Check if pandas and openpyxl are installed
python -c "import pandas" 2>nul
if errorlevel 1 (
    echo pandas not found. Installing...
    pip install pandas openpyxl
)

REM Build the executable
echo Creating standalone executable...
pyinstaller --onefile ^
    --name python_converter ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell._writer ^
    python_converter.py

REM Copy to root directory for electron-builder
if exist "dist\python_converter.exe" (
    copy dist\python_converter.exe python_converter.exe
    echo ✓ Windows executable created: python_converter.exe
)

echo Done! The executable is ready to be bundled with Electron.

