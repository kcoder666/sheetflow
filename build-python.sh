#!/bin/bash
# Script to build Python executable for bundling with Electron app

echo "Building Python executable for Excel to CSV converter..."

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    pip install pyinstaller
fi

# Check if pandas and openpyxl are installed
if ! python3 -c "import pandas" 2>/dev/null; then
    echo "pandas not found. Installing..."
    pip install pandas openpyxl
fi

# Build the executable
echo "Creating standalone executable..."
pyinstaller --onefile \
    --name python_converter \
    --hidden-import pandas \
    --hidden-import openpyxl \
    --hidden-import openpyxl.cell._writer \
    python_converter.py

# Copy to root directory for electron-builder
if [ -f "dist/python_converter" ]; then
    cp dist/python_converter python_converter
    echo "✓ macOS/Linux executable created: python_converter"
elif [ -f "dist/python_converter.exe" ]; then
    cp dist/python_converter.exe python_converter.exe
    echo "✓ Windows executable created: python_converter.exe"
fi

echo "Done! The executable is ready to be bundled with Electron."

