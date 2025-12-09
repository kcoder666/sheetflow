# Excel to CSV Converter

Desktop application to convert Excel files (.xlsx, .xls) to CSV format with optional file splitting. Handles very large Excel files using embedded Python/pandas.

## Installation

```bash
cd excel-to-csv-app
pnpm install
```

## Usage

```bash
pnpm start
```

## Features

- Convert Excel files to CSV format
- Handles very large Excel files (200MB+) using embedded Python
- Optional splitting into multiple files with configurable row limit
- Simple drag-and-drop style interface
- Choose custom output directory
- View conversion results with file paths
- **No Python installation required for end users** - Python is bundled with the app

## Building

### Step 1: Build Python Executable

First, build the Python executable that will be bundled with the app:

**On macOS/Linux:**
```bash
npm run build-python
```

**On Windows:**
```bash
npm run build-python-win
```

This will:
- Install PyInstaller if needed
- Install pandas and openpyxl if needed
- Create a standalone Python executable (`python_converter` or `python_converter.exe`)

### Step 2: Package the Electron App

After building the Python executable, package the Electron app:

```bash
pnpm package
```

The packaged app will be in the `dist` folder and will include the Python executable, so users don't need to install Python separately.

## Development

During development, the app will try to use:
1. Bundled Python executable (if available)
2. System Python with `python_converter.py` script (if Python is installed)
3. System Python with temporary script (fallback)

For production builds, the Python executable is bundled and users don't need Python installed.

## Dependencies

- Electron: Desktop application framework
- xlsx: Excel file parsing library (for smaller files)
- ExcelJS: Alternative Excel library (optional)
- Python/pandas: Bundled executable for large files (no installation needed for end users)

## Requirements for Building

To build the Python executable, you need:
- Python 3.x
- pip (Python package manager)
- PyInstaller (installed automatically by build script)
- pandas and openpyxl (installed automatically by build script)
