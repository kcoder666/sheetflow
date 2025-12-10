# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

SheetFlow is an Electron desktop application that converts Excel files (.xlsx, .xls, .xlsm) to CSV format. It supports large files (200MB+) by using a bundled Python executable with pandas/openpyxl. The app provides a GUI for selecting worksheets, splitting output files by row count or file size, and choosing custom output directories.

**Key characteristic**: No Python installation required for end users—Python is bundled with the packaged app.

## Commands

### Development
```bash
# Start the app in development mode
pnpm start

# Start with hot-reload
pnpm dev
```

### Building

Building follows a two-step process:

**Step 1: Build Python executable**
```bash
# On macOS/Linux
npm run build-python

# On Windows
npm run build-python-win
```

This creates a standalone Python executable (`python_converter` or `python_converter.exe`) that will be bundled with the Electron app.

**Step 2: Package the Electron app**
```bash
pnpm package
```

The packaged app will be in the `dist/` folder and includes the Python executable.

**Note**: The prepackage script automatically runs `build-python` before packaging.

### Installing Dependencies
```bash
pnpm install
```

## Architecture

### Hybrid Conversion Strategy

The app uses a **three-tier fallback approach** for Excel conversion:

1. **XLSX library** (Node.js): Default for accessible sheets, fastest but limited to files that fit in memory
2. **ExcelJS streaming** (Node.js): For large sheets when XLSX fails, processes row-by-row
3. **Python/pandas** (subprocess): Most reliable for very large files (200MB+), called as a bundled executable

The conversion logic in `main.js` automatically determines which method to use based on file accessibility and size.

### Process Architecture

The app follows Electron's standard multi-process architecture:

- **Main process** (`main.js`): Handles file system operations, Python subprocess spawning, menu creation, and conversion logic
- **Renderer process** (`renderer.js`): UI logic for file selection, worksheet selection, and displaying results
- **Preload script** (`preload.js`): Secure bridge between main and renderer via `contextBridge`

### IPC Communication

Key IPC handlers defined in `main.js`:

- `select-file`: Open file dialog for Excel files
- `select-output-dir`: Open directory picker
- `get-worksheets`: Extract worksheet names from Excel file
- `convert-file`: Main conversion handler, processes multiple worksheets and handles splitting

### Python Integration

**In development**: Uses system Python with `python_converter.py` script
**In production**: Uses bundled executable (no Python installation needed)

The `getPythonExecutable()` function in `main.js` handles path resolution for both scenarios.

Python script (`python_converter.py`) accepts arguments:
```bash
python_converter <input_file> <sheet_name> <output_file> [max_rows] [max_file_size_mb]
```

### File Splitting Logic

Both Node.js and Python converters support splitting output CSV files based on:

- **Row limit** (`maxRows`): Split after N rows
- **File size limit** (`maxFileSize` in MB): Split when file size would exceed limit
- Files are named: `{basename}_{sheetname}.csv` for first file, `{basename}_{sheetname}_part2.csv`, etc.

The splitting logic is implemented in:
- `convertSingleWorksheet()` for XLSX
- `convertSheetStreamingExcelJS()` for ExcelJS
- `python_converter.py` for Python/pandas

## Key Files

- `main.js`: Main Electron process, conversion orchestration
- `renderer.js`: Frontend UI logic
- `preload.js`: IPC bridge
- `python_converter.py`: Standalone Python script for large file conversion
- `build-python.sh` / `build-python.bat`: Scripts to build Python executable with PyInstaller
- `index.html` / `style.css`: UI structure and styling

## Packaging Configuration

The `electron-builder` configuration in `package.json` includes:

- App ID: `com.sheetflow.app`
- Product name: `SheetFlow`
- Extra resources: `python_converter.exe`, `python_converter`, `python_converter.py` (bundled in resources directory)

## Development Notes

### Python Executable Resolution

The app tries to use Python in this order:
1. Bundled Python executable (in `process.resourcesPath` for packaged app)
2. System Python with bundled `python_converter.py` script (development)
3. System Python with temporary inline script (fallback)

### Auto-reload

In development mode, `electron-reload` is used for hot-reloading when source files change.

### Empty Row Handling

All conversion methods skip completely empty rows. A row is considered empty if all cells are empty, null, or undefined.

### CSV Escaping

The `escapeCSVCell()` helper in `main.js` properly handles CSV special characters:
- Wraps cells containing commas, quotes, or newlines in double quotes
- Escapes internal quotes by doubling them (`"` → `""`)

## Building Requirements

To build the Python executable, you need:
- Python 3.x
- pip (Python package manager)
- PyInstaller (installed automatically by build script)
- pandas and openpyxl (installed automatically by build script)

**End users do not need Python** installed to run the packaged app.
