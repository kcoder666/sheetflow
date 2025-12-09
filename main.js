const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { spawn } = require('child_process');
const { Readable } = require('stream');

// Set app name for macOS menu bar - must be called before app is ready
app.setName('SheetFlow');

// Try to load ExcelJS, but don't fail if it's not installed
let ExcelJS;
try {
  ExcelJS = require('exceljs');
} catch (e) {
  console.log('ExcelJS not available, will use Python fallback for large files');
}

// Auto-reload for development
if (process.env.NODE_ENV === 'development' || !app.isPackaged) {
  try {
    require('electron-reload')(__dirname, {
      electron: path.join(__dirname, 'node_modules', '.bin', 'electron'),
      hardResetMethod: 'exit'
    });
  } catch (error) {
    // electron-reload not installed, continue without it
    console.log('electron-reload not available, skipping auto-reload');
  }
}

let mainWindow;

// Function to show About dialog
function showAboutDialog() {
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'About SheetFlow',
    message: 'SheetFlow',
    detail: 'Convert Excel files to CSV with ease\n\n' +
             'Created by:\n' +
             'Name: Khoa Tran (Kcoder)\n' +
             'Email: khoatran.geek@gmail.com',
    buttons: ['OK']
  });
}

// Function to open file from menu
async function openFileFromMenu() {
  if (!mainWindow) return;
  
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] }
    ]
  });

  if (!result.canceled && result.filePaths.length > 0) {
    const filePath = result.filePaths[0];
    // Send file path to renderer to update the input field
    mainWindow.webContents.send('file-selected', filePath);
  }
}

// Create application menu
function createMenu() {
  const template = [
    ...(process.platform === 'darwin' ? [{
      label: 'SheetFlow',
      submenu: [
        {
          label: 'About SheetFlow',
          click: showAboutDialog
        },
        { type: 'separator' },
        {
          label: 'Quit',
          accelerator: 'Command+Q',
          click: () => app.quit()
        }
      ]
    }] : []),
    {
      label: 'File',
      submenu: [
        {
          label: 'Open File...',
          accelerator: process.platform === 'darwin' ? 'Command+O' : 'Ctrl+O',
          click: openFileFromMenu
        },
        ...(process.platform === 'darwin' ? [
          { type: 'separator' },
          {
            label: 'Close',
            accelerator: 'Command+W',
            click: () => {
              if (mainWindow) {
                mainWindow.close();
              }
            }
          }
        ] : []),
        ...(process.platform !== 'darwin' ? [
          { type: 'separator' },
          {
            label: 'Close',
            accelerator: 'Ctrl+W',
            click: () => {
              if (mainWindow) {
                mainWindow.close();
              }
            }
          },
          {
            label: 'Exit',
            accelerator: 'Ctrl+Q',
            click: () => app.quit()
          }
        ] : [])
      ]
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'About SheetFlow',
          click: showAboutDialog
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    title: 'SheetFlow',
    icon: path.join(__dirname, 'app_icon.png'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  // Set app icon for macOS dock
  if (process.platform === 'darwin' && app.dock) {
    const iconPath = path.join(__dirname, 'app_icon.png');
    if (fs.existsSync(iconPath)) {
      app.dock.setIcon(iconPath);
    } else {
      console.warn('App icon not found at:', iconPath);
    }
  }
  
  // Create application menu
  createMenu();
  
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Handle file selection
ipcMain.handle('select-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] }
    ]
  });

  if (result.canceled) {
    return null;
  }

  return result.filePaths[0];
});

// Handle output directory selection
ipcMain.handle('select-output-dir', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory']
  });

  if (result.canceled) {
    return null;
  }

  return result.filePaths[0];
});

// Handle getting worksheets from Excel file
ipcMain.handle('get-worksheets', async (event, filePath) => {
  try {
    if (!filePath || !fs.existsSync(filePath)) {
      throw new Error('File not found');
    }

    // Try reading with XLSX first
    let workbook;
    try {
      workbook = XLSX.readFile(filePath, {
        cellDates: false,
        cellNF: false,
        cellText: false,
        dense: false,
        sheetStubs: false
      });
    } catch (error) {
      // If XLSX fails, try to get sheet names using ExcelJS if available
      if (ExcelJS) {
        try {
          const excelWorkbook = new ExcelJS.Workbook();
          await excelWorkbook.xlsx.readFile(filePath);
          return {
            success: true,
            worksheets: excelWorkbook.worksheets.map(ws => ({
              name: ws.name,
              accessible: true
            }))
          };
        } catch (excelError) {
          throw new Error(`Failed to read Excel file: ${error.message}`);
        }
      }
      throw error;
    }

    // Get sheet names and check which are accessible
    const sheetNames = workbook.SheetNames || [];
    const availableSheets = workbook.Sheets ? Object.keys(workbook.Sheets) : [];
    
    const worksheets = sheetNames.map(name => ({
      name: name,
      accessible: availableSheets.includes(name)
    }));

    return {
      success: true,
      worksheets: worksheets
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
});

// Helper function to escape CSV cell
function escapeCSVCell(cell) {
  const cellStr = String(cell || '');
  if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
    return `"${cellStr.replace(/"/g, '""')}"`;
  }
  return cellStr;
}

// Streaming converter using ExcelJS - better for large files
async function convertSheetStreamingExcelJS(inputFile, sheetName, outputDir, maxRows, baseName) {
  if (!ExcelJS) {
    throw new Error('ExcelJS not available');
  }
  
  console.log('Starting ExcelJS streaming conversion for', sheetName);
  
  const output = outputDir || path.dirname(inputFile);
  const finalBaseName = baseName || path.basename(inputFile, path.extname(inputFile));
  
  const workbook = new ExcelJS.Workbook();
  
  console.log('Reading workbook with ExcelJS...');
  await workbook.xlsx.readFile(inputFile);
  
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" not found in workbook`);
  }
  
  console.log(`Sheet "${sheetName}" found. Starting row-by-row processing...`);
  
  let nonEmptyRowCount = 0;
  let fileIndex = 1;
  let currentFileRows = 0;
  let writeStream = null;
  const filesCreated = [];
  
  // Helper to get or create write stream
  function getWriteStream() {
    if (!writeStream || (maxRows && currentFileRows >= maxRows)) {
      if (writeStream) {
        writeStream.end();
        fileIndex++;
        currentFileRows = 0;
      }
      
      const fileName = maxRows && fileIndex > 1 
        ? `${finalBaseName}_${sheetName}_part${fileIndex}.csv`
        : `${finalBaseName}_${sheetName}.csv`;
      const filePath = path.join(output, fileName);
      
      writeStream = fs.createWriteStream(filePath);
      filesCreated.push({ path: filePath, rows: 0 });
      console.log(`Writing to: ${fileName}`);
    }
    return writeStream;
  }
  
  // Process rows one by one using ExcelJS streaming
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const rowValues = [];
    row.eachCell({ includeEmpty: true }, (cell) => {
      let value = '';
      if (cell.value !== null && cell.value !== undefined) {
        if (cell.value instanceof Date) {
          value = cell.value.toISOString();
        } else if (typeof cell.value === 'object') {
          value = cell.value.text || String(cell.value);
        } else {
          value = String(cell.value);
        }
      }
      rowValues.push(value);
    });
    
    // Check if row has content
    const hasContent = rowValues.some(cell => {
      const cellStr = String(cell || '').trim();
      return cellStr !== '' && cellStr !== 'null' && cellStr !== 'undefined';
    });
    
    if (!hasContent) {
      return; // Skip empty rows
    }
    
    nonEmptyRowCount++;
    currentFileRows++;
    
    // Convert row to CSV
    const csvRow = rowValues.map(escapeCSVCell).join(',') + '\n';
    
    // Get write stream (creates new file if needed)
    const stream = getWriteStream();
    stream.write(csvRow);
    
    // Update file row count
    const currentFile = filesCreated[filesCreated.length - 1];
    currentFile.rows = currentFileRows;
    
    // Log progress
    if (nonEmptyRowCount % 10000 === 0) {
      console.log(`Processed ${nonEmptyRowCount} non-empty rows...`);
    }
  });
  
  // Close final stream
  if (writeStream) {
    writeStream.end();
  }
  
  console.log(`ExcelJS streaming conversion complete. Total rows: ${nonEmptyRowCount}, Files: ${filesCreated.length}`);
  
  return {
    totalRows: nonEmptyRowCount,
    files: filesCreated
  };
}

// Helper function to get Python executable path (bundled or system)
function getPythonExecutable() {
  // In packaged app, Python executable is in resources folder
  if (app.isPackaged) {
    const resourcesPath = process.resourcesPath;
    const platform = process.platform;
    
    if (platform === 'win32') {
      // Windows: python_converter.exe in resources
      const exePath = path.join(resourcesPath, 'python_converter.exe');
      if (fs.existsSync(exePath)) {
        return exePath;
      }
    } else if (platform === 'darwin') {
      // macOS: python_converter executable in resources
      const exePath = path.join(resourcesPath, 'python_converter');
      if (fs.existsSync(exePath)) {
        return exePath;
      }
      // Also check for .app bundle structure
      const appPath = path.join(resourcesPath, 'python_converter.app', 'Contents', 'MacOS', 'python_converter');
      if (fs.existsSync(appPath)) {
        return appPath;
      }
    } else {
      // Linux: python_converter executable
      const exePath = path.join(resourcesPath, 'python_converter');
      if (fs.existsSync(exePath)) {
        return exePath;
      }
    }
  } else {
    // Development: try bundled script first, then system Python
    const bundledScript = path.join(__dirname, 'python_converter.py');
    if (fs.existsSync(bundledScript)) {
      // Use system Python to run the script
      const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
      return { command: pythonCmd, script: bundledScript };
    }
  }
  
  // Fallback to system Python
  return process.platform === 'win32' ? 'python' : 'python3';
}

// Python script converter (fallback when Node.js libraries fail)
async function convertSheetPython(inputFile, sheetName, outputDir, maxRows, baseName) {
  console.log('Using Python/pandas for conversion (most reliable for large files)...');
  
  // Ensure baseName is defined
  if (!baseName) {
    baseName = path.basename(inputFile, path.extname(inputFile));
  }
  
  const output = outputDir || path.dirname(inputFile);
  const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
  
  const pythonExec = getPythonExecutable();
  let pythonCmd, scriptPath;
  
  // Check if it's a development mode object
  if (typeof pythonExec === 'object' && pythonExec.command) {
    // Development mode: using system Python with script
    pythonCmd = pythonExec.command;
    scriptPath = pythonExec.script;
  } else if (typeof pythonExec === 'string' && fs.existsSync(pythonExec)) {
    // Packaged app: using bundled executable
    pythonCmd = pythonExec;
    scriptPath = null; // Executable doesn't need script path
  } else {
    // System Python fallback
    pythonCmd = pythonExec;
    // Create temp script for system Python
    const escapePythonString = (str) => str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const pythonScript = `
import pandas as pd
import sys
import os

input_file = r"${escapePythonString(inputFile)}"
sheet_name = "${sheetName}"
output_file = r"${escapePythonString(outputFile)}"
max_rows = ${maxRows || 'None'}

try:
    print(f"Reading {input_file}...")
    df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
    print(f"Read {len(df)} rows")
    
    df = df.dropna(how='all')
    print(f"After removing empty rows: {len(df)} rows")
    
    if len(df) == 0:
        print("ERROR: No data found")
        sys.exit(1)
    
    if max_rows and len(df) > max_rows:
        num_files = (len(df) + max_rows - 1) // max_rows
        print(f"Splitting into {num_files} files...")
        
        for i in range(num_files):
            start_idx = i * max_rows
            end_idx = min((i + 1) * max_rows, len(df))
            chunk = df.iloc[start_idx:end_idx]
            
            if i == 0:
                output_path = output_file
            else:
                base_name = os.path.splitext(output_file)[0]
                output_path = f"{base_name}_part{i + 1}.csv"
            
            chunk.to_csv(output_path, index=False, encoding='utf-8')
            print(f"Created {output_path} with {len(chunk)} rows")
    else:
        df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"Created {output_file} with {len(df)} rows")
    
    print(f"SUCCESS: {len(df)} rows")
    sys.exit(0)
    
except Exception as e:
    print(f"ERROR: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
`;
    scriptPath = path.join(__dirname, 'temp_convert.py');
    fs.writeFileSync(scriptPath, pythonScript);
  }
  
  return new Promise((resolve, reject) => {
    // Prepare command arguments
    const args = scriptPath 
      ? [scriptPath, inputFile, sheetName, outputFile, maxRows || 'None']
      : [inputFile, sheetName, outputFile, maxRows || 'None'];
    
    console.log(`Executing: ${pythonCmd} ${args.join(' ')}`);
    const pythonProcess = spawn(pythonCmd, args);
    
    let stdout = '';
    let stderr = '';
    
    pythonProcess.stdout.on('data', (data) => {
      const output = data.toString();
      stdout += output;
      console.log('Python:', output.trim());
    });
    
    pythonProcess.stderr.on('data', (data) => {
      const output = data.toString();
      stderr += output;
      console.error('Python error:', output.trim());
    });
    
    pythonProcess.on('close', (code) => {
      // Clean up temp script if created
      if (scriptPath && scriptPath.includes('temp_convert.py')) {
        try {
          fs.unlinkSync(scriptPath);
        } catch (e) {
          // Ignore cleanup errors
        }
      }
      
      if (code !== 0) {
        reject(new Error(`Python conversion failed: ${stderr || stdout}`));
        return;
      }
      
      // Parse output to get row count
      const successMatch = stdout.match(/SUCCESS: (\d+) rows/);
      const totalRows = successMatch ? parseInt(successMatch[1]) : 0;
      
      // Find all created files
      const filesCreated = [];
      const fileMatches = stdout.matchAll(/Created (.+?) with (\d+) rows/g);
      for (const match of fileMatches) {
        filesCreated.push({
          path: match[1],
          rows: parseInt(match[2])
        });
      }
      
      resolve({
        totalRows,
        files: filesCreated
      });
    });
    
    pythonProcess.on('error', (error) => {
      // Clean up temp script if created
      if (scriptPath && scriptPath.includes('temp_convert.py')) {
        try {
          fs.unlinkSync(scriptPath);
        } catch (e) {
          // Ignore cleanup errors
        }
      }
      
      if (error.code === 'ENOENT') {
        reject(new Error('Python not found. Please install Python 3 and pandas: pip install pandas openpyxl'));
      } else {
        reject(new Error(`Python execution failed: ${error.message}`));
      }
    });
  });
}

// Helper function to convert a single worksheet
async function convertSingleWorksheet(inputFile, sheetName, outputDir, maxRows, baseName) {
  const output = outputDir || path.dirname(inputFile);
  
  // Check if this sheet needs Python (large file handling)
  let workbook;
  try {
    workbook = XLSX.readFile(inputFile, {
      cellDates: false,
      cellNF: false,
      cellText: false,
      dense: false,
      sheetStubs: false
    });
  } catch (error) {
    throw new Error(`Failed to read workbook: ${error.message}`);
  }

  const availableSheets = workbook.Sheets ? Object.keys(workbook.Sheets) : [];
  const sheetExists = workbook.SheetNames && workbook.SheetNames.includes(sheetName);
  const sheetAccessible = availableSheets.includes(sheetName);

  // If sheet is not accessible but exists, use Python
  if (sheetExists && !sheetAccessible) {
    console.log(`Sheet "${sheetName}" is large - using Python conversion`);
    try {
      return await convertSheetPython(inputFile, sheetName, outputDir, maxRows, baseName);
    } catch (pythonError) {
      if (ExcelJS) {
        console.log('Python failed, trying ExcelJS...');
        return await convertSheetStreamingExcelJS(inputFile, sheetName, outputDir, maxRows, baseName);
      }
      throw pythonError;
    }
  }

  // Regular conversion using XLSX
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Worksheet "${sheetName}" not found`);
  }

  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  const nonEmptyRows = data.filter(row => {
    const hasContent = row.some(cell => {
      const cellStr = String(cell || '').trim();
      return cellStr !== '' && cellStr !== 'null' && cellStr !== 'undefined';
    });
    return hasContent;
  });

  if (nonEmptyRows.length === 0) {
    throw new Error(`Worksheet "${sheetName}" is empty`);
  }

  const totalRows = nonEmptyRows.length;
  let filesCreated = [];

  if (!maxRows || totalRows <= maxRows) {
    const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
    const csv = nonEmptyRows.map(row => row.map(cell => {
      const cellStr = String(cell || '');
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        return `"${cellStr.replace(/"/g, '""')}"`;
      }
      return cellStr;
    }).join(',')).join('\n');
    fs.writeFileSync(outputFile, csv);
    filesCreated.push({ path: outputFile, rows: totalRows });
  } else {
    const numFiles = Math.ceil(totalRows / maxRows);
    for (let i = 0; i < numFiles; i++) {
      const startIdx = i * maxRows;
      const endIdx = Math.min((i + 1) * maxRows, totalRows);
      const chunk = nonEmptyRows.slice(startIdx, endIdx);
      const outputFile = path.join(output, `${baseName}_${sheetName}_part${i + 1}.csv`);
      const csv = chunk.map(row => row.map(cell => {
        const cellStr = String(cell || '');
        if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
          return `"${cellStr.replace(/"/g, '""')}"`;
        }
        return cellStr;
      }).join(',')).join('\n');
      fs.writeFileSync(outputFile, csv);
      filesCreated.push({ path: outputFile, rows: chunk.length });
    }
  }

  return {
    totalRows,
    files: filesCreated
  };
}

// Handle Excel to CSV conversion
ipcMain.handle('convert-file', async (event, options) => {
  try {
    const { inputFile, outputDir, maxRows, worksheets } = options;

    console.log('=== CONVERSION START ===');
    console.log('Options received:', JSON.stringify(options, null, 2));

    // Validate input file
    if (!inputFile || !inputFile.trim()) {
      console.error('ERROR: No input file specified');
      throw new Error('No input file specified');
    }

    // Validate worksheets
    if (!worksheets || !Array.isArray(worksheets) || worksheets.length === 0) {
      console.error('ERROR: No worksheets specified');
      throw new Error('Please select at least one worksheet to convert');
    }

    console.log('Input file path:', inputFile);
    console.log('Selected worksheets:', worksheets);
    console.log('File path exists check:', fs.existsSync(inputFile));

    // Check if file exists
    if (!fs.existsSync(inputFile)) {
      console.error('ERROR: File not found:', inputFile);
      throw new Error(`File not found: ${inputFile}`);
    }

    const baseName = path.basename(inputFile, path.extname(inputFile));
    let allFilesCreated = [];
    let totalRowsAll = 0;
    const errors = [];

    // Process each worksheet
    for (const sheetName of worksheets) {
      try {
        console.log(`\n=== Processing worksheet: ${sheetName} ===`);
        const result = await convertSingleWorksheet(inputFile, sheetName, outputDir, maxRows, baseName);
        allFilesCreated.push(...result.files);
        totalRowsAll += result.totalRows;
        console.log(`✓ Successfully converted "${sheetName}": ${result.totalRows} rows, ${result.files.length} file(s)`);
      } catch (error) {
        console.error(`✗ Failed to convert "${sheetName}": ${error.message}`);
        errors.push(`${sheetName}: ${error.message}`);
      }
    }

    if (errors.length > 0 && allFilesCreated.length === 0) {
      throw new Error(`All worksheets failed to convert:\n${errors.join('\n')}`);
    }

    if (errors.length > 0) {
      console.warn(`Some worksheets failed: ${errors.join('; ')}`);
    }

    console.log('\n=== CONVERSION END ===');
    console.log(`Total worksheets processed: ${worksheets.length}`);
    console.log(`Successful: ${worksheets.length - errors.length}, Failed: ${errors.length}`);
    console.log(`Total rows: ${totalRowsAll}`);
    console.log(`Total files created: ${allFilesCreated.length}`);

    return {
      success: true,
      totalRows: totalRowsAll,
      files: allFilesCreated,
      errors: errors.length > 0 ? errors : undefined
    };

  } catch (error) {
    console.error('=== CONVERSION ERROR ===');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    console.error('=== END ERROR ===');
    return {
      success: false,
      error: error.message
    };
  }
});
