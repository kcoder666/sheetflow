const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { spawn } = require('child_process');
const { Readable } = require('stream');

// Try to load ExcelJS, but don't fail if it's not installed
let ExcelJS;
try {
  ExcelJS = require('exceljs');
} catch (error) {
  console.log('ExcelJS not found - will use XLSX and Python fallback only');
  ExcelJS = null;
}

/**
 * Find available Python executable
 */
async function findPythonExecutable() {
  const candidates = [
    // Common Python 3 executables
    'python3',
    'python',
    // Windows specific
    'py -3',
    'py',
    // macOS/Linux homebrew/pyenv
    '/usr/bin/python3',
    '/usr/local/bin/python3',
    '/opt/homebrew/bin/python3',
    // Common virtual env locations
    './venv/bin/python',
    './env/bin/python',
    './.venv/bin/python'
  ];

  for (const candidate of candidates) {
    try {
      const result = await new Promise((resolve) => {
        const proc = spawn(candidate, ['--version'], { shell: true });
        let output = '';

        proc.stdout.on('data', (data) => output += data.toString());
        proc.stderr.on('data', (data) => output += data.toString());

        proc.on('close', (code) => {
          resolve({ code, output });
        });

        proc.on('error', () => {
          resolve({ code: 1, output: '' });
        });
      });

      if (result.code === 0 && result.output.includes('Python 3')) {
        console.log(`Found Python: ${candidate} - ${result.output.trim()}`);
        return candidate;
      }
    } catch (error) {
      // Continue to next candidate
    }
  }

  // Development fallback
  const devScript = path.join(__dirname, '..', 'python_converter.py');
  if (fs.existsSync(devScript)) {
    return {
      command: 'python3',
      script: devScript
    };
  }

  throw new Error('Python 3 not found. Please install Python 3 and pandas: pip install pandas openpyxl');
}

/**
 * Convert Excel to CSV using Python
 */
async function convertWithPython(inputFile, sheetName, outputFile, maxRows = null, maxFileSize = null) {
  const pythonExec = await findPythonExecutable();

  let pythonCmd, scriptPath;

  if (typeof pythonExec === 'object' && pythonExec.command) {
    // Development mode: using system Python with script
    pythonCmd = pythonExec.command;
    scriptPath = pythonExec.script;
  } else if (typeof pythonExec === 'string' && fs.existsSync(pythonExec)) {
    // Packaged app: using bundled executable
    pythonCmd = pythonExec;
    scriptPath = null; // Executable doesn't need script path
  } else {
    // System Python fallback - use secure python converter
    pythonCmd = pythonExec;
    scriptPath = path.join(__dirname, '..', 'python_converter.py');

    // Verify the secure python converter exists
    if (!fs.existsSync(scriptPath)) {
      throw new Error('Secure Python converter script not found. Please ensure python_converter.py exists.');
    }
  }

  return new Promise((resolve, reject) => {
    // Prepare command arguments
    const args = scriptPath
      ? [scriptPath, inputFile, sheetName, outputFile, maxRows || 'None', maxFileSize || 'None']
      : [inputFile, sheetName, outputFile, maxRows || 'None', maxFileSize || 'None'];

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
      // No temp script cleanup needed since we use secure python_converter.py

      if (code !== 0) {
        reject(new Error(`Python conversion failed: ${stderr || stdout}`));
        return;
      }

      // Parse output to get row count
      const successMatch = stdout.match(/SUCCESS: (\\d+) rows/);
      const totalRows = successMatch ? parseInt(successMatch[1]) : 0;

      // Find all created files
      const filesCreated = [];
      const fileMatches = stdout.matchAll(/Created (.+?) with (\\d+) rows/g);
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
      // No temp script cleanup needed since we use secure python_converter.py

      if (error.code === 'ENOENT') {
        reject(new Error('Python not found. Please install Python 3 and pandas: pip install pandas openpyxl'));
      } else {
        reject(new Error(`Python execution failed: ${error.message}`));
      }
    });
  });
}

/**
 * Convert Excel to CSV using ExcelJS
 */
async function convertWithExcelJS(inputFile, sheetName, outputFile, maxRows, maxFileSize) {
  if (!ExcelJS) {
    throw new Error('ExcelJS not available');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputFile);

  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`Worksheet "${sheetName}" not found`);
  }

  // Convert to CSV with streaming support for large files
  const stream = new Readable({ objectMode: true });
  let rowCount = 0;
  let fileIndex = 1;
  const files = [];

  stream._read = () => {};

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header

    const values = row.values.slice(1); // Remove undefined first element
    const csvRow = values.map(cell =>
      typeof cell === 'string' ? `"${cell.replace(/"/g, '""')}"` : (cell || '')
    ).join(',');

    stream.push(csvRow + '\\n');
    rowCount++;
  });

  stream.push(null); // End stream

  // Write to file(s) with size/row limits
  let outputPath = outputFile;
  if (maxRows && rowCount > maxRows) {
    // Split into multiple files
    // Implementation for file splitting would go here
  }

  const writeStream = fs.createWriteStream(outputPath);
  stream.pipe(writeStream);

  return new Promise((resolve, reject) => {
    writeStream.on('finish', () => {
      resolve({
        totalRows: rowCount,
        files: [{ path: outputPath, rows: rowCount }]
      });
    });

    writeStream.on('error', reject);
  });
}

/**
 * Convert Excel to CSV using XLSX library (for smaller files)
 */
function convertWithXLSX(inputFile, sheetName, outputFile, maxRows, maxFileSize) {
  const workbook = XLSX.readFile(inputFile);

  if (!workbook.Sheets[sheetName]) {
    throw new Error(`Worksheet "${sheetName}" not found`);
  }

  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  if (data.length === 0) {
    throw new Error('No data found in worksheet');
  }

  // Filter out completely empty rows
  const nonEmptyRows = data.filter(row =>
    row.some(cell => cell !== null && cell !== undefined && cell !== '')
  );

  if (nonEmptyRows.length === 0) {
    throw new Error('No non-empty rows found');
  }

  // Convert to CSV
  const csv = nonEmptyRows.map(row =>
    row.map(cell => {
      if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\\n'))) {
        return `"${cell.replace(/"/g, '""')}"`;
      }
      return cell || '';
    }).join(',')
  ).join('\\n');

  // Write to file
  fs.writeFileSync(outputFile, csv, 'utf8');

  return {
    totalRows: nonEmptyRows.length,
    files: [{ path: outputFile, rows: nonEmptyRows.length }]
  };
}

/**
 * Main conversion function that tries different methods
 */
async function convertSingleWorksheet(inputFile, sheetName, outputDir, maxRows, maxFileSize, baseName) {
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
    console.log(`Sheet "${sheetName}" exists but not accessible with XLSX - using Python`);
    const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
    return await convertWithPython(inputFile, sheetName, outputFile, maxRows, maxFileSize);
  }

  if (!sheetExists) {
    throw new Error(`Worksheet "${sheetName}" not found in file`);
  }

  // Try XLSX first for smaller files
  try {
    console.log(`Converting "${sheetName}" with XLSX...`);
    const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
    return convertWithXLSX(inputFile, sheetName, outputFile, maxRows, maxFileSize);
  } catch (xlsxError) {
    console.log(`XLSX conversion failed: ${xlsxError.message}`);

    // Try ExcelJS if available
    if (ExcelJS) {
      try {
        console.log(`Converting "${sheetName}" with ExcelJS...`);
        const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
        return await convertWithExcelJS(inputFile, sheetName, outputFile, maxRows, maxFileSize);
      } catch (exceljsError) {
        console.log(`ExcelJS conversion failed: ${exceljsError.message}`);
      }
    }

    // Fall back to Python
    console.log(`Converting "${sheetName}" with Python...`);
    const outputFile = path.join(output, `${baseName}_${sheetName}.csv`);
    return await convertWithPython(inputFile, sheetName, outputFile, maxRows, maxFileSize);
  }
}

module.exports = {
  findPythonExecutable,
  convertWithPython,
  convertWithExcelJS,
  convertWithXLSX,
  convertSingleWorksheet
};