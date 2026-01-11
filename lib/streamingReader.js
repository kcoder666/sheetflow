const fs = require('fs');
const path = require('path');
const readline = require('readline');
const csv = require('csv-parser');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { spawn } = require('child_process');
const { validateFilePath } = require('./security');
const { logger } = require('./logger');

/**
 * Streaming file reader for extremely large Excel and CSV files
 * Handles millions of rows with memory-efficient processing
 */
class StreamingFileReader {
  constructor() {
    this.currentFile = null;
    this.currentSheet = null;
    this.fileType = null;
    this.totalRows = 0;
    this.columns = [];
    this.pageSize = 100;
  }

  /**
   * Initialize reader with a file
   */
  async initializeFile(filePath) {
    try {
      // Validate file path
      const validatedPath = validateFilePath(filePath);
      this.currentFile = validatedPath;

      // Check file size
      const fileStats = fs.statSync(validatedPath);
      const fileSizeBytes = fileStats.size;
      const fileSizeMB = fileSizeBytes / (1024 * 1024);

      // Detect file type
      const ext = path.extname(validatedPath).toLowerCase();

      if (ext === '.csv') {
        this.fileType = 'csv';
        await this.analyzeCsvFile();
      } else if (['.xlsx', '.xls', '.xlsm'].includes(ext)) {
        this.fileType = 'excel';

        // Use Python for large Excel files (> 200MB)
        if (fileSizeMB > 200) {
          logger.info('Large Excel file detected, using Python fallback', {
            filePath: validatedPath,
            sizeMB: fileSizeMB.toFixed(2)
          });
          await this.analyzeExcelFilePython();
        } else {
          await this.analyzeExcelFile();
        }
      } else {
        throw new Error('Unsupported file type. Only CSV and Excel files are supported.');
      }

      logger.info('File initialized for streaming', {
        filePath: validatedPath,
        type: this.fileType,
        totalRows: this.totalRows,
        columns: this.columns.length
      });

      return {
        success: true,
        fileType: this.fileType,
        totalRows: this.totalRows,
        columns: this.columns,
        sheets: this.fileType === 'excel' ? await this.getExcelSheets() : null
      };

    } catch (error) {
      logger.error('Failed to initialize file for streaming', { error: error.message, filePath });
      throw error;
    }
  }

  /**
   * Analyze CSV file structure and get basic info
   */
  async analyzeCsvFile() {
    return new Promise((resolve, reject) => {
      let rowCount = 0;
      let headersFound = false;

      const stream = fs.createReadStream(this.currentFile, { encoding: 'utf8' });
      const rl = readline.createInterface({
        input: stream,
        crlfDelay: Infinity
      });

      rl.on('line', (line) => {
        rowCount++;

        // Get headers from first row
        if (!headersFound && line.trim()) {
          try {
            // Parse CSV line to get column headers
            const headers = this.parseCsvLine(line);
            this.columns = headers.map((header, index) => ({
              index,
              name: header || `Column ${index + 1}`,
              type: 'text'
            }));
            headersFound = true;
          } catch (error) {
            // If parsing fails, create generic headers
            const commaCount = (line.match(/,/g) || []).length;
            this.columns = Array.from({ length: commaCount + 1 }, (_, i) => ({
              index: i,
              name: `Column ${i + 1}`,
              type: 'text'
            }));
            headersFound = true;
          }
        }
      });

      rl.on('close', () => {
        this.totalRows = Math.max(0, rowCount - 1); // Subtract header row
        resolve();
      });

      rl.on('error', (error) => {
        reject(new Error(`Failed to analyze CSV: ${error.message}`));
      });
    });
  }

  /**
   * Analyze Excel file structure
   */
  async analyzeExcelFile() {
    try {
      // Use XLSX for quick analysis
      const workbook = XLSX.readFile(this.currentFile, {
        cellDates: false,
        cellNF: false,
        cellText: false,
        dense: false,
        sheetStubs: false
      });

      const sheetName = workbook.SheetNames[0];
      if (!sheetName) {
        throw new Error('No worksheets found in Excel file');
      }

      await this.setExcelSheet(sheetName);

    } catch (error) {
      throw new Error(`Failed to analyze Excel file: ${error.message}`);
    }
  }

  /**
   * Set current Excel sheet for reading
   */
  async setExcelSheet(sheetName) {
    try {
      // Check file size to determine method
      const fileStats = fs.statSync(this.currentFile);
      const fileSizeMB = fileStats.size / (1024 * 1024);

      if (fileSizeMB > 200) {
        // Use Python for large files
        logger.info('Using Python for large file sheet selection', {
          sizeMB: fileSizeMB.toFixed(2),
          sheetName
        });
        await this.setExcelSheetPython(sheetName);
        return;
      }

      // Use XLSX for smaller files
      const workbook = XLSX.readFile(this.currentFile);
      const worksheet = workbook.Sheets[sheetName];

      if (!worksheet) {
        throw new Error(`Worksheet '${sheetName}' not found`);
      }

      this.currentSheet = sheetName;

      // Get sheet dimensions
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      this.totalRows = Math.max(0, range.e.r - range.s.r); // Subtract header row

      // Get column headers
      this.columns = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const cell = worksheet[cellAddress];
        const header = cell ? String(cell.v) : `Column ${col + 1}`;

        this.columns.push({
          index: col,
          name: header,
          type: 'text'
        });
      }

      logger.info('Excel sheet set for streaming', {
        sheetName,
        totalRows: this.totalRows,
        columns: this.columns.length
      });

    } catch (error) {
      // Try Python fallback if XLSX fails
      try {
        logger.info('XLSX sheet selection failed, trying Python fallback', { sheetName });
        await this.setExcelSheetPython(sheetName);
      } catch (pythonError) {
        throw new Error(`Failed to set Excel sheet: ${error.message}. Python fallback also failed: ${pythonError.message}`);
      }
    }
  }

  /**
   * Get list of Excel sheets
   */
  async getExcelSheets() {
    if (this.fileType !== 'excel') return null;

    try {
      // Check file size to determine method
      const fileStats = fs.statSync(this.currentFile);
      const fileSizeMB = fileStats.size / (1024 * 1024);

      if (fileSizeMB > 200) {
        // Use Python for large files
        logger.info('Using Python for large file worksheet detection', { sizeMB: fileSizeMB.toFixed(2) });
        return await this.getExcelSheetsPython();
      } else {
        // Use XLSX for smaller files
        const workbook = XLSX.readFile(this.currentFile, { bookSheets: true });
        return workbook.SheetNames.map(name => ({
          name,
          estimated: true
        }));
      }
    } catch (error) {
      logger.error('Failed to get Excel sheets', { error: error.message });
      // Try Python fallback if XLSX fails
      try {
        logger.info('XLSX failed, trying Python fallback');
        return await this.getExcelSheetsPython();
      } catch (pythonError) {
        logger.error('Python fallback also failed', { error: pythonError.message });
        return [];
      }
    }
  }

  /**
   * Read a page of data (streaming)
   */
  async readPage(startRow = 0, pageSize = this.pageSize) {
    try {
      if (!this.currentFile) {
        throw new Error('No file initialized');
      }

      if (this.fileType === 'csv') {
        return await this.readCsvPage(startRow, pageSize);
      } else if (this.fileType === 'excel') {
        return await this.readExcelPage(startRow, pageSize);
      } else {
        throw new Error('Unsupported file type');
      }

    } catch (error) {
      logger.error('Failed to read page', { error: error.message, startRow, pageSize });
      throw error;
    }
  }

  /**
   * Read CSV page using streaming
   */
  async readCsvPage(startRow, pageSize) {
    return new Promise((resolve, reject) => {
      const rows = [];
      let currentRow = -1; // Start at -1 to account for header row
      let foundRows = 0;

      const stream = fs.createReadStream(this.currentFile, { encoding: 'utf8' });
      const rl = readline.createInterface({
        input: stream,
        crlfDelay: Infinity
      });

      rl.on('line', (line) => {
        currentRow++;

        // Skip header row
        if (currentRow === 0) return;

        // Check if we're in the target range
        const dataRow = currentRow - 1; // Adjust for header
        if (dataRow >= startRow && foundRows < pageSize) {
          try {
            const values = this.parseCsvLine(line);
            const rowData = {};

            this.columns.forEach((col, index) => {
              rowData[col.name] = values[index] || '';
            });

            rows.push({
              rowNumber: dataRow + 1,
              data: rowData
            });

            foundRows++;
          } catch (error) {
            // Skip malformed rows
            logger.warn('Skipping malformed CSV row', { rowNumber: currentRow, error: error.message });
          }
        }

        // Stop if we have enough rows
        if (foundRows >= pageSize) {
          rl.close();
        }
      });

      rl.on('close', () => {
        resolve({
          rows,
          startRow,
          endRow: startRow + foundRows - 1,
          totalRows: this.totalRows,
          hasMore: (startRow + foundRows) < this.totalRows
        });
      });

      rl.on('error', (error) => {
        reject(new Error(`Failed to read CSV page: ${error.message}`));
      });
    });
  }

  /**
   * Read Excel page using XLSX
   */
  async readExcelPage(startRow, pageSize) {
    try {
      const workbook = XLSX.readFile(this.currentFile);
      const worksheet = workbook.Sheets[this.currentSheet];

      if (!worksheet) {
        throw new Error(`Worksheet '${this.currentSheet}' not found`);
      }

      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const rows = [];

      // Read the requested range of rows
      const endRow = Math.min(startRow + pageSize, this.totalRows);

      for (let row = startRow; row < endRow; row++) {
        const actualRow = row + range.s.r + 1; // Adjust for header row and sheet offset
        const rowData = {};

        this.columns.forEach((col) => {
          const cellAddress = XLSX.utils.encode_cell({ r: actualRow, c: col.index });
          const cell = worksheet[cellAddress];
          rowData[col.name] = cell ? String(cell.v) : '';
        });

        rows.push({
          rowNumber: row + 1,
          data: rowData
        });
      }

      return {
        rows,
        startRow,
        endRow: endRow - 1,
        totalRows: this.totalRows,
        hasMore: endRow < this.totalRows
      };

    } catch (error) {
      throw new Error(`Failed to read Excel page: ${error.message}`);
    }
  }

  /**
   * Parse CSV line handling quotes and escapes
   */
  parseCsvLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    let i = 0;

    while (i < line.length) {
      const char = line[i];

      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          // Escaped quote
          current += '"';
          i += 2;
        } else {
          // Toggle quote state
          inQuotes = !inQuotes;
          i++;
        }
      } else if (char === ',' && !inQuotes) {
        // End of field
        result.push(current);
        current = '';
        i++;
      } else {
        current += char;
        i++;
      }
    }

    // Add the last field
    result.push(current);

    return result;
  }

  /**
   * Get file info
   */
  getFileInfo() {
    if (!this.currentFile) {
      return null;
    }

    try {
      const stats = fs.statSync(this.currentFile);

      return {
        filePath: this.currentFile,
        fileName: path.basename(this.currentFile),
        fileSize: stats.size,
        fileSizeMB: (stats.size / (1024 * 1024)).toFixed(2),
        fileType: this.fileType,
        currentSheet: this.currentSheet,
        totalRows: this.totalRows,
        totalColumns: this.columns.length,
        columns: this.columns
      };
    } catch (error) {
      logger.error('Failed to get file info', { error: error.message });
      return null;
    }
  }

  /**
   * Analyze Excel file using Python (for large files)
   */
  async analyzeExcelFilePython() {
    try {
      // First get worksheet names using Python
      const worksheets = await this.getExcelSheetsPython();

      if (!worksheets || worksheets.length === 0) {
        throw new Error('No worksheets found in Excel file');
      }

      // Use the first worksheet by default
      const firstSheet = worksheets[0];
      await this.setExcelSheetPython(firstSheet.name);

      logger.info('Excel file analyzed using Python', {
        sheetCount: worksheets.length,
        firstSheet: firstSheet.name
      });
    } catch (error) {
      throw new Error(`Failed to analyze Excel file with Python: ${error.message}`);
    }
  }

  /**
   * Get Excel worksheets using Python
   */
  async getExcelSheetsPython() {
    return new Promise((resolve, reject) => {
      const pythonScript = `
import pandas as pd
import json
import sys

try:
    file_path = r"${this.currentFile.replace(/\\/g, '\\\\')}"
    xl_file = pd.ExcelFile(file_path)

    sheets = []
    for sheet_name in xl_file.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
            sheets.append({
                "name": sheet_name,
                "columns": len(df.columns),
                "accessible": True
            })
        except Exception as e:
            sheets.append({
                "name": sheet_name,
                "columns": 0,
                "accessible": False,
                "error": str(e)
            })

    result = {"success": True, "sheets": sheets}
    print(json.dumps(result))

except Exception as e:
    result = {"success": False, "error": str(e)}
    print(json.dumps(result))
`;

      const python = spawn('python3', ['-c', pythonScript], {
        stdio: ['pipe', 'pipe', 'pipe']
      });

      let output = '';
      let errorOutput = '';

      python.stdout.on('data', (data) => {
        output += data.toString();
      });

      python.stderr.on('data', (data) => {
        errorOutput += data.toString();
      });

      python.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`Python process failed: ${errorOutput || 'Unknown error'}`));
          return;
        }

        try {
          const result = JSON.parse(output.trim());
          if (result.success) {
            resolve(result.sheets.map(sheet => ({
              name: sheet.name,
              accessible: sheet.accessible,
              estimated: true
            })));
          } else {
            reject(new Error(result.error));
          }
        } catch (parseError) {
          reject(new Error(`Failed to parse Python output: ${parseError.message}`));
        }
      });

      python.on('error', (error) => {
        reject(new Error(`Failed to start Python process: ${error.message}`));
      });
    });
  }

  /**
   * Set Excel sheet using Python (for large files)
   */
  async setExcelSheetPython(sheetName) {
    return new Promise((resolve, reject) => {
      const pythonScript = `
import pandas as pd
import json

try:
    file_path = r"${this.currentFile.replace(/\\/g, '\\\\')}"

    # Get basic info about the sheet
    df_info = pd.read_excel(file_path, sheet_name="${sheetName}", nrows=0)
    df_sample = pd.read_excel(file_path, sheet_name="${sheetName}", nrows=1)

    # Get total rows (this might be slow for very large files)
    try:
        df_full = pd.read_excel(file_path, sheet_name="${sheetName}")
        total_rows = len(df_full)
    except:
        # Fallback: estimate based on file size
        total_rows = 100000  # Conservative estimate

    columns = []
    for i, col in enumerate(df_info.columns):
        columns.append({
            "index": i,
            "name": str(col),
            "type": "text"
        })

    result = {
        "success": True,
        "totalRows": total_rows,
        "columns": columns,
        "sheetName": "${sheetName}"
    }
    print(json.dumps(result))

except Exception as e:
    result = {"success": False, "error": str(e)}
    print(json.dumps(result))
`;

      const python = spawn('python3', ['-c', pythonScript], {
        stdio: ['pipe', 'pipe', 'pipe']
      });

      let output = '';
      let errorOutput = '';

      python.stdout.on('data', (data) => {
        output += data.toString();
      });

      python.stderr.on('data', (data) => {
        errorOutput += data.toString();
      });

      python.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`Python process failed: ${errorOutput || 'Unknown error'}`));
          return;
        }

        try {
          const result = JSON.parse(output.trim());
          if (result.success) {
            this.currentSheet = result.sheetName;
            this.totalRows = result.totalRows;
            this.columns = result.columns;
            resolve();
          } else {
            reject(new Error(result.error));
          }
        } catch (parseError) {
          reject(new Error(`Failed to parse Python output: ${parseError.message}`));
        }
      });

      python.on('error', (error) => {
        reject(new Error(`Failed to start Python process: ${error.message}`));
      });
    });
  }

  /**
   * Close and cleanup
   */
  close() {
    this.currentFile = null;
    this.currentSheet = null;
    this.fileType = null;
    this.totalRows = 0;
    this.columns = [];
  }
}

module.exports = { StreamingFileReader };