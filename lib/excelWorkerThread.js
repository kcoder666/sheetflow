const { parentPort, workerData } = require('worker_threads');
const fs = require('fs');
const path = require('path');

// Import Excel libraries
let XLSX, ExcelJS;

try {
  XLSX = require('xlsx');
} catch (error) {
  // XLSX not available
}

try {
  ExcelJS = require('exceljs');
} catch (error) {
  // ExcelJS not available
}

/**
 * Worker thread for processing Excel files
 * Runs in background to prevent UI blocking
 */
class ExcelProcessor {
  constructor() {
    this.workerId = workerData.workerId;
    this.filePath = workerData.filePath;
    this.options = workerData.options || {};
  }

  /**
   * Send progress update to main thread
   */
  sendProgress(progress, message) {
    parentPort.postMessage({
      type: 'progress',
      workerId: this.workerId,
      data: { progress, message }
    });
  }

  /**
   * Send worksheets result to main thread
   */
  sendWorksheets(worksheets) {
    parentPort.postMessage({
      type: 'worksheets',
      workerId: this.workerId,
      data: {
        success: true,
        worksheets: worksheets.map(name => ({
          name: name,
          accessible: true,
          size: 'unknown'
        }))
      }
    });
  }

  /**
   * Send success result to main thread
   */
  sendSuccess(result) {
    parentPort.postMessage({
      type: 'success',
      workerId: this.workerId,
      data: result
    });
  }

  /**
   * Send error to main thread
   */
  sendError(error) {
    parentPort.postMessage({
      type: 'error',
      workerId: this.workerId,
      error: error.message || error
    });
  }

  /**
   * Validate file before processing
   */
  async validateFile() {
    this.sendProgress(5, 'Validating file...');

    if (!fs.existsSync(this.filePath)) {
      throw new Error('File does not exist');
    }

    const stats = fs.statSync(this.filePath);
    if (!stats.isFile()) {
      throw new Error('Path is not a file');
    }

    // Check file extension
    const ext = path.extname(this.filePath).toLowerCase();
    const allowedExtensions = ['.xlsx', '.xls', '.xlsm'];
    if (!allowedExtensions.includes(ext)) {
      throw new Error('Invalid file type. Only Excel files are supported');
    }

    const fileSizeMB = stats.size / (1024 * 1024);
    this.sendProgress(10, `File size: ${fileSizeMB.toFixed(2)} MB`);

    return { fileSizeMB, stats };
  }

  /**
   * Process Excel file to extract worksheets
   */
  async processWorksheets() {
    this.sendProgress(20, 'Reading Excel file structure...');

    // Try XLSX first (fastest for most files)
    if (XLSX) {
      try {
        this.sendProgress(30, 'Using XLSX library...');

        const workbook = XLSX.readFile(this.filePath, {
          cellDates: false,
          cellNF: false,
          cellText: false,
          dense: false,
          sheetStubs: false
        });

        this.sendProgress(60, 'Extracting worksheet information...');

        const sheets = workbook.SheetNames || [];

        this.sendProgress(80, `Found ${sheets.length} worksheets`);

        // Send worksheets to UI immediately for user interaction
        this.sendWorksheets(sheets);

        this.sendProgress(100, 'Worksheet analysis complete');

        return {
          success: true,
          method: 'XLSX',
          worksheets: sheets,
          workbook: workbook // Keep for potential future use
        };

      } catch (xlsxError) {
        this.sendProgress(40, 'XLSX failed, trying alternative method...');

        // Try ExcelJS if XLSX fails
        if (ExcelJS) {
          return await this.tryExcelJS(xlsxError);
        } else {
          throw new Error(`Unable to read Excel file with XLSX: ${xlsxError.message}`);
        }
      }
    } else if (ExcelJS) {
      return await this.tryExcelJS();
    } else {
      throw new Error('No Excel processing libraries available');
    }
  }

  /**
   * Try processing with ExcelJS
   */
  async tryExcelJS(previousError = null) {
    try {
      this.sendProgress(50, 'Using ExcelJS library...');

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.filePath);

      this.sendProgress(70, 'Extracting worksheet information...');

      const sheets = workbook.worksheets.map(ws => ws.name);

      this.sendProgress(90, `Found ${sheets.length} worksheets`);

      // Send worksheets to UI
      this.sendWorksheets(sheets);

      this.sendProgress(100, 'Worksheet analysis complete');

      return {
        success: true,
        method: 'ExcelJS',
        worksheets: sheets,
        workbook: workbook
      };

    } catch (exceljsError) {
      const errorMsg = previousError
        ? `Failed with both XLSX (${previousError.message}) and ExcelJS (${exceljsError.message})`
        : `Failed with ExcelJS: ${exceljsError.message}`;

      throw new Error(errorMsg);
    }
  }

  /**
   * Analyze worksheet complexity for better processing decisions
   */
  async analyzeWorksheets(result) {
    if (!result.workbook) {
      return result;
    }

    this.sendProgress(85, 'Analyzing worksheet complexity...');

    try {
      const worksheetInfo = [];

      if (result.method === 'XLSX') {
        // Analyze XLSX worksheets
        for (const sheetName of result.worksheets) {
          const worksheet = result.workbook.Sheets[sheetName];
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
          const estimatedRows = range.e.r - range.s.r + 1;
          const estimatedCols = range.e.c - range.s.c + 1;

          worksheetInfo.push({
            name: sheetName,
            accessible: true,
            estimatedRows,
            estimatedCols,
            complexity: estimatedRows > 10000 ? 'large' : estimatedRows > 1000 ? 'medium' : 'small'
          });
        }
      } else if (result.method === 'ExcelJS') {
        // Analyze ExcelJS worksheets
        for (const worksheet of result.workbook.worksheets) {
          const estimatedRows = worksheet.rowCount || 0;
          const estimatedCols = worksheet.columnCount || 0;

          worksheetInfo.push({
            name: worksheet.name,
            accessible: true,
            estimatedRows,
            estimatedCols,
            complexity: estimatedRows > 10000 ? 'large' : estimatedRows > 1000 ? 'medium' : 'small'
          });
        }
      }

      // Send updated worksheet information
      parentPort.postMessage({
        type: 'worksheets',
        workerId: this.workerId,
        data: {
          success: true,
          worksheets: worksheetInfo
        }
      });

      return {
        ...result,
        worksheetInfo
      };

    } catch (error) {
      // If analysis fails, just return basic info
      return result;
    }
  }

  /**
   * Convert worksheet to CSV
   */
  async convertToCSV(sheetName, outputPath, maxRows = null, maxFileSize = null) {
    try {
      this.sendProgress(10, `Starting conversion of ${sheetName}...`);

      // Load workbook if not already loaded
      if (!this.workbook) {
        this.sendProgress(20, 'Loading Excel file...');
        await this.loadWorkbook();
      }

      this.sendProgress(30, `Processing worksheet ${sheetName}...`);

      if (this.method === 'XLSX') {
        return await this.convertWithXLSX(sheetName, outputPath, maxRows, maxFileSize);
      } else if (this.method === 'ExcelJS') {
        return await this.convertWithExcelJS(sheetName, outputPath, maxRows, maxFileSize);
      } else {
        throw new Error('No conversion method available');
      }

    } catch (error) {
      throw new Error(`Failed to convert ${sheetName}: ${error.message}`);
    }
  }

  /**
   * Load workbook for conversion
   */
  async loadWorkbook() {
    if (XLSX) {
      try {
        this.sendProgress(25, 'Loading with XLSX...');
        this.workbook = XLSX.readFile(this.filePath);
        this.method = 'XLSX';
        return;
      } catch (error) {
        // Fall through to ExcelJS
      }
    }

    if (ExcelJS) {
      this.sendProgress(25, 'Loading with ExcelJS...');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(this.filePath);
      this.workbook = workbook;
      this.method = 'ExcelJS';
    } else {
      throw new Error('No Excel libraries available for conversion');
    }
  }

  /**
   * Convert using XLSX with streaming for large files
   */
  async convertWithXLSX(sheetName, outputPath, maxRows, maxFileSize) {
    const worksheet = this.workbook.Sheets[sheetName];
    if (!worksheet) {
      throw new Error(`Worksheet "${sheetName}" not found`);
    }

    this.sendProgress(40, 'Converting to CSV format...');

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (data.length === 0) {
      throw new Error('No data found in worksheet');
    }

    this.sendProgress(60, 'Processing rows...');

    // Filter out completely empty rows
    const nonEmptyRows = data.filter(row =>
      row.some(cell => cell !== null && cell !== undefined && cell !== '')
    );

    if (nonEmptyRows.length === 0) {
      throw new Error('No non-empty rows found');
    }

    this.sendProgress(80, 'Writing CSV file...');

    // Handle file splitting if limits are specified
    if (maxRows || maxFileSize) {
      return await this.writeCSVWithSplitting(nonEmptyRows, outputPath, maxRows, maxFileSize);
    } else {
      return await this.writeCSVSingle(nonEmptyRows, outputPath);
    }
  }

  /**
   * Convert using ExcelJS with streaming
   */
  async convertWithExcelJS(sheetName, outputPath, maxRows, maxFileSize) {
    const worksheet = this.workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${sheetName}" not found`);
    }

    this.sendProgress(40, 'Converting to CSV format...');

    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      const values = row.values.slice(1); // Remove undefined first element
      rows.push(values);
    });

    if (rows.length === 0) {
      throw new Error('No data found in worksheet');
    }

    this.sendProgress(60, 'Processing rows...');

    // Filter out empty rows
    const nonEmptyRows = rows.filter(row =>
      row.some(cell => cell !== null && cell !== undefined && cell !== '')
    );

    this.sendProgress(80, 'Writing CSV file...');

    // Handle file splitting if limits are specified
    if (maxRows || maxFileSize) {
      return await this.writeCSVWithSplitting(nonEmptyRows, outputPath, maxRows, maxFileSize);
    } else {
      return await this.writeCSVSingle(nonEmptyRows, outputPath);
    }
  }

  /**
   * Write CSV without splitting
   */
  async writeCSVSingle(rows, outputPath) {
    const fs = require('fs');

    this.sendProgress(90, 'Writing CSV file...');

    const csv = rows.map(row =>
      row.map(cell => {
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
          return `"${cell.replace(/"/g, '""')}"`;
        }
        return cell || '';
      }).join(',')
    ).join('\n');

    fs.writeFileSync(outputPath, csv, 'utf8');

    this.sendProgress(100, 'Conversion complete');

    return {
      totalRows: rows.length,
      files: [{
        path: outputPath,
        rows: rows.length
      }]
    };
  }

  /**
   * Write CSV with file splitting for large files
   */
  async writeCSVWithSplitting(rows, outputPath, maxRows, maxFileSize) {
    const fs = require('fs');
    const path = require('path');

    const files = [];
    let fileIndex = 1;
    let currentRows = [];
    let currentSize = 0;
    const maxSizeBytes = maxFileSize ? maxFileSize * 1024 * 1024 : null;

    this.sendProgress(85, 'Preparing file splitting...');

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const csvRow = row.map(cell => {
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
          return `"${cell.replace(/"/g, '""')}"`;
        }
        return cell || '';
      }).join(',');

      const rowSize = Buffer.byteLength(csvRow + '\n', 'utf8');

      // Check if adding this row would exceed limits
      const wouldExceedRows = maxRows && currentRows.length >= maxRows;
      const wouldExceedSize = maxSizeBytes && (currentSize + rowSize) > maxSizeBytes;

      // Write current chunk if limits would be exceeded
      if (currentRows.length > 0 && (wouldExceedRows || wouldExceedSize)) {
        const filePath = this.getChunkFilePath(outputPath, fileIndex);
        const csv = currentRows.join('\n');
        fs.writeFileSync(filePath, csv, 'utf8');

        files.push({
          path: filePath,
          rows: currentRows.length
        });

        this.sendProgress(85 + (i / rows.length) * 10, `Created file ${fileIndex} with ${currentRows.length} rows`);

        fileIndex++;
        currentRows = [];
        currentSize = 0;
      }

      currentRows.push(csvRow);
      currentSize += rowSize;
    }

    // Write remaining rows
    if (currentRows.length > 0) {
      const filePath = fileIndex === 1 ? outputPath : this.getChunkFilePath(outputPath, fileIndex);
      const csv = currentRows.join('\n');
      fs.writeFileSync(filePath, csv, 'utf8');

      files.push({
        path: filePath,
        rows: currentRows.length
      });
    }

    this.sendProgress(100, `Conversion complete - ${files.length} file(s) created`);

    return {
      totalRows: rows.length,
      files: files
    };
  }

  /**
   * Generate chunk file path
   */
  getChunkFilePath(originalPath, index) {
    const path = require('path');
    const dir = path.dirname(originalPath);
    const name = path.basename(originalPath, path.extname(originalPath));
    const ext = path.extname(originalPath);
    return path.join(dir, `${name}_part${index}${ext}`);
  }

  /**
   * Main processing function
   */
  async process() {
    try {
      const operation = this.options.operation || 'getWorksheets';

      if (operation === 'getWorksheets') {
        this.sendProgress(0, 'Starting Excel file analysis...');

        // Step 1: Validate file
        const fileInfo = await this.validateFile();

        // Step 2: Process worksheets
        const result = await this.processWorksheets();

        // Step 3: Analyze worksheet complexity
        const finalResult = await this.analyzeWorksheets(result);

        // Send final success
        this.sendSuccess({
          success: true,
          method: finalResult.method,
          worksheets: finalResult.worksheets,
          worksheetInfo: finalResult.worksheetInfo,
          fileInfo
        });

      } else if (operation === 'convertToCSV') {
        this.sendProgress(0, 'Starting CSV conversion...');

        // Validate file
        await this.validateFile();

        // Convert worksheet
        const result = await this.convertToCSV(
          this.options.sheetName,
          this.options.outputPath,
          this.options.maxRows,
          this.options.maxFileSize
        );

        // Send success
        this.sendSuccess(result);

      } else {
        throw new Error(`Unknown operation: ${operation}`);
      }

    } catch (error) {
      this.sendError(error);
    }
  }
}

// Handle messages from main thread
parentPort.on('message', async (message) => {
  if (message.type === 'start') {
    const processor = new ExcelProcessor();
    await processor.process();
  }
});

// Handle uncaught errors
process.on('uncaughtException', (error) => {
  parentPort.postMessage({
    type: 'error',
    workerId: workerData.workerId,
    error: `Uncaught exception: ${error.message}`
  });
  process.exit(1);
});

process.on('unhandledRejection', (reason) => {
  parentPort.postMessage({
    type: 'error',
    workerId: workerData.workerId,
    error: `Unhandled rejection: ${reason}`
  });
  process.exit(1);
});