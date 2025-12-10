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
   * Main processing function
   */
  async process() {
    try {
      this.sendProgress(0, 'Starting Excel file processing...');

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