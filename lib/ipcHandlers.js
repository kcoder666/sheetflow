const { ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { validateFilePath, validateDirectoryPath, sanitizeSheetName } = require('./security');
const { convertSingleWorksheet } = require('./converter');
const { logger } = require('./logger');
const { excelWorkerManager } = require('./excelWorker');

/**
 * Set up Excel worker progress reporting
 */
function setupExcelWorkerProgress(mainWindow) {
  // Forward progress events from worker to renderer
  excelWorkerManager.on('progress', (data) => {
    mainWindow.webContents.send('excel-progress', data);
  });

  // Forward worksheet events from worker to renderer
  excelWorkerManager.on('worksheets', (data) => {
    mainWindow.webContents.send('excel-worksheets', data);
  });

  // Forward cancellation events
  excelWorkerManager.on('cancelled', (data) => {
    mainWindow.webContents.send('excel-cancelled', data);
  });
}

/**
 * Set up all IPC handlers for the application
 */
function setupIpcHandlers(mainWindow) {
  // Set up Excel worker progress reporting
  setupExcelWorkerProgress(mainWindow);
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

    try {
      // Validate selected file
      const validatedPath = validateFilePath(result.filePaths[0]);
      return validatedPath;
    } catch (error) {
      logger.warn('File validation failed', { error: error.message, path: result.filePaths[0] });
      throw new Error(`Invalid file selection: ${error.message}`);
    }
  });

  // Handle output directory selection
  ipcMain.handle('select-output-dir', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory', 'createDirectory']
    });

    if (result.canceled) {
      return null;
    }

    try {
      // Validate selected directory
      const validatedPath = validateDirectoryPath(result.filePaths[0]);
      return validatedPath;
    } catch (error) {
      console.error('Directory validation failed:', error.message);
      throw new Error(`Invalid directory selection: ${error.message}`);
    }
  });

  // Handle getting worksheets from Excel file (using background worker)
  ipcMain.handle('get-worksheets', async (event, filePath) => {
    try {
      // Validate file path
      const validatedPath = validateFilePath(filePath);

      logger.info('Starting background Excel processing', { filePath: validatedPath });

      // Use background worker for Excel processing
      const result = await excelWorkerManager.processExcelFile(validatedPath, {
        operation: 'getWorksheets'
      });

      return {
        success: true,
        worksheets: result.worksheetInfo || result.worksheets.map(name => ({
          name: typeof name === 'string' ? name : name.name,
          accessible: true,
          complexity: typeof name === 'object' ? name.complexity : 'unknown'
        }))
      };

    } catch (error) {
      logger.error('Failed to load worksheets with background worker', {
        error: error.message,
        filePath
      });

      return {
        success: false,
        error: error.message || 'Unable to load worksheets from Excel file'
      };
    }
  });

  // Handle canceling Excel processing
  ipcMain.handle('cancel-excel-processing', async (event, workerId) => {
    try {
      if (workerId) {
        excelWorkerManager.cancelWorker(workerId);
        logger.info('Cancelled Excel processing', { workerId });
      } else {
        excelWorkerManager.cancelAllWorkers();
        logger.info('Cancelled all Excel processing');
      }
      return { success: true };
    } catch (error) {
      logger.error('Failed to cancel Excel processing', { error: error.message, workerId });
      return { success: false, error: error.message };
    }
  });

  // Handle file conversion
  ipcMain.handle('convert-file', async (event, options) => {
    try {
      const { inputFile, outputDir, maxRows, maxFileSize, worksheets } = options;

      logger.info('Starting conversion', { options });

      // Validate and sanitize input file
      const validatedInputFile = validateFilePath(inputFile);

      // Validate output directory if provided
      let validatedOutputDir = null;
      if (outputDir && outputDir.trim()) {
        validatedOutputDir = validateDirectoryPath(outputDir);
      }

      // Validate worksheets
      if (!worksheets || !Array.isArray(worksheets) || worksheets.length === 0) {
        console.error('ERROR: No worksheets specified');
        throw new Error('Please select at least one worksheet to convert');
      }

      // Validate and sanitize worksheet names
      const validatedWorksheets = worksheets.map(sheet => {
        if (typeof sheet === 'string') {
          return sanitizeSheetName(sheet);
        } else if (sheet && typeof sheet.name === 'string') {
          return sanitizeSheetName(sheet.name);
        } else {
          throw new Error('Invalid worksheet name format');
        }
      });

      logger.info('Input validation complete', {
        inputFile: validatedInputFile,
        outputDir: validatedOutputDir,
        worksheets: validatedWorksheets
      });

      const baseName = path.basename(validatedInputFile, path.extname(validatedInputFile));
      let allFilesCreated = [];
      let totalRowsAll = 0;
      const errors = [];

      // Process each worksheet
      for (const sheetName of validatedWorksheets) {
        const timer = logger.time(`Convert worksheet ${sheetName}`);
        try {
          logger.info('Processing worksheet', { sheetName });
          const result = await convertSingleWorksheet(validatedInputFile, sheetName, validatedOutputDir, maxRows, maxFileSize, baseName);
          allFilesCreated.push(...result.files);
          totalRowsAll += result.totalRows;
          timer.end(`${result.totalRows} rows, ${result.files.length} file(s) created`);
          logger.info('Worksheet conversion successful', {
            sheetName,
            totalRows: result.totalRows,
            filesCreated: result.files.length
          });
        } catch (error) {
          timer.end('failed');
          logger.error('Worksheet conversion failed', { sheetName, error: error.message });
          errors.push(`${sheetName}: ${error.message}`);
        }
      }

      if (errors.length > 0 && allFilesCreated.length === 0) {
        throw new Error(`All worksheets failed to convert:\\n${errors.join('\\n')}`);
      }

      if (errors.length > 0) {
        logger.warn('Some worksheets failed conversion', { failedWorksheets: errors });
      }

      logger.info('Conversion completed', {
        totalWorksheets: validatedWorksheets.length,
        successful: validatedWorksheets.length - errors.length,
        failed: errors.length,
        totalRows: totalRowsAll,
        totalFiles: allFilesCreated.length
      });

      return {
        success: true,
        totalRows: totalRowsAll,
        files: allFilesCreated,
        errors: errors.length > 0 ? errors : null
      };
    } catch (error) {
      logger.error('Conversion failed', { error: error.message, stack: error.stack });
      return {
        success: false,
        error: error.message
      };
    }
  });
}

/**
 * Clean up IPC handlers
 */
function cleanupIpcHandlers() {
  ipcMain.removeHandler('select-file');
  ipcMain.removeHandler('select-output-dir');
  ipcMain.removeHandler('get-worksheets');
  ipcMain.removeHandler('cancel-excel-processing');
  ipcMain.removeHandler('convert-file');

  // Shutdown Excel worker manager
  excelWorkerManager.shutdown();
}

module.exports = {
  setupIpcHandlers,
  cleanupIpcHandlers
};