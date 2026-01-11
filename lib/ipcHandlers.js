const { ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { validateFilePath, validateDirectoryPath, sanitizeSheetName } = require('./security');
const { convertSingleWorksheet } = require('./converter');
const { logger } = require('./logger');
const { excelWorkerManager } = require('./excelWorker');
const { StreamingFileReader } = require('./streamingReader');

// Global streaming reader instance
let streamingReader = null;

// Reference to main window and viewer creation function
let mainWindowRef = null;
let createViewerWindowFunc = null;

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
function setupIpcHandlers(mainWindow, createViewerWindow) {
  logger.info('Setting up IPC handlers...');

  // Store references
  mainWindowRef = mainWindow;
  createViewerWindowFunc = createViewerWindow;

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

  // Handle file conversion (using background workers)
  ipcMain.handle('convert-file', async (event, options) => {
    try {
      const { inputFile, outputDir, maxRows, maxFileSize, worksheets } = options;

      logger.info('Starting background conversion', { options });

      // Validate and sanitize input file
      const validatedInputFile = validateFilePath(inputFile);

      // Validate output directory if provided
      let validatedOutputDir = null;
      if (outputDir && outputDir.trim()) {
        validatedOutputDir = validateDirectoryPath(outputDir);
      } else {
        validatedOutputDir = path.dirname(validatedInputFile);
      }

      // Validate worksheets
      if (!worksheets || !Array.isArray(worksheets) || worksheets.length === 0) {
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

      // Process each worksheet using background workers
      for (const sheetName of validatedWorksheets) {
        const timer = logger.time(`Convert worksheet ${sheetName} (background)`);
        try {
          logger.info('Starting background conversion for worksheet', { sheetName });

          // Generate output path for this worksheet
          const outputPath = path.join(validatedOutputDir, `${baseName}_${sheetName}.csv`);

          // Convert using background worker
          const result = await excelWorkerManager.convertWorksheetToCSV(
            validatedInputFile,
            sheetName,
            outputPath,
            { maxRows, maxFileSize }
          );

          allFilesCreated.push(...result.files);
          totalRowsAll += result.totalRows;

          timer.end(`${result.totalRows} rows, ${result.files.length} file(s) created`);
          logger.info('Background worksheet conversion successful', {
            sheetName,
            totalRows: result.totalRows,
            filesCreated: result.files.length
          });

          // Send progress update to renderer
          mainWindow.webContents.send('conversion-progress', {
            currentSheet: sheetName,
            completed: allFilesCreated.length,
            total: validatedWorksheets.length,
            totalRows: totalRowsAll
          });

        } catch (error) {
          timer.end('failed');
          logger.error('Background worksheet conversion failed', { sheetName, error: error.message });
          errors.push(`${sheetName}: ${error.message}`);
        }
      }

      if (errors.length > 0 && allFilesCreated.length === 0) {
        throw new Error(`All worksheets failed to convert:\\n${errors.join('\\n')}`);
      }

      if (errors.length > 0) {
        logger.warn('Some worksheets failed conversion', { failedWorksheets: errors });
      }

      logger.info('Background conversion completed', {
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
      logger.error('Background conversion failed', { error: error.message, stack: error.stack });
      return {
        success: false,
        error: error.message
      };
    }
  });

  // File viewer handlers

  // Handle file selection for viewer (supports Excel and CSV)
  ipcMain.handle('select-viewer-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      filters: [
        { name: 'All Supported Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
        { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] },
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (result.canceled) {
      return null;
    }

    try {
      // Validate selected file (allow both Excel and CSV)
      const filePath = result.filePaths[0];
      const ext = path.extname(filePath).toLowerCase();
      const allowedExtensions = ['.xlsx', '.xls', '.xlsm', '.csv'];

      if (!allowedExtensions.includes(ext)) {
        throw new Error('Invalid file type. Only Excel (.xlsx, .xls, .xlsm) and CSV (.csv) files are supported.');
      }

      // Basic file validation
      if (!fs.existsSync(filePath)) {
        throw new Error('File does not exist');
      }

      if (!fs.statSync(filePath).isFile()) {
        throw new Error('Path is not a file');
      }

      return filePath;
    } catch (error) {
      logger.warn('Viewer file validation failed', { error: error.message, path: result.filePaths[0] });
      throw new Error(`Invalid file selection: ${error.message}`);
    }
  });

  // Handle multiple file selection for viewer
  ipcMain.handle('select-multiple-viewer-files', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: 'All Supported Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
        { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] },
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });

    if (result.canceled) {
      return null;
    }

    try {
      // Validate selected files
      const validatedPaths = [];
      for (const filePath of result.filePaths) {
        const ext = path.extname(filePath).toLowerCase();
        const allowedExtensions = ['.xlsx', '.xls', '.xlsm', '.csv'];

        if (!allowedExtensions.includes(ext)) {
          continue; // Skip invalid files
        }

        // Basic file validation
        if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
          validatedPaths.push(filePath);
        }
      }

      return validatedPaths;
    } catch (error) {
      logger.warn('Multiple viewer files validation failed', { error: error.message });
      throw new Error(`Invalid file selection: ${error.message}`);
    }
  });

  // Open new viewer window
  ipcMain.handle('open-new-viewer', async () => {
    try {
      if (createViewerWindowFunc) {
        const viewerWindow = createViewerWindowFunc();
        logger.info('New viewer window opened');
        return { success: true };
      } else {
        throw new Error('Viewer window creation function not available');
      }
    } catch (error) {
      logger.error('Failed to open new viewer window', { error: error.message });
      return { success: false, error: error.message };
    }
  });

  // Open new viewer window with file selection
  logger.info('Setting up open-viewer-with-file handler');
  ipcMain.handle('open-viewer-with-file', async () => {
    try {
      // First, let user select a file
      const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [
          { name: 'All Supported Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
          { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] },
          { name: 'CSV Files', extensions: ['csv'] },
          { name: 'All Files', extensions: ['*'] }
        ]
      });

      if (result.canceled) {
        return { success: false, cancelled: true };
      }

      const filePath = result.filePaths[0];

      // Validate the selected file
      const ext = path.extname(filePath).toLowerCase();
      const allowedExtensions = ['.xlsx', '.xls', '.xlsm', '.csv'];

      if (!allowedExtensions.includes(ext)) {
        throw new Error('Invalid file type. Only Excel (.xlsx, .xls, .xlsm) and CSV (.csv) files are supported.');
      }

      if (!fs.existsSync(filePath)) {
        throw new Error('File does not exist');
      }

      if (!fs.statSync(filePath).isFile()) {
        throw new Error('Path is not a file');
      }

      // Create new viewer window
      if (createViewerWindowFunc) {
        const viewerWindow = createViewerWindowFunc();

        // Send the file path to the viewer window once it's ready
        viewerWindow.webContents.once('did-finish-load', () => {
          viewerWindow.webContents.send('open-file-directly', filePath);
        });

        logger.info('New viewer window opened with file', { filePath });
        return { success: true, filePath };
      } else {
        throw new Error('Viewer window creation function not available');
      }
    } catch (error) {
      logger.error('Failed to open viewer with file', { error: error.message });
      return { success: false, error: error.message };
    }
  });

  // Initialize file for viewing
  ipcMain.handle('viewer-init-file', async (event, filePath) => {
    try {
      const validatedPath = validateFilePath(filePath);

      // Create new streaming reader instance
      if (streamingReader) {
        streamingReader.close();
      }
      streamingReader = new StreamingFileReader();

      logger.info('Initializing file for viewing', { filePath: validatedPath });

      const result = await streamingReader.initializeFile(validatedPath);
      const fileInfo = streamingReader.getFileInfo();

      return {
        success: true,
        fileInfo,
        ...result
      };

    } catch (error) {
      logger.error('Failed to initialize file for viewing', { error: error.message, filePath });
      return {
        success: false,
        error: error.message
      };
    }
  });

  // Set Excel worksheet for viewing
  ipcMain.handle('viewer-set-sheet', async (event, sheetName) => {
    try {
      if (!streamingReader) {
        throw new Error('No file initialized for viewing');
      }

      await streamingReader.setExcelSheet(sheetName);
      const fileInfo = streamingReader.getFileInfo();

      logger.info('Set Excel sheet for viewing', { sheetName });

      return {
        success: true,
        fileInfo
      };

    } catch (error) {
      logger.error('Failed to set Excel sheet for viewing', { error: error.message, sheetName });
      return {
        success: false,
        error: error.message
      };
    }
  });

  // Read page of data
  ipcMain.handle('viewer-read-page', async (event, startRow, pageSize) => {
    try {
      if (!streamingReader) {
        throw new Error('No file initialized for viewing');
      }

      logger.info('Reading page for viewer', { startRow, pageSize });

      const result = await streamingReader.readPage(startRow, pageSize || 100);

      return {
        success: true,
        ...result
      };

    } catch (error) {
      logger.error('Failed to read page for viewer', { error: error.message, startRow, pageSize });
      return {
        success: false,
        error: error.message
      };
    }
  });

  // Get file info
  ipcMain.handle('viewer-get-file-info', async (event) => {
    try {
      if (!streamingReader) {
        return {
          success: false,
          error: 'No file initialized for viewing'
        };
      }

      const fileInfo = streamingReader.getFileInfo();

      return {
        success: true,
        fileInfo
      };

    } catch (error) {
      logger.error('Failed to get file info for viewer', { error: error.message });
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

  // File viewer handlers
  ipcMain.removeHandler('select-viewer-file');
  ipcMain.removeHandler('select-multiple-viewer-files');
  ipcMain.removeHandler('open-new-viewer');
  ipcMain.removeHandler('open-viewer-with-file');
  ipcMain.removeHandler('viewer-init-file');
  ipcMain.removeHandler('viewer-set-sheet');
  ipcMain.removeHandler('viewer-read-page');
  ipcMain.removeHandler('viewer-get-file-info');

  // Shutdown Excel worker manager
  excelWorkerManager.shutdown();

  // Close streaming reader
  if (streamingReader) {
    streamingReader.close();
    streamingReader = null;
  }
}

module.exports = {
  setupIpcHandlers,
  cleanupIpcHandlers
};