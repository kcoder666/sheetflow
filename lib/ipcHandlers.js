const { ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { validateFilePath, validateDirectoryPath, sanitizeSheetName } = require('./security');
const { convertSingleWorksheet } = require('./converter');
const { logger } = require('./logger');

/**
 * Set up all IPC handlers for the application
 */
function setupIpcHandlers(mainWindow) {
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

  // Handle getting worksheets from Excel file
  ipcMain.handle('get-worksheets', async (event, filePath) => {
    try {
      // Validate file path
      const validatedPath = validateFilePath(filePath);

      // Try reading with XLSX first
      let workbook;
      try {
        workbook = XLSX.readFile(validatedPath, {
          cellDates: false,
          cellNF: false,
          cellText: false,
          dense: false,
          sheetStubs: false
        });

        const sheets = workbook.SheetNames || [];
        logger.info('Successfully loaded worksheets', { count: sheets.length, sheets });

        // Return format expected by renderer
        return {
          success: true,
          worksheets: sheets.map(name => ({
            name: name,
            accessible: true // XLSX can read these sheets
          }))
        };
      } catch (xlsxError) {
        logger.warn('XLSX failed to read file, trying alternative methods', { error: xlsxError.message });

        // Try to get basic sheet info using fs
        // This is a fallback that might work for some corrupted files
        try {
          const stats = fs.statSync(validatedPath);
          logger.info('File stats', { size: `${(stats.size / 1024 / 1024).toFixed(2)} MB` });

          // If file is very large, suggest using Python method
          if (stats.size > 50 * 1024 * 1024) { // 50MB
            return {
              success: false,
              error: 'File is too large for worksheet preview. You can still proceed with conversion using Python.'
            };
          }

          // For smaller files, return the XLSX error
          return {
            success: false,
            error: `Unable to read Excel file: ${xlsxError.message}`
          };
        } catch (fsError) {
          return {
            success: false,
            error: `Unable to read Excel file: ${xlsxError.message}`
          };
        }
      }
    } catch (error) {
      logger.error('Failed to load worksheets', { error: error.message, filePath });
      return {
        success: false,
        error: error.message || 'Unknown error occurred while loading worksheets'
      };
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
  ipcMain.removeHandler('convert-file');
}

module.exports = {
  setupIpcHandlers,
  cleanupIpcHandlers
};