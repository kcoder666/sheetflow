const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { logger } = require('./logger');

/**
 * Security helpers for file validation and sanitization
 */

function validateFilePath(filePath) {
  if (!filePath || typeof filePath !== 'string') {
    throw new Error('Invalid file path');
  }

  // Normalize path to prevent path traversal
  const normalizedPath = path.resolve(filePath);

  // Check if file exists and is actually a file
  if (!fs.existsSync(normalizedPath)) {
    throw new Error('File does not exist');
  }

  if (!fs.statSync(normalizedPath).isFile()) {
    throw new Error('Path is not a file');
  }

  // Check file extension for input files
  const ext = path.extname(normalizedPath).toLowerCase();
  const allowedExtensions = ['.xlsx', '.xls', '.xlsm'];
  if (!allowedExtensions.includes(ext)) {
    throw new Error('Invalid file type. Only Excel files (.xlsx, .xls, .xlsm) are allowed');
  }

  return normalizedPath;
}

function validateDirectoryPath(dirPath) {
  if (!dirPath || typeof dirPath !== 'string') {
    throw new Error('Invalid directory path');
  }

  // Normalize path to prevent path traversal
  const normalizedPath = path.resolve(dirPath);

  // Check if directory exists and is actually a directory
  if (!fs.existsSync(normalizedPath)) {
    throw new Error('Directory does not exist');
  }

  if (!fs.statSync(normalizedPath).isDirectory()) {
    throw new Error('Path is not a directory');
  }

  // Check write permissions
  try {
    const testFile = path.join(normalizedPath, `test_${crypto.randomBytes(8).toString('hex')}.tmp`);
    fs.writeFileSync(testFile, 'test');
    fs.unlinkSync(testFile);
  } catch (error) {
    throw new Error('Directory is not writable');
  }

  return normalizedPath;
}

function sanitizeSheetName(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Invalid sheet name');
  }

  // Remove any potentially dangerous characters
  const sanitized = sheetName.replace(/[<>:"/\\|?*]/g, '').trim();

  if (!sanitized) {
    throw new Error('Sheet name cannot be empty after sanitization');
  }

  return sanitized;
}

function setupWindowSecurity(mainWindow) {
  // Security: Prevent navigation to external URLs
  mainWindow.webContents.on('will-navigate', (event, navigationUrl) => {
    const parsedUrl = new URL(navigationUrl);

    // Only allow file:// protocol for local files
    if (parsedUrl.protocol !== 'file:') {
      event.preventDefault();
      logger.security('Blocked external navigation', { url: navigationUrl, protocol: parsedUrl.protocol });
    }
  });

  // Security: Prevent new window creation
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    logger.security('Blocked new window creation', { url });
    return { action: 'deny' };
  });

  // Security: Disable web security features that could be exploited
  mainWindow.webContents.session.setPermissionRequestHandler((webContents, permission, callback) => {
    // Deny all permissions by default
    logger.security('Blocked permission request', { permission });
    callback(false);
  });
}

function getSecureWebPreferences(preloadPath) {
  return {
    nodeIntegration: false,
    contextIsolation: true,
    enableRemoteModule: false,
    allowRunningInsecureContent: false,
    experimentalFeatures: false,
    webSecurity: true,
    sandbox: false, // Required for preload script to work
    preload: preloadPath
  };
}

module.exports = {
  validateFilePath,
  validateDirectoryPath,
  sanitizeSheetName,
  setupWindowSecurity,
  getSecureWebPreferences
};