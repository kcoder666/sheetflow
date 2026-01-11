const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const { setupWindowSecurity, getSecureWebPreferences } = require('./lib/security');
const { setupIpcHandlers, cleanupIpcHandlers } = require('./lib/ipcHandlers');
const { createMenu } = require('./lib/menu');
const { logger } = require('./lib/logger');

// Set app name for macOS menu bar - must be called before app is ready
app.setName('SheetFlow');

// Global variables
let mainWindow;
let viewerWindows = new Set();

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

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 700,
    minWidth: 1000,
    minHeight: 600,
    title: 'SheetFlow',
    icon: path.join(__dirname, 'app_icon.png'),
    webPreferences: getSecureWebPreferences(path.join(__dirname, 'preload.js'))
  });

  mainWindow.loadFile('index.html');

  // Set up security for the window
  setupWindowSecurity(mainWindow);

  // Quit app when main window is closed
  mainWindow.on('closed', () => {
    app.quit();
  });
}

function createViewerWindow() {
  const viewerWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    webPreferences: getSecureWebPreferences(path.join(__dirname, 'preload.js')),
    show: false,
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default'
  });

  // Load the viewer HTML
  viewerWindow.loadFile('viewer.html');

  // Set up security for the window
  setupWindowSecurity(viewerWindow);

  // Add to viewer windows set
  viewerWindows.add(viewerWindow);

  // Show window when ready
  viewerWindow.once('ready-to-show', () => {
    viewerWindow.show();
  });

  // Remove from set when closed
  viewerWindow.on('closed', () => {
    viewerWindows.delete(viewerWindow);
  });

  return viewerWindow;
}

// Initialize app when ready
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

  createWindow();
  createMenu();
  setupIpcHandlers(mainWindow, createViewerWindow);

  logger.info('SheetFlow started successfully');
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on('window-all-closed', () => {
  cleanupIpcHandlers();
  app.quit();
});