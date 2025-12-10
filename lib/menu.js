const { Menu, dialog, shell, app } = require('electron');

/**
 * Show about dialog
 */
function showAboutDialog() {
  dialog.showMessageBox({
    type: 'info',
    title: 'About SheetFlow',
    message: 'SheetFlow v1.0.0',
    detail: 'Convert Excel files to CSV format with support for large files and multiple worksheets.\\n\\nBuilt with Electron and Node.js',
    buttons: ['OK']
  });
}

/**
 * Create application menu
 */
function createMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'Open Excel File...',
          accelerator: 'CmdOrCtrl+O',
          click: (menuItem, browserWindow) => {
            if (browserWindow) {
              browserWindow.webContents.send('menu-open-file');
            }
          }
        },
        {
          type: 'separator'
        },
        ...(process.platform !== 'darwin' ? [
          {
            label: 'Exit',
            accelerator: 'Ctrl+Q',
            click: () => app.quit()
          }
        ] : [])
      ]
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', role: 'undo' },
        { label: 'Redo', accelerator: 'Shift+CmdOrCtrl+Z', role: 'redo' },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'CmdOrCtrl+X', role: 'cut' },
        { label: 'Copy', accelerator: 'CmdOrCtrl+C', role: 'copy' },
        { label: 'Paste', accelerator: 'CmdOrCtrl+V', role: 'paste' },
        { label: 'Select All', accelerator: 'CmdOrCtrl+A', role: 'selectall' }
      ]
    },
    {
      label: 'View',
      submenu: [
        { label: 'Reload', accelerator: 'CmdOrCtrl+R', role: 'reload' },
        { label: 'Force Reload', accelerator: 'CmdOrCtrl+Shift+R', role: 'forceReload' },
        { label: 'Toggle Developer Tools', accelerator: process.platform === 'darwin' ? 'Alt+Cmd+I' : 'Ctrl+Shift+I', role: 'toggleDevTools' },
        { type: 'separator' },
        { label: 'Actual Size', accelerator: 'CmdOrCtrl+0', role: 'resetZoom' },
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+Plus', role: 'zoomIn' },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', role: 'zoomOut' },
        { type: 'separator' },
        { label: 'Toggle Fullscreen', accelerator: process.platform === 'darwin' ? 'Ctrl+Cmd+F' : 'F11', role: 'togglefullscreen' }
      ]
    },
    {
      label: 'Window',
      submenu: [
        { label: 'Minimize', accelerator: 'CmdOrCtrl+M', role: 'minimize' },
        { label: 'Close', accelerator: 'CmdOrCtrl+W', role: 'close' },
        ...(process.platform === 'darwin' ? [
          { type: 'separator' },
          { label: 'Bring All to Front', role: 'front' }
        ] : [])
      ]
    },
    ...(process.platform === 'darwin' ? [
      {
        label: app.getName(),
        submenu: [
          { label: 'About ' + app.getName(), click: showAboutDialog },
          { type: 'separator' },
          { label: 'Services', role: 'services', submenu: [] },
          { type: 'separator' },
          { label: 'Hide ' + app.getName(), accelerator: 'Command+H', role: 'hide' },
          { label: 'Hide Others', accelerator: 'Command+Shift+H', role: 'hideothers' },
          { label: 'Show All', role: 'unhide' },
          { type: 'separator' },
          {
            label: 'Quit',
            accelerator: 'Command+Q',
            click: () => app.quit()
          }
        ]
      }
    ] : []),
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

module.exports = {
  createMenu,
  showAboutDialog
};