const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // File operations
  selectFile: () => ipcRenderer.invoke('select-file'),
  selectOutputDir: () => ipcRenderer.invoke('select-output-dir'),
  getWorksheets: (filePath) => ipcRenderer.invoke('get-worksheets', filePath),
  convertFile: (options) => ipcRenderer.invoke('convert-file', options),

  // Progress and cancellation
  cancelExcelProcessing: (workerId) => ipcRenderer.invoke('cancel-excel-processing', workerId),

  // Event listeners for progress updates
  onFileSelected: (callback) => ipcRenderer.on('file-selected', (event, filePath) => callback(filePath)),
  onExcelProgress: (callback) => ipcRenderer.on('excel-progress', (event, data) => callback(data)),
  onExcelWorksheets: (callback) => ipcRenderer.on('excel-worksheets', (event, data) => callback(data)),
  onExcelCancelled: (callback) => ipcRenderer.on('excel-cancelled', (event, data) => callback(data)),
  onConversionProgress: (callback) => ipcRenderer.on('conversion-progress', (event, data) => callback(data)),

  // Cleanup functions for event listeners
  removeExcelProgressListener: () => ipcRenderer.removeAllListeners('excel-progress'),
  removeExcelWorksheetsListener: () => ipcRenderer.removeAllListeners('excel-worksheets'),
  removeExcelCancelledListener: () => ipcRenderer.removeAllListeners('excel-cancelled'),
  removeConversionProgressListener: () => ipcRenderer.removeAllListeners('conversion-progress')
});
