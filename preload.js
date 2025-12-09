const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('select-file'),
  selectOutputDir: () => ipcRenderer.invoke('select-output-dir'),
  getWorksheets: (filePath) => ipcRenderer.invoke('get-worksheets', filePath),
  convertFile: (options) => ipcRenderer.invoke('convert-file', options),
  onFileSelected: (callback) => ipcRenderer.on('file-selected', (event, filePath) => callback(filePath))
});
