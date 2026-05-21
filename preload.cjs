const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('env', {
  API_KEY: process.env.GEMINI_API_KEY || ''
});

contextBridge.exposeInMainWorld('windowControls', {
  minimize: () => ipcRenderer.send('window-minimize'),
  maximize: () => ipcRenderer.send('window-maximize'),
  close: () => ipcRenderer.send('window-close'),
});

// Auto Updater API
contextBridge.exposeInMainWorld('updater', {
  checkForUpdates: () => ipcRenderer.send('check-for-updates'),
  installUpdate: () => ipcRenderer.send('install-update'),
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  onUpdateStatus: (callback) => {
    const handler = (_event, data) => callback(data);
    ipcRenderer.on('update-status', handler);
    return () => ipcRenderer.removeListener('update-status', handler);
  }
});

// Backup API - veriler güncelleme/yeniden yükleme sonrası korunsun diye
// userData/backups klasörüne otomatik yedek alır, Belgeler klasörüne dışa
// aktarma ve dosya seçerek geri yükleme imkanı sağlar.
contextBridge.exposeInMainWorld('backup', {
  readAuto: () => ipcRenderer.invoke('backup:read-auto'),
  writeAuto: (data) => ipcRenderer.invoke('backup:write-auto', data),
  writeDaily: (data) => ipcRenderer.invoke('backup:write-daily', data),
  exportToDocuments: (data) => ipcRenderer.invoke('backup:export-to-documents', data),
  saveAs: (data) => ipcRenderer.invoke('backup:save-as', data),
  import: () => ipcRenderer.invoke('backup:import'),
  openFolder: () => ipcRenderer.invoke('backup:open-folder'),
  openDocumentsFolder: () => ipcRenderer.invoke('backup:open-documents-folder'),
  list: () => ipcRenderer.invoke('backup:list'),
  readFile: (filePath) => ipcRenderer.invoke('backup:read-file', filePath)
});

// Printer API - Electron'un native PDF/yazdırma desteği.
// html2pdf yerine bu kullanılır; çok daha hızlı ve güvenilir.
contextBridge.exposeInMainWorld('printer', {
  savePDF: (options) => ipcRenderer.invoke('printer:save-pdf', options || {}),
  print: (options) => ipcRenderer.invoke('printer:print', options || {})
});
