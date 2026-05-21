import { createRequire } from 'module';
const require = createRequire(import.meta.url);
// Electron preload ESM uyumluluğu için require kullanıyoruz
const { contextBridge, ipcRenderer } = require('electron') as typeof import('electron');

contextBridge.exposeInMainWorld('env', {
  API_KEY: process.env.GEMINI_API_KEY || ''
});

contextBridge.exposeInMainWorld('windowControls', {
  minimize: () => ipcRenderer.send('window-minimize'),
  maximize: () => ipcRenderer.send('window-maximize'),
  close: () => ipcRenderer.send('window-close')
});

// Auto Updater API
contextBridge.exposeInMainWorld('updater', {
  checkForUpdates: () => ipcRenderer.send('check-for-updates'),
  installUpdate: () => ipcRenderer.send('install-update'),
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  onUpdateStatus: (callback: (data: { status: string; message: string; version?: string; percent?: number }) => void) => {
    ipcRenderer.on('update-status', (_event, data) => callback(data));
    return () => ipcRenderer.removeAllListeners('update-status');
  }
});