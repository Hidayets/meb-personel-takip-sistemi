const { contextBridge } = require('electron');

contextBridge.exposeInMainWorld('env', {
  // Expose the API key from the main process environment to the renderer
  API_KEY: process.env.API_KEY
});