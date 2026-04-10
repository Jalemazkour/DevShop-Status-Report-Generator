/**
 * Preload script — exposes safe IPC bridge to renderer
 * Context isolation is ON; no direct Node access in renderer
 */
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('studio', {
  validate:         (data)              => ipcRenderer.invoke('validate', data),
  generate:         (data)              => ipcRenderer.invoke('generate', data),
  openFile:         (filePath)          => ipcRenderer.invoke('open-file', filePath),
  openOutputFolder: ()                  => ipcRenderer.invoke('open-output-folder'),
  saveFile:         (srcPath, name)     => ipcRenderer.invoke('save-file', srcPath, name),
  generateScript:   (data)              => ipcRenderer.invoke('generate-script', data),
  generateRca:      (data)              => ipcRenderer.invoke('generate-rca', data),
  listOutputFiles:  ()                  => ipcRenderer.invoke('list-output-files'),
  windowMinimize:   ()                  => ipcRenderer.send('window-minimize'),
  windowMaximize:   ()                  => ipcRenderer.send('window-maximize'),
  windowClose:      ()                  => ipcRenderer.send('window-close'),
});
