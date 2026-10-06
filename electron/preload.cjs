// Мост между окном и главным процессом: только явно разрешённые функции.
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('pero', {
  openFile: () => ipcRenderer.invoke('file:open'),
  openImage: () => ipcRenderer.invoke('file:openImage'),
  print: () => ipcRenderer.invoke('doc:print'),
  exportPdf: (name) => ipcRenderer.invoke('doc:pdf', name),
  saveFile: (path, data) => ipcRenderer.invoke('file:save', path, data),
  saveFileAs: (name, data) => ipcRenderer.invoke('file:saveAs', name, data),
  confirmDiscard: () => ipcRenderer.invoke('file:confirmDiscard'),
  showError: (title, msg) => ipcRenderer.invoke('dialog:error', title, msg),
  setDirty: (v) => ipcRenderer.send('state:dirty', v),
  setTitle: (t) => ipcRenderer.send('state:title', t),
  pathForFile: (f) => { try { return webUtils.getPathForFile(f) || null; } catch { return null; } },
  edit: (cmd) => ipcRenderer.send('edit', cmd),
  closeNow: () => ipcRenderer.send('app:closeNow'),
  onMenu: (cb) => {
    const h = (_e, cmd, payload) => cb(cmd, payload);
    ipcRenderer.on('menu', h);
    return () => ipcRenderer.removeListener('menu', h);
  },
});
