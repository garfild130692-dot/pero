// Мост между окном и главным процессом: только явно разрешённые функции.
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('pero', {
  // что открыть в окне при загрузке: { kind, file, recovery, notice, autosaveSec }
  ready: () => ipcRenderer.invoke('app:ready'),
  newWindow: (kind) => ipcRenderer.send('app:newWindow', kind),
  openFile: () => ipcRenderer.invoke('file:open'),
  openPath: (p) => ipcRenderer.invoke('file:openPath', p),
  openImage: () => ipcRenderer.invoke('file:openImage'),
  recentList: (check) => ipcRenderer.invoke('recent:list', !!check),
  recentPin: (p, v) => ipcRenderer.invoke('recent:pin', p, v),
  recentRemove: (p) => ipcRenderer.invoke('recent:remove', p),
  recoveryList: () => ipcRenderer.invoke('recovery:list'),
  settingsGet: () => ipcRenderer.invoke('settings:get'),
  clipboardRead: () => ipcRenderer.invoke('clipboard:read'),
  winList: () => ipcRenderer.invoke('win:list'),
  winFocus: (id) => ipcRenderer.send('win:focus', id),
  winArrange: () => ipcRenderer.send('win:arrange'),
  settingsSet: (k, v) => ipcRenderer.invoke('settings:set', k, v),
  recentClear: () => ipcRenderer.invoke('recent:clear'),
  print: () => ipcRenderer.invoke('doc:print'),
  userName: () => ipcRenderer.invoke('app:userName'),
  exportPdf: (name) => ipcRenderer.invoke('doc:pdf', name),
  saveFile: (path, data) => ipcRenderer.invoke('file:save', path, data),
  saveFileAs: (name, data) => ipcRenderer.invoke('file:saveAs', name, data),
  confirmDiscard: () => ipcRenderer.invoke('file:confirmDiscard'),
  showError: (title, msg) => ipcRenderer.invoke('dialog:error', title, msg),
  setDirty: (v) => ipcRenderer.send('state:dirty', v),
  setDoc: (meta) => ipcRenderer.send('state:doc', meta),
  setTitle: (t) => ipcRenderer.send('state:title', t),
  autosave: (meta, data) => ipcRenderer.invoke('autosave:write', meta, data),
  autosaveClear: () => ipcRenderer.send('autosave:clear'),
  recoveryOpen: (id) => ipcRenderer.invoke('recovery:open', id),
  recoveryDiscard: (id) => ipcRenderer.invoke('recovery:discard', id),
  logError: (msg) => ipcRenderer.send('log:error', String(msg).slice(0, 4000)),
  reportProblem: () => ipcRenderer.invoke('app:report'),
  about: () => ipcRenderer.send('app:about'),
  openLogs: () => ipcRenderer.send('app:logs'),
  pathForFile: (f) => { try { return webUtils.getPathForFile(f) || null; } catch { return null; } },
  edit: (cmd) => ipcRenderer.send('edit', cmd),
  closeNow: () => ipcRenderer.send('app:closeNow'),
  onMenu: (cb) => {
    const h = (_e, cmd, payload) => cb(cmd, payload);
    ipcRenderer.on('menu', h);
    return () => ipcRenderer.removeListener('menu', h);
  },
});
