// Главный процесс Electron: окно, меню, работа с файлами.
const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs/promises');

const DOCX_FILTER = [{ name: 'Документ Word', extensions: ['docx'] }];
let win;
let dirty = false;
let forceClose = false;

function send(cmd, payload) {
  if (win) win.webContents.send('menu', cmd, payload);
}

async function readDoc(filePath) {
  const buf = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
}

function buildMenu() {
  const template = [
    {
      label: 'Файл',
      submenu: [
        { label: 'Создать', accelerator: 'CmdOrCtrl+N', click: () => send('new') },
        { label: 'Открыть…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { type: 'separator' },
        { label: 'Сохранить', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Сохранить как…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('saveAs') },
        { type: 'separator' },
        { label: 'Печать…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { label: 'Экспорт в PDF…', click: () => send('pdf') },
        { type: 'separator' },
        { label: 'Выход', role: 'quit' },
      ],
    },
    {
      label: 'Правка',
      submenu: [
        { label: 'Отменить', role: 'undo' },
        { label: 'Повторить', role: 'redo' },
        { type: 'separator' },
        { label: 'Вырезать', role: 'cut' },
        { label: 'Копировать', role: 'copy' },
        { label: 'Вставить', role: 'paste' },
        { label: 'Выделить всё', role: 'selectAll' },
        { type: 'separator' },
        // Ctrl+F/Ctrl+H ловит само окно (как и в браузерном режиме) — здесь только подписи
        { label: 'Найти…', accelerator: 'CmdOrCtrl+F', registerAccelerator: false, click: () => send('find') },
        { label: 'Заменить…', accelerator: 'CmdOrCtrl+H', registerAccelerator: false, click: () => send('replace') },
      ],
    },
    {
      label: 'Вид',
      submenu: [
        // Масштаб страницы документа (не всего окна). Без горячих клавиш: стандартные роли zoomIn/resetZoom
        // перехватывали Ctrl+Shift+= (надстрочный) и Ctrl+0 (интервал перед абзацем). Масштаб — Ctrl+колесо.
        { label: 'Масштаб 100%', click: () => send('zoom', 0) },
        { label: 'Крупнее', click: () => send('zoom', 1) },
        { label: 'Мельче', click: () => send('zoom', -1) },
        { type: 'separator' },
        { label: 'Во весь экран', role: 'togglefullscreen' },
        { label: 'Инструменты разработчика', role: 'toggleDevTools' },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 800,
    minHeight: 500,
    title: 'Перо',
    backgroundColor: '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });
  win.webContents.session.setSpellCheckerLanguages(['ru', 'en-US']);

  if (process.env.VITE_DEV_SERVER_URL) win.loadURL(process.env.VITE_DEV_SERVER_URL);
  else win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));

  // Ссылки из документа открываем во внешнем браузере, а не в окне редактора
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) require('electron').shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  // Файл, переданный при запуске (двойной клик по .docx)
  win.webContents.once('did-finish-load', async () => {
    const arg = process.argv.slice(app.isPackaged ? 1 : 2).find((a) => /\.docx$/i.test(a));
    if (arg) {
      try { send('openData', await readDoc(path.resolve(arg))); } catch (e) { dialog.showErrorBox('Не удалось открыть файл', String(e.message || e)); }
    }
  });

  win.on('close', (e) => {
    if (!dirty || forceClose) return;
    const choice = dialog.showMessageBoxSync(win, {
      type: 'question',
      buttons: ['Сохранить', 'Не сохранять', 'Отмена'],
      defaultId: 0,
      cancelId: 2,
      title: 'Перо',
      message: 'Сохранить изменения в документе?',
    });
    if (choice === 1) return;
    e.preventDefault();
    if (choice === 0) send('saveAndClose');
  });
}

// ---------- IPC ----------
ipcMain.handle('file:open', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openFile'], filters: DOCX_FILTER });
  if (r.canceled || !r.filePaths[0]) return null;
  return readDoc(r.filePaths[0]);
});

// Печать: системный диалог; страницы режет Chromium по CSS @page (размер, поля, колонтитулы)
ipcMain.handle('doc:print', () => new Promise((resolve) => {
  win.webContents.print({ printBackground: true }, (ok, reason) => resolve({ ok, reason }));
}));

// Экспорт в PDF: тот же CSS печати, размер листа — из @page
ipcMain.handle('doc:pdf', async (_e, suggestedName) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: suggestedName || 'Документ.pdf', filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (r.canceled || !r.filePath) return null;
  const p = /\.pdf$/i.test(r.filePath) ? r.filePath : r.filePath + '.pdf';
  const data = await win.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true });
  await fs.writeFile(p, data);
  return { path: p, name: path.basename(p) };
});

// Имя пользователя Windows — автор примечаний и исправлений
ipcMain.handle('app:userName', () => {
  try { return require('os').userInfo().username || 'Пользователь'; } catch { return 'Пользователь'; }
});

// Выбор картинки для «Вставка → Рисунок»
const IMAGE_FILTER = [{ name: 'Изображения', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'] }];
ipcMain.handle('file:openImage', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openFile'], filters: IMAGE_FILTER });
  if (r.canceled || !r.filePaths[0]) return null;
  return readDoc(r.filePaths[0]);
});

ipcMain.handle('file:save', async (_e, filePath, data) => {
  await fs.writeFile(filePath, Buffer.from(data));
  return { path: filePath, name: path.basename(filePath) };
});

ipcMain.handle('file:saveAs', async (_e, suggestedName, data) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: suggestedName || 'Документ.docx', filters: DOCX_FILTER });
  if (r.canceled || !r.filePath) return null;
  const p = /\.docx$/i.test(r.filePath) ? r.filePath : r.filePath + '.docx';
  await fs.writeFile(p, Buffer.from(data));
  return { path: p, name: path.basename(p) };
});

ipcMain.handle('file:confirmDiscard', async () => {
  if (!dirty) return true;
  const choice = dialog.showMessageBoxSync(win, {
    type: 'question',
    buttons: ['Продолжить без сохранения', 'Отмена'],
    defaultId: 1,
    cancelId: 1,
    title: 'Перо',
    message: 'В документе есть несохранённые изменения. Продолжить?',
  });
  return choice === 0;
});

ipcMain.on('edit', (e, cmd) => { if (['cut', 'copy', 'paste'].includes(cmd)) e.sender[cmd](); });
ipcMain.on('state:dirty', (_e, value) => { dirty = !!value; });
ipcMain.on('state:title', (_e, title) => { if (win) win.setTitle(title); });
ipcMain.on('app:closeNow', () => { forceClose = true; if (win) win.close(); });
ipcMain.handle('dialog:error', (_e, title, msg) => dialog.showErrorBox(title, msg));

app.whenReady().then(() => {
  buildMenu();
  createWindow();
});
app.on('window-all-closed', () => app.quit());
