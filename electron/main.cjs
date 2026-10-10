// Главный процесс Electron: окна, меню, работа с файлами, автосохранение, журнал.
//
// Один процесс на все окна (как Word: каждый файл — своё окно, процесс один). Повторный запуск программы
// (двойной клик по файлу, ярлык) передаёт аргументы первому экземпляру и завершается.
// Ключ --new-instance запускает независимый процесс (отладка, особые случаи).
const { app, BrowserWindow, Menu, dialog, ipcMain, shell, clipboard } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs/promises');
const { createLog, stamp } = require('./log.cjs');
const { RecentFiles, samePath } = require('./recent.cjs');
const { createAutosave } = require('./autosave.cjs');

const DOCX_FILTER = [{ name: 'Документ Word', extensions: ['docx'] }];
const XLSX_FILTER = [{ name: 'Книга Excel', extensions: ['xlsx'] }];
const CSV_FILTER = [{ name: 'CSV (разделитель — точка с запятой)', extensions: ['csv'] }];
const OPEN_FILTER = [{ name: 'Документы и книги', extensions: ['docx', 'xlsx', 'csv'] }, ...DOCX_FILTER, ...XLSX_FILTER, ...CSV_FILTER];
const FILE_RE = /\.(docx|xlsx|csv)$/i;
const NEW_INSTANCE = process.argv.includes('--new-instance');
const AUTOSAVE_SEC = Math.max(1, +process.env.PERO_AUTOSAVE_SEC || 60);

// Папка настроек: у разработческого запуска — своя, чтобы не мешать установленной программе
// (блокировка единственного экземпляра привязана к этой папке)
if (process.env.PERO_USER_DATA) app.setPath('userData', process.env.PERO_USER_DATA);
else if (!app.isPackaged) app.setPath('userData', `${app.getPath('userData')}-dev`);

const userData = app.getPath('userData');
const log = createLog(path.join(userData, 'logs'));
const recent = new RecentFiles(path.join(userData, 'recent.json'));
const autosave = createAutosave(path.join(userData, 'autosave'));

// Состояние окон: webContents.id → { win, kind, path, dirty, forceClose, initial, key, crashes }
const wins = new Map();
const stateOf = (e) => wins.get(e.sender.id);
const pristine = (s) => s && !s.path && !s.dirty;
let recoveryOffered = false;

const send = (s, cmd, payload) => { if (s && !s.win.isDestroyed()) s.win.webContents.send('menu', cmd, payload); };
const stateOfWindow = (bw) => (bw ? wins.get(bw.webContents.id) : null);

async function readFile(filePath) {
  const buf = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
}

// Документ или книга пользователя: читаем и запоминаем в «Последних»
async function readUserFile(filePath) {
  const f = await readFile(filePath);
  rememberRecent(filePath);
  log.info(`Открыт файл: ${filePath}`);
  return f;
}

function rememberRecent(p) {
  recent.add(p);
  try { app.addRecentDocument(p); } catch { /* список переходов Windows — по возможности */ }
  buildMenu();
}

// Файлы из командной строки (первый запуск или повторный — second-instance)
const fileArgs = (argv, cwd) => argv.filter((a) => FILE_RE.test(a) && !a.startsWith('-')).map((a) => path.resolve(cwd || process.cwd(), a));

// Открыть файл по пути. from — окно, откуда пришёл запрос (renderer ждёт данные, если файл откроется в нём).
// Возвращает данные файла, если открывать в окне from; null — если открыто в другом окне или уже было открыто.
async function openPath(filePath, from = null) {
  const already = [...wins.values()].find((w) => w.path && samePath(w.path, filePath));
  if (already) { focus(already.win); return null; }
  let file;
  try {
    file = await readUserFile(filePath);
  } catch (e) {
    log.error(`Не удалось открыть ${filePath}: ${e.message}`);
    recent.remove(filePath);
    buildMenu();
    dialog.showErrorBox('Не удалось открыть файл', `${filePath}\n\n${e.message || e}`);
    return null;
  }
  if (from && pristine(from)) return file;
  if (!from) {
    // запуск с файлом извне: пустое нетронутое окно используем, как Word «Документ1»
    const target = [...wins.values()].find(pristine);
    if (target) {
      target.path = file.path;
      // окно ещё грузится — отдадим файл в app:ready, иначе сообщение потеряется
      if (target.ready) send(target, 'openData', file); else target.initial = { kind: target.kind, file };
      focus(target.win);
      return null;
    }
  }
  createWindow({ file });
  return null;
}

function focus(win) {
  if (win.isMinimized()) win.restore();
  win.focus();
}

// ---------- Меню ----------
function buildMenu() {
  const to = (cmd, payload) => (_item, bw) => send(stateOfWindow(bw), cmd, payload);
  const recentItems = recent.list().map((r, i) => ({
    label: `${i + 1 < 10 ? `&${i + 1} ` : ''}${r.name.replace(/&/g, '&&')}`,
    sublabel: path.dirname(r.path),
    click: () => { openPath(r.path); },
  }));
  const template = [
    {
      label: 'Файл',
      submenu: [
        { label: 'Создать документ', accelerator: 'CmdOrCtrl+N', click: () => createWindow({ kind: 'doc' }) },
        { label: 'Создать книгу', click: () => createWindow({ kind: 'sheet' }) },
        { label: 'Открыть…', accelerator: 'CmdOrCtrl+O', click: to('open') },
        {
          label: 'Последние файлы',
          submenu: recentItems.length
            ? [...recentItems, { type: 'separator' }, { label: 'Очистить список', click: () => { recent.clear(); try { app.clearRecentDocuments(); } catch { /* */ } buildMenu(); } }]
            : [{ label: 'Список пуст', enabled: false }],
        },
        { type: 'separator' },
        { label: 'Сохранить', accelerator: 'CmdOrCtrl+S', click: to('save') },
        { label: 'Сохранить как…', accelerator: 'CmdOrCtrl+Shift+S', click: to('saveAs') },
        { type: 'separator' },
        { label: 'Печать…', accelerator: 'CmdOrCtrl+P', click: to('print') },
        { label: 'Экспорт в PDF…', click: to('pdf') },
        { type: 'separator' },
        { label: 'Закрыть окно', accelerator: 'CmdOrCtrl+W', role: 'close' },
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
        { label: 'Найти…', accelerator: 'CmdOrCtrl+F', registerAccelerator: false, click: to('find') },
        { label: 'Заменить…', accelerator: 'CmdOrCtrl+H', registerAccelerator: false, click: to('replace') },
      ],
    },
    {
      label: 'Вид',
      submenu: [
        // Масштаб страницы документа (не всего окна). Без горячих клавиш: стандартные роли zoomIn/resetZoom
        // перехватывали Ctrl+Shift+= (надстрочный) и Ctrl+0 (интервал перед абзацем). Масштаб — Ctrl+колесо.
        { label: 'Масштаб 100%', click: to('zoom', 0) },
        { label: 'Крупнее', click: to('zoom', 1) },
        { label: 'Мельче', click: to('zoom', -1) },
        { type: 'separator' },
        { label: 'Во весь экран', role: 'togglefullscreen' },
        { label: 'Инструменты разработчика', role: 'toggleDevTools' },
      ],
    },
    {
      label: 'Справка',
      submenu: [
        { label: 'Сообщить о проблеме…', click: (_i, bw) => reportProblem(bw) },
        { label: 'Открыть папку журналов', click: () => shell.openPath(log.dir) },
        { type: 'separator' },
        { label: 'О программе', click: (_i, bw) => about(bw) },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function about(bw) {
  dialog.showMessageBox(bw, {
    type: 'info',
    title: 'О программе',
    message: `Перо ${app.getVersion()}`,
    detail: `Редактор документов .docx и книг .xlsx для внутреннего использования.\n\n` +
      `Electron ${process.versions.electron}, Chromium ${process.versions.chrome}\nWindows ${os.release()} (${os.arch()})\n\n© Перо`,
  });
}

// «Сообщить о проблеме»: версия, система и хвост журнала → файл в папке журналов + буфер обмена.
// Никуда не отправляется: пользователь сам вставляет текст в письмо администратору.
async function reportProblem(bw) {
  const info = [
    `Перо ${app.getVersion()} (${app.isPackaged ? 'установленная' : 'разработческая'} сборка)`,
    `Electron ${process.versions.electron}, Chromium ${process.versions.chrome}, Node ${process.versions.node}`,
    `Windows ${os.release()} ${os.arch()}, память ${Math.round(os.totalmem() / 2 ** 30)} ГБ`,
    `Время: ${stamp()}`,
    `Открытых окон: ${wins.size}`,
  ];
  const text = `${info.join('\n')}\n\n--- Журнал (последние записи) ---\n${log.tail(300).join('\n')}\n`;
  const file = path.join(log.dir, `report-${stamp().replace(/[: ]/g, '-')}.txt`);
  await fs.writeFile(file, text, 'utf8');
  clipboard.writeText(text);
  log.info(`Собран отчёт о проблеме: ${file}`);
  const r = await dialog.showMessageBox(bw, {
    type: 'info',
    title: 'Сообщить о проблеме',
    message: 'Сведения о проблеме собраны',
    detail: `Они скопированы в буфер обмена и сохранены в файл:\n${file}\n\n` +
      'Вставьте их в письмо администратору (Ctrl+V) и опишите, что вы делали, когда возникла проблема.',
    buttons: ['Показать файл', 'ОК'],
    defaultId: 1,
    cancelId: 1,
  });
  if (r.response === 0) shell.showItemInFolder(file);
  return file;
}

// ---------- Окна ----------
// kind: 'doc' | 'sheet' — что создать в пустом окне; file — открыть файл; recovered — данные восстановления
function createWindow({ kind = 'doc', file = null } = {}) {
  const ref = BrowserWindow.getFocusedWindow() || [...wins.values()].map((s) => s.win).pop();
  const b = ref && !ref.isDestroyed() && !ref.isMaximized() ? ref.getBounds() : null;
  const win = new BrowserWindow({
    width: b ? b.width : 1280,
    height: b ? b.height : 860,
    ...(b ? { x: b.x + 26, y: b.y + 26 } : {}), // каскадом от текущего, как новые окна Word
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
  if (ref && ref.isMaximized()) win.maximize();
  win.webContents.session.setSpellCheckerLanguages(['ru', 'en-US']);
  const id = win.webContents.id;
  const s = {
    win,
    kind: file ? (/\.(xlsx|csv)$/i.test(file.name) ? 'sheet' : 'doc') : kind,
    path: file && file.path ? file.path : null,
    dirty: false,
    forceClose: false,
    initial: { kind, file },
    key: `${process.pid}-${win.id}`,
    crashes: [],
  };
  wins.set(id, s);

  if (process.env.VITE_DEV_SERVER_URL) win.loadURL(process.env.VITE_DEV_SERVER_URL);
  else win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));

  // Ссылки из документа открываем во внешнем браузере, а не в окне редактора
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  win.on('close', (e) => {
    if (!s.dirty || s.forceClose) return;
    const choice = dialog.showMessageBoxSync(win, {
      type: 'question',
      buttons: ['Сохранить', 'Не сохранять', 'Отмена'],
      defaultId: 0,
      cancelId: 2,
      title: 'Перо',
      message: s.kind === 'sheet' ? 'Сохранить изменения в книге?' : 'Сохранить изменения в документе?',
    });
    if (choice === 1) return;
    e.preventDefault();
    if (choice === 0) send(s, 'saveAndClose');
  });
  // Окно закрыто штатно (сохранено или «Не сохранять») — копия для восстановления больше не нужна
  win.on('closed', () => {
    if (!s.keepAutosave) autosave.clear(s.key);
    wins.delete(id);
  });
  win.on('unresponsive', () => log.warn(`Окно не отвечает: ${s.path || s.kind}`));

  // Сбой процесса окна: перезапускаем окно, документ восстанавливается из его автосохранения (app:ready).
  // Больше двух сбоев за минуту — не зацикливаемся: сообщаем и оставляем копию до следующего запуска.
  win.webContents.on('render-process-gone', (_e, d) => {
    log.error(`Сбой окна (${d.reason}, код ${d.exitCode}): ${s.path || s.kind}`);
    if (d.reason === 'clean-exit' || win.isDestroyed()) return;
    const now = Date.now();
    s.crashes = [...s.crashes.filter((t) => now - t < 60_000), now];
    if (s.crashes.length > 2) {
      dialog.showErrorBox('Перо', 'Окно несколько раз подряд завершилось с ошибкой и будет закрыто.\n' +
        'Несохранённые изменения будут предложены к восстановлению при следующем запуске программы.');
      s.forceClose = true;
      s.keepAutosave = true; // копия нужна для восстановления при следующем запуске
      win.destroy();
      return;
    }
    s.afterCrash = true;
    s.ready = false;
    win.reload();
  });
  return s;
}

// ---------- IPC ----------
// Окно загрузилось: что в нём открыть. Перезапуск окна (сбой, перезагрузка) — восстановить его же автосохранение.
ipcMain.handle('app:ready', (e) => {
  const s = stateOf(e);
  if (!s) return { kind: 'doc', file: null, recovery: null, autosaveSec: AUTOSAVE_SEC };
  let { kind, file } = s.initial;
  let notice = null;
  const own = autosave.readOwn(s.key);
  if (own) {
    const m = own.meta;
    file = { path: m.path || null, name: m.name, data: toArrayBuffer(own.data), recovered: true };
    kind = m.kind || kind;
    notice = s.afterCrash ? 'Окно было перезапущено после сбоя. Несохранённые изменения восстановлены из автосохранения.' : null;
    log.info(`Окно восстановлено из автосохранения: ${m.path || m.name}`);
  }
  s.initial = { kind: s.kind, file: null };
  s.afterCrash = false;
  s.ready = true;
  let recovery = null;
  if (!recoveryOffered) {
    recoveryOffered = true;
    recovery = autosave.scan();
    if (recovery.length) log.info(`Найдено документов для восстановления: ${recovery.length}`);
    else recovery = null;
  }
  return { kind, file, recovery, notice, autosaveSec: AUTOSAVE_SEC };
});

const toArrayBuffer = (buf) => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

ipcMain.on('app:newWindow', (_e, kind) => createWindow({ kind: kind === 'sheet' ? 'sheet' : 'doc' }));

// Открыть: диалог; файл вернётся окну, если оно пустое и нетронутое, иначе откроется в новом окне
ipcMain.handle('file:open', async (e) => {
  const s = stateOf(e);
  const r = await dialog.showOpenDialog(s.win, { properties: ['openFile', 'multiSelections'], filters: OPEN_FILTER });
  if (r.canceled || !r.filePaths.length) return null;
  let here = null;
  for (const p of r.filePaths) {
    const f = await openPath(p, here ? null : s);
    if (f) { here = f; s.path = f.path; }
  }
  return here;
});

// Открыть по пути (перетаскивание, «Последние» на ленте) — та же логика выбора окна
ipcMain.handle('file:openPath', async (e, p) => {
  const s = stateOf(e);
  const f = await openPath(p, s);
  if (f) s.path = f.path;
  return f;
});

ipcMain.handle('recent:list', () => recent.list());
ipcMain.handle('recent:clear', () => { recent.clear(); try { app.clearRecentDocuments(); } catch { /* */ } buildMenu(); return true; });

// Печать: системный диалог; страницы режет Chromium по CSS @page (размер, поля, колонтитулы)
ipcMain.handle('doc:print', (e) => new Promise((resolve) => {
  e.sender.print({ printBackground: true }, (ok, reason) => resolve({ ok, reason }));
}));

// Экспорт в PDF: тот же CSS печати, размер листа — из @page
ipcMain.handle('doc:pdf', async (e, suggestedName) => {
  const s = stateOf(e);
  const r = await dialog.showSaveDialog(s.win, { defaultPath: suggestedName || 'Документ.pdf', filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (r.canceled || !r.filePath) return null;
  const p = /\.pdf$/i.test(r.filePath) ? r.filePath : r.filePath + '.pdf';
  const data = await e.sender.printToPDF({ preferCSSPageSize: true, printBackground: true });
  await fs.writeFile(p, data);
  log.info(`Экспорт в PDF: ${p}`);
  return { path: p, name: path.basename(p) };
});

// Имя пользователя Windows — автор примечаний и исправлений
ipcMain.handle('app:userName', () => {
  try { return os.userInfo().username || 'Пользователь'; } catch { return 'Пользователь'; }
});

// Выбор картинки для «Вставка → Рисунок»
const IMAGE_FILTER = [{ name: 'Изображения', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'] }];
ipcMain.handle('file:openImage', async (e) => {
  const r = await dialog.showOpenDialog(stateOf(e).win, { properties: ['openFile'], filters: IMAGE_FILTER });
  if (r.canceled || !r.filePaths[0]) return null;
  return readFile(r.filePaths[0]);
});

async function writeUserFile(s, p, data) {
  try {
    await fs.writeFile(p, Buffer.from(data));
  } catch (err) {
    log.error(`Не удалось сохранить ${p}: ${err.message}`);
    throw err;
  }
  s.path = p;
  rememberRecent(p);
  log.info(`Сохранён файл: ${p}`);
  return { path: p, name: path.basename(p) };
}

ipcMain.handle('file:save', (e, filePath, data) => writeUserFile(stateOf(e), filePath, data));

ipcMain.handle('file:saveAs', async (e, suggestedName, data) => {
  const s = stateOf(e);
  // расширение — по предложенному имени: книга сохраняется в .xlsx (или .csv), документ — в .docx
  const ext = (/\.(xlsx|csv)$/i.exec(suggestedName || '') || [, 'docx'])[1].toLowerCase();
  const filters = { xlsx: XLSX_FILTER, csv: CSV_FILTER, docx: DOCX_FILTER }[ext];
  const r = await dialog.showSaveDialog(s.win, { defaultPath: suggestedName || 'Документ.docx', filters });
  if (r.canceled || !r.filePath) return null;
  const p = r.filePath.toLowerCase().endsWith(`.${ext}`) ? r.filePath : `${r.filePath}.${ext}`;
  return writeUserFile(s, p, data);
});

ipcMain.handle('file:confirmDiscard', async (e) => {
  const s = stateOf(e);
  if (!s.dirty) return true;
  const choice = dialog.showMessageBoxSync(s.win, {
    type: 'question',
    buttons: ['Продолжить без сохранения', 'Отмена'],
    defaultId: 1,
    cancelId: 1,
    title: 'Перо',
    message: 'В документе есть несохранённые изменения. Продолжить?',
  });
  return choice === 0;
});

// Автосохранение: копия несохранённого документа окна; очистка — после сохранения или если правок нет
ipcMain.handle('autosave:write', (e, meta, data) => {
  const s = stateOf(e);
  try {
    autosave.write(s.key, { name: meta.name, path: meta.path || null, kind: meta.kind, pid: process.pid }, data);
    return true;
  } catch (err) {
    log.error(`Автосохранение не удалось: ${err.message}`);
    return false;
  }
});
ipcMain.on('autosave:clear', (e) => { const s = stateOf(e); if (s) autosave.clear(s.key); });

// Восстановление документов прошлого сеанса: открыть (в этом окне, если оно нетронутое, иначе в новом) или удалить
ipcMain.handle('recovery:open', (e, id) => {
  const s = stateOf(e);
  const here = pristine(s);
  const target = here ? s : createWindow({ kind: 'doc' });
  // копия переходит окну, где откроется документ: дальше оно само её обновляет, сбой до этого не теряет данные
  const got = autosave.adopt(id, target.key);
  if (!got) return null;
  const m = got.meta;
  const ext = m.kind === 'sheet' ? '.xlsx' : '.docx';
  // «(восстановлен)» — один раз, даже если копию восстанавливают повторно
  const base = String(m.name || 'Документ').replace(/\.(docx|xlsx|csv)$/i, '').replace(/( \(восстановлен\))+$/, '');
  const name = `${base} (восстановлен)${ext}`;
  autosave.write(target.key, { ...m, name, path: null, pid: process.pid }, got.data);
  log.info(`Восстановлен документ: ${m.path || m.name}`);
  // новое окно при загрузке прочитает своё автосохранение (app:ready)
  return here ? { path: null, name, data: toArrayBuffer(got.data), recovered: true } : null;
});
ipcMain.handle('recovery:discard', (_e, id) => { autosave.clear(id); log.info(`Удалена копия для восстановления: ${id}`); return true; });

ipcMain.on('edit', (e, cmd) => { if (['cut', 'copy', 'paste'].includes(cmd)) e.sender[cmd](); });
ipcMain.on('state:dirty', (e, value) => { const s = stateOf(e); if (s) s.dirty = !!value; });
ipcMain.on('state:doc', (e, meta) => { const s = stateOf(e); if (s && meta) { s.path = meta.path || null; s.kind = meta.kind === 'sheet' ? 'sheet' : 'doc'; } });
ipcMain.on('state:title', (e, title) => { const s = stateOf(e); if (s) s.win.setTitle(title); });
ipcMain.on('app:closeNow', (e) => { const s = stateOf(e); if (s) { s.forceClose = true; s.win.close(); } });
ipcMain.handle('dialog:error', (e, title, msg) => {
  log.error(`${title}: ${msg}`);
  dialog.showErrorBox(title, msg);
});
ipcMain.on('log:error', (e, msg) => { const s = stateOf(e); log.error(`[окно ${s ? s.path || s.kind : '?'}] ${msg}`); });
ipcMain.handle('app:report', (e) => reportProblem(stateOf(e) && stateOf(e).win));

// ---------- Запуск ----------
process.on('uncaughtException', (err) => { log.error(`Необработанная ошибка: ${err.stack || err}`); });
process.on('unhandledRejection', (err) => { log.error(`Необработанный отказ: ${(err && err.stack) || err}`); });

if (!NEW_INSTANCE && !app.requestSingleInstanceLock()) {
  // программа уже запущена — аргументы ушли первому экземпляру (second-instance)
  app.quit();
} else {
  app.on('second-instance', (_e, argv, cwd) => {
    const files = fileArgs(argv.slice(1), cwd);
    log.info(`Повторный запуск${files.length ? `: ${files.join(', ')}` : ''}`);
    if (files.length) files.forEach((f) => openPath(f));
    else createWindow({ kind: 'doc' });
  });
  app.on('child-process-gone', (_e, d) => log.error(`Сбой служебного процесса ${d.type}: ${d.reason} (${d.exitCode})`));

  app.whenReady().then(() => {
    log.info(`Запуск: Перо ${app.getVersion()}, Electron ${process.versions.electron}, Windows ${os.release()}, pid ${process.pid}${NEW_INSTANCE ? ', отдельный экземпляр' : ''}`);
    buildMenu();
    const files = fileArgs(process.argv.slice(app.isPackaged ? 1 : 2));
    // файлы — каждый в своё окно; не открылся ни один — пустой документ
    (async () => {
      for (const f of files) await openPath(f);
      if (!wins.size) createWindow({ kind: 'doc' });
    })();
  });
  app.on('window-all-closed', () => { log.info('Завершение работы'); app.quit(); });
}
