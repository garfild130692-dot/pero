// Доступ к файлам: в Electron — через preload-мост, в обычном браузере (для отладки) — через input/скачивание.
const native = typeof window !== 'undefined' ? window.pero : undefined;

export const isDesktop = !!native;

export async function openFile() {
  if (native) return native.openFile();
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx,.xlsx';
    input.onchange = async () => {
      const f = input.files[0];
      resolve(f ? { path: null, name: f.name, data: await f.arrayBuffer() } : null);
    };
    input.click();
  });
}

// Выбрать картинку: { name, data } или null
export async function openImage() {
  if (native && native.openImage) return native.openImage();
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = async () => {
      const f = input.files[0];
      resolve(f ? { path: null, name: f.name, data: await f.arrayBuffer() } : null);
    };
    input.click();
  });
}

export async function saveFile(path, name, data) {
  if (native) return path ? native.saveFile(path, data) : native.saveFileAs(name, data);
  return download(name, data);
}

export async function saveFileAs(name, data) {
  if (native) return native.saveFileAs(name, data);
  return download(name, data);
}

function download(name, data) {
  const type = /\.xlsx$/i.test(name || '') ? 'spreadsheetml.sheet' : 'wordprocessingml.document';
  const blob = new Blob([data], { type: `application/vnd.openxmlformats-officedocument.${type}` });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name || 'Документ.docx';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  return { path: null, name: a.download };
}

// Печать и PDF (в браузере — окно печати, PDF — через «Сохранить как PDF» в нём)
export const print = () => (native && native.print ? native.print() : window.print());
export const exportPdf = (name) => (native && native.exportPdf ? native.exportPdf(name) : window.print());

// Имя пользователя: в Electron — из Windows, в браузере — «Пользователь»
export const about = () => (native && native.about ? native.about() : alert('Перо'));
export const openLogs = () => native && native.openLogs && native.openLogs();
export const userName = async () => (native && native.userName ? native.userName() : 'Пользователь');

export const confirmDiscard = async () => (native ? native.confirmDiscard() : window.confirm('Есть несохранённые изменения. Продолжить?'));
export const showError = (title, msg) => (native ? native.showError(title, msg) : alert(`${title}\n\n${msg}`));
export const setDirty = (v) => native && native.setDirty(v);
export const setTitle = (t) => { document.title = t; native && native.setTitle(t); };
export const closeNow = () => native && native.closeNow();
export const onMenu = (cb) => (native ? native.onMenu(cb) : () => {});
export const clipboard = (cmd) => (native ? native.edit(cmd) : document.execCommand(cmd));
export const pathForFile = (f) => (native && native.pathForFile ? native.pathForFile(f) : null);

// ---------- Окна, последние файлы, автосохранение, журнал (этап 9) ----------
// Что открыть при загрузке окна; в браузере (отладка) — ?sheet для книги
export const ready = async () => (native && native.ready ? native.ready() : {
  kind: /[?&]sheet\b/.test(window.location.search) ? 'sheet' : /[?&]start\b/.test(window.location.search) ? 'start' : 'doc', file: null, recovery: null, notice: null, autosaveSec: 0,
});
// Новый документ/книга в отдельном окне (как Ctrl+N в Word). false — нет окон (браузер): заменить содержимое текущего
export const newWindow = (kind) => { if (!native || !native.newWindow) return false; native.newWindow(kind); return true; };
// Открыть по пути: данные файла — если открывать здесь (окно пустое), null — открылось в другом окне
export const openPath = (p) => (native && native.openPath ? native.openPath(p) : Promise.resolve(null));
export const recentList = async (check = false) => (native && native.recentList ? native.recentList(check) : []);
export const recentPin = async (p, v) => native && native.recentPin && native.recentPin(p, v);
export const recentRemove = async (p) => native && native.recentRemove && native.recentRemove(p);
export const recoveryList = async () => (native && native.recoveryList ? native.recoveryList() : []);
export const settingsGet = async () => (native && native.settingsGet ? native.settingsGet() : { startScreen: true });
export const settingsSet = async (k, v) => native && native.settingsSet && native.settingsSet(k, v);
export const recentClear = async () => native && native.recentClear && native.recentClear();
export const setDoc = (meta) => native && native.setDoc && native.setDoc(meta);
export const autosave = async (meta, data) => native && native.autosave && native.autosave(meta, data);
export const autosaveClear = () => native && native.autosaveClear && native.autosaveClear();
export const recoveryOpen = (id) => native.recoveryOpen(id);
export const recoveryDiscard = (id) => native.recoveryDiscard(id);
export const logError = (msg) => { console.error(msg); if (native && native.logError) native.logError(msg); };
export const reportProblem = () => (native && native.reportProblem ? native.reportProblem()
  : alert('В браузерном режиме отчёт не собирается: откройте консоль разработчика.'));
