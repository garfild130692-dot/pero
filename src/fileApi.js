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
export const userName = async () => (native && native.userName ? native.userName() : 'Пользователь');

export const confirmDiscard = async () => (native ? native.confirmDiscard() : window.confirm('Есть несохранённые изменения. Продолжить?'));
export const showError = (title, msg) => (native ? native.showError(title, msg) : alert(`${title}\n\n${msg}`));
export const setDirty = (v) => native && native.setDirty(v);
export const setTitle = (t) => { document.title = t; native && native.setTitle(t); };
export const closeNow = () => native && native.closeNow();
export const onMenu = (cb) => (native ? native.onMenu(cb) : () => {});
export const clipboard = (cmd) => (native ? native.edit(cmd) : document.execCommand(cmd));
export const pathForFile = (f) => (native && native.pathForFile ? native.pathForFile(f) : null);
