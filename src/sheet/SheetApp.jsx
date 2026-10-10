// Табличный редактор «Перо»: книга .xlsx — сетка, строка формул, лента, листы, файлы.
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Grid, { cellText } from './Grid.jsx';
import SheetRibbon from './SheetRibbon.jsx';
import { addr, key, parseRange, rangeName } from './address.js';
import { isErr } from './formula.js';
import { formatValue } from './format.js';
import {
  DEFAULT_COL_W, addSheet, autoSumFormula, blockToText, clearRange, collapseGroup, copyBlock, deleteSheet, expandGroup, fill, inputOf,
  isHidden, mergeAt, newWorkbook, pasteBlock, pasteText, renameSheet, setColWidth, setInput, setRowHeight, showLevels, usedRange,
} from './model.js';
import { exportXlsx, importXlsx } from './xlsx.js';
import * as fileApi from '../fileApi.js';
import { useAutosave } from '../autosave.js';

const DEFAULT_NAME = 'Книга1.xlsx';
const normSel = (s) => ({ r1: Math.min(s.r, s.r2), c1: Math.min(s.c, s.c2), r2: Math.max(s.r, s.r2), c2: Math.max(s.c, s.c2) });

export default function SheetApp({ initialFile, onOpenOther, onNewDoc, autosaveSec }) {
  const [wb, setWb] = useState(newWorkbook);
  const history = useRef({ undo: [], redo: [] });
  const [sel, setSelState] = useState({ r: 0, c: 0, r2: 0, c2: 0 });
  const [editing, setEditingState] = useState(null);
  const [doc, setDoc] = useState({ path: null, name: DEFAULT_NAME });
  const [dirty, setDirty] = useState(false);
  const [warnings, setWarnings] = useState({});
  const [kept, setKept] = useState({}); // что Перо не редактирует, но сохраняет как есть
  const source = useRef(null); // байты открытого .xlsx — экспорт строится поверх них
  const [renaming, setRenaming] = useState(null);
  const gridRef = useRef(null);
  const clip = useRef(null); // { block, text } — последний скопированный блок (с формулами и стилями)
  const si = wb.active;
  const sheet = wb.sheets[si];

  // изменение книги с записью в историю отмены
  const change = useCallback((fn) => {
    setWb((cur) => {
      const next = fn(cur);
      if (next === cur) return cur;
      history.current.undo.push(cur);
      if (history.current.undo.length > 200) history.current.undo.shift();
      history.current.redo = [];
      setDirty(true);
      return next;
    });
  }, []);
  const undo = () => { const h = history.current; if (!h.undo.length) return; h.redo.push(wb); setWb(h.undo.pop()); setDirty(true); };
  const redo = () => { const h = history.current; if (!h.redo.length) return; h.undo.push(wb); setWb(h.redo.pop()); setDirty(true); };

  const setSel = useCallback((s) => setSelState((cur) => (typeof s === 'function' ? s(cur) : s)), []);
  const setEditing = (e) => {
    if (!e && gridRef.current) gridRef.current.focus(); // Esc — фокус обратно сетке
    if (e && e.text === null) e = { ...e, text: inputOf(sheet.cells.get(key(e.r, e.c))) };
    setEditingState(e);
  };

  // ---------- Файлы ----------
  const load = useCallback(async (file) => {
    try {
      const { wb: book, warnings: w, kept: k } = await importXlsx(file.data);
      source.current = file.data;
      setKept(k || {});
      setWb(book);
      history.current = { undo: [], redo: [] };
      setSelState({ r: 0, c: 0, r2: 0, c2: 0 });
      setEditingState(null);
      setDoc({ path: file.path, name: file.name });
      setWarnings(w);
      setDirty(!!file.recovered); // восстановленная после сбоя — ещё не сохранена
    } catch (e) {
      fileApi.showError('Не удалось открыть книгу', `${file.name}\n\n${e.message || e}`);
    }
  }, []);
  useEffect(() => { if (initialFile) load(initialFile); }, [initialFile, load]);

  // только для отладки (npm run dev): выгрузить текущую книгу без диалога сохранения
  useEffect(() => { if (import.meta.env.DEV) window.__peroSheet = { wb, exportXlsx: () => exportXlsx(wb, { source: source.current }) }; }, [wb]);

  useEffect(() => { fileApi.setDirty(dirty); fileApi.setTitle(`${doc.name}${dirty ? ' •' : ''} — Перо`); }, [dirty, doc.name]);
  useEffect(() => { fileApi.setDoc({ path: doc.path, kind: 'sheet' }); }, [doc.path]);
  // книга неизменяема: новая книга после каждой правки — сама и есть «номер версии» для автосохранения
  useAutosave({ intervalSec: autosaveSec, dirty, revision: wb, kind: 'sheet', name: doc.name, path: doc.path, getBytes: () => exportXlsx(wb, { source: source.current }) });

  const write = async (asNew) => {
    try {
      const bytes = await exportXlsx(wb, { source: source.current });
      const r = asNew || !doc.path ? await fileApi.saveFileAs(doc.name, bytes) : await fileApi.saveFile(doc.path, doc.name, bytes);
      if (!r) return false;
      setDoc({ path: r.path, name: r.name });
      setDirty(false);
      setWarnings({});
      return true;
    } catch (e) {
      fileApi.showError('Не удалось сохранить книгу', String(e.message || e));
      return false;
    }
  };
  const save = async (then) => {
    if (doc.path && Object.keys(warnings).length && !window.confirm(
      `В книге есть то, что Перо пока не сохраняет (${Object.keys(warnings).join(', ')}). Перезаписать файл? (Отмена — сохранить копию)`)) {
      const ok = await write(true);
      if (ok && typeof then === 'function') then();
      return;
    }
    const ok = await write(false);
    if (ok && typeof then === 'function') then();
  };
  // файл для этого окна (в программе — только если окно было пустым, иначе он уже открыт в новом окне)
  const openHere = (f) => {
    if (!f) return;
    if (/\.docx$/i.test(f.name)) { fileApi.setDirty(false); onOpenOther(f); } else load(f);
  };
  const open = async () => {
    if (!fileApi.isDesktop && dirty && !(await fileApi.confirmDiscard())) return;
    openHere(await fileApi.openFile());
  };
  const newBook = async () => {
    if (fileApi.newWindow('sheet')) return; // в программе — новое окно
    if (dirty && !(await fileApi.confirmDiscard())) return;
    setWb(newWorkbook()); history.current = { undo: [], redo: [] }; setDoc({ path: null, name: DEFAULT_NAME }); setDirty(false); setWarnings({}); setKept({}); source.current = null;
    setSelState({ r: 0, c: 0, r2: 0, c2: 0 });
  };
  const newDocument = async () => {
    if (fileApi.newWindow('doc')) return;
    if (dirty && !(await fileApi.confirmDiscard())) return;
    setDirty(false);
    onNewDoc();
  };

  // команды меню Electron
  const handlers = useRef({});
  handlers.current = { save, saveAs: () => write(true), open, newBook, newDocument, load };
  useEffect(() => fileApi.onMenu((cmd, payload) => {
    const h = handlers.current;
    if (cmd === 'new') h.newDocument();
    else if (cmd === 'newSheet') h.newBook();
    else if (cmd === 'open') h.open();
    else if (cmd === 'save') h.save();
    else if (cmd === 'saveAs') h.saveAs();
    else if (cmd === 'openData') { if (/\.docx$/i.test(payload.name)) onOpenOther(payload); else h.load(payload); }
    else if (cmd === 'saveAndClose') h.save(() => fileApi.closeNow());
    else if (cmd === 'print' || cmd === 'pdf') { if (cmd === 'print') fileApi.print(); else fileApi.exportPdf(doc.name.replace(/\.xlsx$/i, '') + '.pdf'); }
  }), [onOpenOther, doc.name]);

  // ---------- Ввод ----------
  // следующая позиция курсора (с учётом объединённых ячеек)
  const step = (s, dir, extend) => {
    const d = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] }[dir];
    if (extend) return { ...s, r2: Math.max(0, s.r2 + d[0]), c2: Math.max(0, s.c2 + d[1]) };
    const m = mergeAt(sheet, s.r, s.c);
    let r = s.r + d[0]; let c = s.c + d[1];
    if (m) { if (d[0] > 0) r = m.r2 + 1; if (d[1] > 0) c = m.c2 + 1; }
    r = Math.max(0, r); c = Math.max(0, c);
    // скрытые строки/столбцы пропускаются (в начале листа — назад, к видимой)
    for (let n = 0; n < 100000 && d[0] && isHidden(sheet, 'row', r); n++) r = r + d[0] < 0 ? r + 1 : r + d[0];
    for (let n = 0; n < 16384 && d[1] && isHidden(sheet, 'col', c); n++) c = c + d[1] < 0 ? c + 1 : c + d[1];
    if (isHidden(sheet, 'row', r) || isHidden(sheet, 'col', c)) return { r: s.r, c: s.c, r2: s.r, c2: s.c };
    const m2 = mergeAt(sheet, r, c);
    if (m2) { r = m2.r1; c = m2.c1; }
    return { r, c, r2: r, c2: c };
  };
  const move = (dir, extend = false) => setSel((s) => step(s, dir, extend));
  // завершить ввод; фокус возвращается сетке сразу, чтобы следующие нажатия не потерялись
  const commit = (text, dir) => {
    const ed = editing;
    if (gridRef.current) gridRef.current.focus();
    setEditingState(null);
    if (!ed) return;
    if (text !== inputOf(sheet.cells.get(key(ed.r, ed.c)))) change((w) => setInput(w, si, ed.r, ed.c, text));
    if (dir) setSelState(step({ r: ed.r, c: ed.c, r2: ed.r, c2: ed.c }, dir));
  };
  // щелчок по ячейке при наборе формулы — вставить ссылку (replace — заменить последнюю вставленную при протягивании)
  const lastRef = useRef(null);
  const insertRef = (rg, replace) => {
    setEditingState((ed) => {
      if (!ed) return ed;
      const ref = rangeName(rg);
      let text = ed.text;
      if (replace && lastRef.current && text.endsWith(lastRef.current)) text = text.slice(0, -lastRef.current.length);
      else if (/[A-Z]+\d+(?::[A-Z]+\d+)?$/.test(text) && lastRef.current && text.endsWith(lastRef.current)) text = text.slice(0, -lastRef.current.length);
      lastRef.current = ref;
      return { ...ed, text: text + ref };
    });
  };

  // Ctrl+стрелка — к краю блока данных
  const jump = (dir) => {
    const d = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] }[dir];
    const has = (r, c) => { const cell = sheet.cells.get(key(r, c)); return cell && (cell.v != null || cell.f != null); };
    let { r, c } = sel;
    const lim = (rr, cc) => rr >= 0 && cc >= 0 && rr < 1048576 && cc < 16384;
    if (has(r, c) && lim(r + d[0], c + d[1]) && has(r + d[0], c + d[1])) {
      while (lim(r + d[0], c + d[1]) && has(r + d[0], c + d[1])) { r += d[0]; c += d[1]; }
    } else {
      let steps = 0;
      do { r += d[0]; c += d[1]; steps++; } while (lim(r, c) && !has(r, c) && steps < 20000);
      if (!lim(r, c) || !has(r, c)) { const u = usedRange(sheet); r = Math.max(0, dir === 'up' ? 0 : dir === 'down' ? (u ? u.r2 : 0) : sel.r); c = Math.max(0, dir === 'left' ? 0 : dir === 'right' ? (u ? u.c2 : 0) : sel.c); }
    }
    setSel({ r, c, r2: r, c2: c });
  };

  // перетаскивание .xlsx/.docx в окно
  const onDrop = async (e) => {
    const f = [...e.dataTransfer.files].find((x) => /\.(xlsx|docx)$/i.test(x.name));
    if (!f) return;
    e.preventDefault();
    const p = fileApi.pathForFile(f);
    if (p) { openHere(await fileApi.openPath(p)); return; } // программа сама решит: это окно или новое
    if (dirty && !(await fileApi.confirmDiscard())) return;
    openHere({ path: null, name: f.name, data: await f.arrayBuffer() });
  };

  // кнопки поля структуры: свернуть/развернуть группу, показать уровень
  const outlineAction = (axis, a) => change((w) => (a.type === 'collapse' ? collapseGroup(w, si, axis, a.a, a.b)
    : a.type === 'expand' ? expandGroup(w, si, axis, a.a, a.b, a.level) : showLevels(w, si, axis, a.n)));

  const rg = normSel(sel);
  const ribbonApi = {
    wb, sheet, si, sel, rg, change, undo, redo, canUndo: history.current.undo.length > 0, canRedo: history.current.redo.length > 0,
    file: { newDocument, newBook, open, openHere, save: () => save(), saveAs: () => write(true), print: () => fileApi.print(), pdf: () => fileApi.exportPdf(doc.name.replace(/\.xlsx$/i, '') + '.pdf') },
    autoSum: () => {
      const target = rg.r2 > rg.r1 || rg.c2 > rg.c1 ? { r: rg.r2 + 1, c: rg.c1 } : { r: sel.r, c: sel.c };
      const f = rg.r2 > rg.r1 && rg.c1 === rg.c2 ? `SUM(${addr(rg.r1, rg.c1)}:${addr(rg.r2, rg.c1)})` : autoSumFormula(sheet, target.r, target.c);
      if (f === 'SUM()') { setEditing({ r: target.r, c: target.c, text: '=SUM()' }); return; }
      change((w) => setInput(w, si, target.r, target.c, `=${f}`));
      setSel({ r: target.r, c: target.c, r2: target.r, c2: target.c });
    },
  };

  const onKeyDown = (e) => {
    if (editing) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (ctrl && e.code === 'KeyZ') { e.preventDefault(); undo(); return; }
    if (ctrl && (e.code === 'KeyY')) { e.preventDefault(); redo(); return; }
    if (ctrl && e.code === 'KeyA') { e.preventDefault(); const u = usedRange(sheet); setSel({ r: 0, c: 0, r2: u ? u.r2 : 0, c2: u ? u.c2 : 0 }); return; }
    if (ctrl && e.code === 'KeyB') { e.preventDefault(); ribbonApi.toggle('bold'); return; }
    if (ctrl && e.code === 'KeyI') { e.preventDefault(); ribbonApi.toggle('italic'); return; }
    if (ctrl && e.code === 'KeyU') { e.preventDefault(); ribbonApi.toggle('underline'); return; }
    if (ctrl && e.code === 'KeyD') { e.preventDefault(); change((w) => fill(w, si, rg, 'down')); return; }
    if (ctrl && e.code === 'KeyR') { e.preventDefault(); change((w) => fill(w, si, rg, 'right')); return; }
    if (ctrl && e.code === 'KeyS') { if (!fileApi.isDesktop) { e.preventDefault(); save(); } return; }
    if (e.altKey && (e.code === 'Equal')) { e.preventDefault(); ribbonApi.autoSum(); return; }
    if (k.startsWith('Arrow')) {
      e.preventDefault();
      const dir = k.slice(5).toLowerCase();
      if (ctrl) jump(dir); else move(dir, e.shiftKey);
      return;
    }
    if (k === 'Tab') { e.preventDefault(); move(e.shiftKey ? 'left' : 'right'); return; }
    if (k === 'Enter') { e.preventDefault(); move(e.shiftKey ? 'up' : 'down'); return; }
    if (k === 'Home') { e.preventDefault(); const r = ctrl ? 0 : sel.r; setSel({ r, c: 0, r2: r, c2: 0 }); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); const r = Math.max(0, sel.r + (k === 'PageDown' ? 25 : -25)); setSel({ r, c: sel.c, r2: r, c2: sel.c }); return; }
    if (k === 'F2') { e.preventDefault(); setEditing({ r: sel.r, c: sel.c, text: null }); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); change((w) => clearRange(w, si, rg, 'contents')); return; }
    if (!ctrl && !e.altKey && k.length === 1) { e.preventDefault(); lastRef.current = null; setEditingState({ r: sel.r, c: sel.c, text: k }); }
  };

  // буфер обмена: TSV (совместимо с Excel) + свой блок с формулами и стилями
  const onCopy = (e, cut) => {
    if (editing) return;
    e.preventDefault();
    const block = copyBlock(sheet, rg);
    const text = blockToText(block, cellText);
    e.clipboardData.setData('text/plain', text);
    clip.current = { block, text };
    if (cut) change((w) => clearRange(w, si, rg, 'all'));
  };
  const onPaste = (e) => {
    if (editing) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    if (clip.current && text === clip.current.text) change((w) => pasteBlock(w, si, rg.r1, rg.c1, clip.current.block));
    else if (text) change((w) => pasteText(w, si, rg.r1, rg.c1, text));
  };

  // строка состояния: сумма/среднее/количество выделенного
  const stats = useMemo(() => {
    let sum = 0; let count = 0; let nums = 0;
    for (const [k, cell] of sheet.cells) {
      const i = k.indexOf(','); const r = +k.slice(0, i); const c = +k.slice(i + 1);
      if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) continue;
      if (cell.v == null || cell.v === '') continue;
      count++;
      if (typeof cell.v === 'number') { sum += cell.v; nums++; }
    }
    return { sum, count, nums };
  }, [sheet.cells, rg.r1, rg.r2, rg.c1, rg.c2]);

  const activeCell = sheet.cells.get(key(sel.r, sel.c));
  const barText = editing && editing.r === sel.r && editing.c === sel.c ? editing.text : inputOf(activeCell);

  return (
    <div className="app sheet-app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <div className="titlebar">
        <div className="quick-access">
          <span className="app-mark">П</span>
          <button type="button" className="rb-btn" title="Сохранить (Ctrl+S)" onClick={() => save()}>💾</button>
          <button type="button" className="rb-btn" title="Отменить (Ctrl+Z)" disabled={!history.current.undo.length} onClick={undo}>↶</button>
          <button type="button" className="rb-btn" title="Повторить (Ctrl+Y)" disabled={!history.current.redo.length} onClick={redo}>↷</button>
        </div>
        <div className="doc-title">{doc.name}{dirty ? ' •' : ''}</div>
      </div>
      <SheetRibbon api={ribbonApi} />
      <div className="formula-bar">
        <input className="name-box" aria-label="Адрес" value={rangeName(rg).replace(/^(\w+):\1$/, '$1')} onChange={() => {}}
          onKeyDown={(e) => {
            if (e.key !== 'Enter') return;
            const p = parseRange(e.currentTarget.value.toUpperCase());
            if (p) setSel({ r: p.r1, c: p.c1, r2: p.r2, c2: p.c2 });
            gridRef.current && gridRef.current.focus();
          }}
          onFocus={(e) => e.target.select()} />
        <span className="fx">fx</span>
        <input className="fx-input" aria-label="Строка формул" value={barText}
          onFocus={() => { if (!editing) { lastRef.current = null; setEditingState({ r: sel.r, c: sel.c, text: inputOf(activeCell), fromBar: true }); } }}
          onChange={(e) => setEditingState({ r: sel.r, c: sel.c, text: e.target.value, fromBar: true })}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { e.preventDefault(); commit(editing ? editing.text : barText, 'down'); }
            if (e.key === 'Escape') { e.preventDefault(); setEditingState(null); gridRef.current && gridRef.current.focus(); }
          }} />
      </div>
      <div className="sheet-main" ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={(e) => onCopy(e, false)} onCut={(e) => onCopy(e, true)} onPaste={onPaste}>
        <Grid sheet={sheet} sel={sel} setSel={setSel} editing={editing} setEditing={setEditing} onCommit={commit}
          onColWidth={(c, px) => change((w) => setColWidth(w, si, c, px))} onRowHeight={(r, px) => change((w) => setRowHeight(w, si, r, px))}
          onRefInsert={insertRef} gridRef={gridRef} onOutline={outlineAction} />
      </div>
      <div className="sheet-tabs">
        {wb.sheets.map((s, i) => (renaming === i ? (
          <input key={i} className="sheet-tab editing" autoFocus defaultValue={s.name}
            onBlur={(e) => { const v = e.target.value; setRenaming(null); if (v !== s.name) change((w) => renameSheet(w, i, v)); }}
            onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); if (e.key === 'Escape') setRenaming(null); }} />
        ) : (
          <button key={i} type="button" className={`sheet-tab${i === si ? ' active' : ''}`} title="Двойной щелчок — переименовать"
            onClick={() => { setEditingState(null); setWb((w) => ({ ...w, active: i })); setSelState({ r: 0, c: 0, r2: 0, c2: 0 }); }}
            onDoubleClick={() => setRenaming(i)}>
            {s.name}
            {i === si && wb.sheets.length > 1 && (
              <span className="sheet-del" title="Удалить лист" onClick={(e) => {
                e.stopPropagation();
                if (window.confirm(`Удалить лист «${s.name}»? Ссылки на него станут ошибками #REF!.`)) change((w) => deleteSheet(w, i));
              }}>×</span>
            )}
          </button>
        )))}
        <button type="button" className="sheet-tab add" title="Новый лист" onClick={() => change((w) => addSheet(w))}>+</button>
      </div>
      <div className="statusbar">
        <span>{Object.keys(warnings).length ? <span className="warn" title="Не сохранится при перезаписи">⚠ Не поддерживается: {Object.keys(warnings).join(', ')}</span> : 'Готово'}</span>
        {Object.keys(kept).length > 0 && (
          <span className="kept" title={'Перо пока не редактирует эти элементы, но сохраняет их без изменений:\n' + Object.keys(kept).map((k) => `• ${k}`).join('\n')}>
            Сохраняются без изменений: {Object.keys(kept).join(', ')}
          </span>
        )}
        <span className="spacer" />
        {stats.count > 1 && (
          <>
            {stats.nums > 0 && <span>Среднее: {formatValue(stats.sum / stats.nums)}</span>}
            <span>Количество: {stats.count}</span>
            {stats.nums > 0 && <span>Сумма: {formatValue(stats.sum)}</span>}
          </>
        )}
        {isErr(activeCell && activeCell.v) && <span className="warn">{errorHint(activeCell.v.error)}</span>}
      </div>
      <style>{'@media print { @page { size: A4 landscape; margin: 15mm; } }'}</style>
      <PrintTable sheet={sheet} />
    </div>
  );
}

function errorHint(code) {
  return {
    '#DIV/0!': 'Деление на ноль', '#VALUE!': 'Неверный тип значения', '#REF!': 'Ссылка на удалённую ячейку или лист',
    '#NAME?': 'Неизвестная функция или имя', '#N/A': 'Значение не найдено', '#NUM!': 'Недопустимое число', '#CIRC!': 'Циклическая ссылка',
  }[code] || code;
}

// Для печати/PDF — таблица использованной области листа (видна только при печати)
function PrintTable({ sheet }) {
  const u = usedRange(sheet);
  if (!u) return null;
  const widths = Array.from({ length: u.c2 + 1 }, (_, c) => sheet.colW[c] ?? DEFAULT_COL_W);
  const rows = [];
  for (let r = 0; r <= u.r2; r++) {
    const tds = [];
    for (let c = 0; c <= u.c2; c++) {
      const m = mergeAt(sheet, r, c);
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const cell = sheet.cells.get(key(r, c));
      const s = (cell && cell.s) || {};
      tds.push(
        <td key={c} rowSpan={m ? m.r2 - m.r1 + 1 : undefined} colSpan={m ? m.c2 - m.c1 + 1 : undefined}
          style={{ fontWeight: s.bold ? 'bold' : undefined, fontStyle: s.italic ? 'italic' : undefined, background: s.fill, color: s.color,
            textAlign: s.align || (cell && typeof cell.v === 'number' ? 'right' : 'left') }}>
          {cellText(cell)}
        </td>,
      );
    }
    rows.push(<tr key={r}>{tds}</tr>);
  }
  return (
    <table className="sheet-print" style={{ width: widths.reduce((a, b) => a + b, 0) }}>
      <colgroup>{widths.map((w, c) => <col key={c} style={{ width: w }} />)}</colgroup>
      <tbody>{rows}</tbody>
    </table>
  );
}
