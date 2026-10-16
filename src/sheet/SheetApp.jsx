// Табличный редактор «Перо»: книга .xlsx — сетка, строка формул, лента, листы, файлы.
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Grid, { cellText } from './Grid.jsx';
import SheetRibbon from './SheetRibbon.jsx';
import { addr, colName, key, parseRange, rangeName } from './address.js';
import { cycleRef, errText, formulaRefs, isErr, localize, sheetPrefix } from './formula.js';

// цвета ссылок редактируемой формулы (рамки на листе) — как в Excel
const REF_COLORS = ['#2f6fbf', '#c0392b', '#7a5ab5', '#2e8b57', '#b8860b', '#1f8a9e', '#c2185b', '#6d4c41'];

// $A$1:$B$5 — абсолютная ссылка на выделение (для имён)
const absRange = (rg) => {
  const a = `$${colName(rg.c1)}$${rg.r1 + 1}`;
  return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a : `${a}:$${colName(rg.c2)}$${rg.r2 + 1}`;
};
import { formatValue } from './format.js';
import {
  DEFAULT_COL_W, DEFAULT_ROW_H, addSheet, autoSumFormula, blockToText, clearRange, collapseGroup, copyBlock, deleteSheet, expandGroup, fill, inputOf,
  addCfRule, applyCellFormat, clearCf, setPageBreak, setValidation, currentRegion, fillSeries, isHidden, isValidName, mergeAt, nameTarget, newWorkbook, reapplyFilter, recalc, setAutoFilter, setCalcMode,
  setColumnFilter, setNames, sortRows, spillAnchorAt, valueAt, viewContext, pasteBlock, pasteText, renameSheet, setColWidth, setInput, setRowHeight, showLevels, usedRange,
} from './model.js';
import { exportXlsx, importXlsx } from './xlsx.js';
import { exportCsv, importCsv } from './csv.js';
import { importXls } from './xls.js';
import { exportXls } from './xlsWrite.js';
import { setPage } from './pageSetup.js';
import { isLocked, isProtected, protectionViolation, setProtection } from './protect.js';
import ProtectDialog from './ProtectDialog.jsx';
import NamesDialog from './NamesDialog.jsx';
import FilterMenu, { CustomFilterDialog } from './FilterMenu.jsx';
import SortDialog from './SortDialog.jsx';
import FormatCellsDialog from './FormatCellsDialog.jsx';
import FindDialog from './FindDialog.jsx';
import DvDialog, { DvErrorDialog } from './DvDialog.jsx';
import PrintPreview, { PrintPages } from './PrintPreview.jsx';
import { paginate } from './printPages.js';
import { dvAt, listItems, validate } from './dv.js';
import { CfDialog } from './CfMenu.jsx';
import { cfView } from './cf.js';
import { filterCount, shownText } from './filter.js';
import EvalDialog from './EvalDialog.jsx';
import InsertFunctionDialog from './InsertFunctionDialog.jsx';
import { AssistPopup, useFormulaAssist } from './FormulaAssist.jsx';
import TitleBar, { QuickButton } from '../TitleBar.jsx';

const PROTECTED_MSG = 'Ячейка на защищённом листе. Чтобы изменить её, снимите защиту: «Рецензирование → Снять защиту листа».';
import * as fileApi from '../fileApi.js';
import { useAutosave } from '../autosave.js';

const DEFAULT_NAME = 'Книга1.xlsx';
const isCsv = (name) => /\.csv$/i.test(name || '');
const isXlsName = (name) => /\.xls$/i.test(name || '');
const CSV_LOSS = 'В формате CSV сохраняется только текущий лист: значения без формул и оформления. Сохранить в CSV?';
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
  const [notice, setNotice] = useState(null); // временное сообщение в строке состояния
  const [protectDlg, setProtectDlg] = useState(null); // 'protect' | 'unprotect'
  const [namesDlg, setNamesDlg] = useState(false);
  const [filterMenu, setFilterMenu] = useState(null); // { j, x, y } — меню фильтра столбца
  const [customFilter, setCustomFilter] = useState(null); // { j, kind, preset }
  const [sortDlg, setSortDlg] = useState(null); // диапазон
  const [cfDlg, setCfDlg] = useState(null); // вид правила условного форматирования
  const [fcDlg, setFcDlg] = useState(false); // «Формат ячеек»
  const [findDlg, setFindDlg] = useState(null); // 'find' | 'replace'
  const [dvDlg, setDvDlg] = useState(false); // «Проверка данных»
  const [preview, setPreview] = useState(false); // «Печать» — предварительный просмотр
  const [showBreaks, setShowBreaks] = useState(false); // пунктир разрывов страниц на листе (после просмотра, как Excel)
  const [dvErr, setDvErr] = useState(null); // { rule, ed, text, dir } — значение не прошло проверку
  const [dvPop, setDvPop] = useState(null); // { r, c, items, x, y } — выпадающий список ячейки
  const [showFormulas, setShowFormulas] = useState(false); // Ctrl+` — формулы вместо значений
  const [arrows, setArrows] = useState([]); // стрелки влияющих/зависимых ячеек
  const [evalAt, setEvalAt] = useState(null); // { r, c } — «Вычислить формулу»
  const [fnDlg, setFnDlg] = useState(null); // мастер функций: { initial }
  const [errPop, setErrPop] = useState(null); // пояснение к ошибке: { r, c, x, y }
  const noticeTimer = useRef(null);
  const flash = useCallback((text) => {
    setNotice(text);
    clearTimeout(noticeTimer.current);
    noticeTimer.current = setTimeout(() => setNotice(null), 5000);
  }, []);
  const gridRef = useRef(null);
  const clip = useRef(null); // { block, text } — последний скопированный блок (с формулами и стилями)
  const si = wb.active;
  const sheet = wb.sheets[si];
  const cfMap = useMemo(() => cfView(wb, si), [wb, si]); // условное форматирование: что показать поверх стилей
  const pageBreaks = useMemo(() => (showBreaks ? paginate(sheet) : null), [showBreaks, sheet]);

  // изменение книги с записью в историю отмены
  const change = useCallback((fn) => {
    setWb((cur) => {
      const next = fn(cur);
      if (next === cur) return cur;
      if (protectionViolation(cur, next)) { setTimeout(() => flash(PROTECTED_MSG)); return cur; }
      history.current.undo.push(cur);
      if (history.current.undo.length > 200) history.current.undo.shift();
      history.current.redo = [];
      setDirty(true);
      return next;
    });
  }, [flash]);
  const undo = () => { const h = history.current; if (!h.undo.length) return; h.redo.push(wb); setWb(h.undo.pop()); setDirty(true); };
  const redo = () => { const h = history.current; if (!h.redo.length) return; h.undo.push(wb); setWb(h.redo.pop()); setDirty(true); };

  const setSel = useCallback((s) => setSelState((cur) => (typeof s === 'function' ? s(cur) : s)), []);
  const setEditing = (e) => {
    if (!e && gridRef.current) gridRef.current.focus(); // Esc — фокус обратно сетке
    if (e && isProtected(sheet) && isLocked(sheet.cells.get(key(e.r, e.c)))) { flash(PROTECTED_MSG); return; }
    if (e && e.text === null) e = { ...e, text: inputOf(sheet.cells.get(key(e.r, e.c))) };
    setEditingState(e);
  };

  // ---------- Файлы ----------
  const load = useCallback(async (file) => {
    try {
      // CSV: книга из одного листа, исходного пакета нет; восстановленная копия CSV — уже .xlsx
      const csv = isCsv(file.name) && !file.recovered;
      // .xls (Excel 97–2003) — свой разбор BIFF8 (xls.js); исходный пакет .xlsx у такой книги нет
      const xls = isXlsName(file.name) && !file.recovered;
      const { wb: book, warnings: w, kept: k } = csv ? { ...importCsv(file.data, file.name), warnings: {}, kept: {} } : xls ? await importXls(file.data) : await importXlsx(file.data);
      source.current = csv || xls ? null : file.data;
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
      // открытый CSV: «Сохранить» — снова в CSV, «Сохранить как» — книгой .xlsx или .xls
      const csv = isCsv(doc.name) && !asNew;
      if (csv && (wb.sheets.length > 1 || [...wb.sheets[wb.active].cells.values()].some((c) => c.f != null || c.s)) && !window.confirm(CSV_LOSS)) return false;
      let r;
      if (csv) r = await fileApi.saveFile(doc.path, doc.name, exportCsv(wb.sheets[wb.active]));
      else {
        // путь — до сборки байтов: формат выбирается в диалоге (.xlsx или Excel 97–2003)
        const suggested = isCsv(doc.name) ? doc.name.replace(/\.csv$/i, '.xlsx') : doc.name;
        const p = asNew || !doc.path ? await fileApi.pickSavePath(suggested) : doc.path;
        if (!p) return false;
        let bytes;
        if (isXlsName(p)) {
          const { bytes: b, lost } = exportXls(wb);
          const what = Object.entries(lost).map(([k, n]) => (n > 1 ? `${k} (${n})` : k));
          if (what.length && !window.confirm(`В формате Excel 97–2003 (.xls) не сохранится: ${what.join(', ')}. Сохранить?`)) return false;
          bytes = b;
        } else {
          // книга из .xls: рисунки, диаграммы, примечания и т. п. переносятся только в .xls (записями как есть)
          const fromXls = [...new Set(wb.sheets.flatMap((s) => (s.xlsKeep ? s.xlsKeep.labels : [])))].filter((l) => l !== 'параметры принтера');
          if (fromXls.length && !window.confirm(`В формате .xlsx не сохранится (из книги Excel 97–2003 Перо переносит это только в .xls): ${fromXls.join(', ')}. Сохранить?`)) return false;
          bytes = await exportXlsx(wb, { source: source.current });
        }
        const name = p.split(/[\\/]/).pop();
        r = fileApi.isDesktop ? await fileApi.saveFile(p, name, bytes) : await fileApi.saveFileAs(name, bytes);
      }
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
  // «Экспорт в CSV»: копия текущего листа, книга остаётся открытой как была
  const saveCsv = async () => {
    try {
      await fileApi.saveFileAs(`${doc.name.replace(/\.(xlsx|xls|csv)$/i, '')} (${sheet.name}).csv`, exportCsv(sheet));
    } catch (e) {
      fileApi.showError('Не удалось сохранить CSV', String(e.message || e));
    }
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
    else if (cmd === 'print') { setPreview(true); setShowBreaks(true); }
    else if (cmd === 'pdf') fileApi.exportPdf(doc.name.replace(/\.(xlsx|xls|csv)$/i, '') + '.pdf');
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
  // как Excel: ввод строки через Tab, затем Enter — в столбец, с которого начали, на строку ниже
  const tabStart = useRef(null); // { r, c }
  const nav = (from, dir) => {
    if (dir === 'tab' || dir === 'tabBack') {
      if (!tabStart.current || tabStart.current.r !== from.r) tabStart.current = { r: from.r, c: from.c };
      return step({ r: from.r, c: from.c, r2: from.r, c2: from.c }, dir === 'tab' ? 'right' : 'left');
    }
    const t = tabStart.current; tabStart.current = null;
    if (dir === 'down' && t && t.r === from.r) return { r: from.r + 1, c: t.c, r2: from.r + 1, c2: t.c };
    return step({ r: from.r, c: from.c, r2: from.r, c2: from.c }, dir);
  };
  // завершить ввод; фокус возвращается сетке сразу, чтобы следующие нажатия не потерялись
  const commit = (text, dir) => {
    const ed = editing;
    if (gridRef.current) gridRef.current.focus();
    setEditingState(null);
    if (!ed) return;
    if (text !== inputOf(sheet.cells.get(key(ed.r, ed.c)))) {
      // проверка данных: значение (результат формулы — тоже) должно подходить под правило ячейки
      const rule = dvAt(sheet, ed.r, ed.c);
      if (rule && rule.showError && rule.type !== 'any') {
        const next = setInput(wb, si, ed.r, ed.c, text);
        if (!validate(next, si, ed.r, ed.c, valueAt(next.sheets[si], ed.r, ed.c))) { setDvErr({ rule, ed, text, dir }); return; }
      }
      change((w) => setInput(w, si, ed.r, ed.c, text));
    }
    if (dir) setSelState(nav(ed, dir));
  };
  // выпадающий список проверки данных у ячейки (стрелка справа или Alt+↓)
  const openDvList = (r, c, rect) => {
    const rule = dvAt(sheet, r, c);
    if (!rule || rule.type !== 'list' || rule.inCellDropdown === false) return false;
    setDvPop({ r, c, items: listItems(wb, si, rule, r, c), x: rect ? rect.left : 300, y: rect ? rect.bottom : 300, w: rect && rect.width ? rect.width + 18 : 140 });
    return true;
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
    const f = [...e.dataTransfer.files].find((x) => /\.(xlsx|xls|docx|csv)$/i.test(x.name));
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

  const refMarks = useMemo(() => {
    if (!editing || !editing.text || !editing.text.startsWith('=')) return [];
    let refs = [];
    try { refs = formulaRefs(editing.text.slice(1)); } catch { return []; }
    return refs.filter((x) => !x.sheet || x.sheet.toLowerCase() === sheet.name.toLowerCase()).map((x, i) => ({ ...x, color: REF_COLORS[i % REF_COLORS.length] }));
  }, [editing, sheet.name]);
  // стрелки «Зависимости формул»: влияющие — ссылки формулы ячейки; зависимые — формулы листа, ссылающиеся на неё
  const sameSheet = (s) => !s || s.toLowerCase() === sheet.name.toLowerCase();
  const precedents = () => {
    const cell = sheet.cells.get(key(sel.r, sel.c));
    if (!cell || cell.f == null) { flash('В ячейке нет формулы'); return; }
    const refs = formulaRefs(cell.f);
    const local = refs.filter((x) => sameSheet(x.sheet));
    if (refs.length > local.length) flash('Есть ссылки на другие листы — стрелки к ним не рисуются');
    setArrows((a) => [...a, ...local.map((x) => ({ from: x, to: { r: sel.r, c: sel.c } }))]);
  };
  const dependents = () => {
    const out = [];
    for (const [k, cell] of sheet.cells) {
      if (cell.f == null) continue;
      if (formulaRefs(cell.f).some((x) => sameSheet(x.sheet) && sel.r >= x.r1 && sel.r <= x.r2 && sel.c >= x.c1 && sel.c <= x.c2)) {
        const [r, c] = k.split(',').map(Number);
        out.push({ from: { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }, to: { r, c } });
      }
    }
    if (!out.length) flash('На эту ячейку не ссылаются формулы листа');
    setArrows((a) => [...a, ...out]);
  };
  const evaluateFormula = () => {
    const cell = sheet.cells.get(key(sel.r, sel.c));
    if (!cell || cell.f == null) { flash('В ячейке нет формулы'); return; }
    setEvalAt({ r: sel.r, c: sel.c });
  };

  const ribbonApi = {
    names: () => setNamesDlg(true),
    insertFunction: (en) => setFnDlg({ initial: en }),
    showFormulas, toggleFormulas: () => setShowFormulas((v) => !v),
    precedents, dependents, clearArrows: () => setArrows([]), evaluateFormula,
    // F9 — пересчитать всё (в том числе отложенное в режиме «вручную» и летучие функции); в отмену не попадает
    recalc: () => setWb((w) => recalc(w, { full: true })),
    formatCells: () => setFcDlg(true),
    validation: () => setDvDlg(true),
    showBreaks, toggleBreaks: () => setShowBreaks((v) => !v),
    pageBreak: (mode) => { change((w) => setPageBreak(w, si, sel.r, sel.c, mode)); setShowBreaks(true); },
    find: (mode) => setFindDlg(mode),
    // условное форматирование выделенного диапазона
    cfDialog: (kind) => setCfDlg(kind),
    addCf: (rule) => change((w) => addCfRule(w, si, [rg], rule)),
    clearCf: (all) => change((w) => clearCf(w, si, all ? null : rg)),
    // автофильтр: включить на текущей области/выделении или снять; очистить условия; применить повторно
    filterOn: !!sheet.autoFilter,
    toggleFilter: () => change((w) => (w.sheets[si].autoFilter ? setAutoFilter(w, si, null) : setAutoFilter(w, si, rg))),
    clearFilter: () => change((w) => { let x = w; for (const j of Object.keys((w.sheets[si].autoFilter || {}).cols || {})) x = setColumnFilter(x, si, +j, null); return x; }),
    reapplyFilter: () => change((w) => reapplyFilter(w, si)),
    sortDialog: () => setSortDlg(rg.r1 === rg.r2 && rg.c1 === rg.c2 ? (sheet.autoFilter || currentRegion(sheet, rg.r1, rg.c1)) && (sheet.autoFilter ? { ...sheet.autoFilter.ref } : currentRegion(sheet, rg.r1, rg.c1)) : rg),
    calcMode: wb.calcMode === 'manual' ? 'manual' : 'auto',
    setCalcMode: (mode) => change((w) => setCalcMode(w, mode)),
    protect: () => setProtectDlg(isProtected(sheet) ? 'unprotect' : 'protect'),
    wb, sheet, si, sel, rg, change, undo, redo, canUndo: history.current.undo.length > 0, canRedo: history.current.redo.length > 0,
    file: { newDocument, newBook, open, openHere, save: () => save(), saveAs: () => write(true), saveCsv, print: () => { setPreview(true); setShowBreaks(true); }, pdf: () => fileApi.exportPdf(doc.name.replace(/\.(xlsx|xls|csv)$/i, '') + '.pdf') },
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
    if (e.key === 'F9') { e.preventDefault(); ribbonApi.recalc(); return; }
    if (ctrl && e.code === 'Backquote') { e.preventDefault(); setShowFormulas((v) => !v); return; }
    if (e.shiftKey && e.key === 'F3') { e.preventDefault(); setFnDlg({}); return; }
    if (ctrl && e.code === 'KeyA') { e.preventDefault(); const u = usedRange(sheet); setSel({ r: 0, c: 0, r2: u ? u.r2 : 0, c2: u ? u.c2 : 0 }); return; }
    if (e.altKey && k === 'ArrowDown') { const b = document.querySelector('.gdvbtn'); if (b) { e.preventDefault(); b.click(); return; } }
    if (ctrl && (e.code === 'KeyF' || e.code === 'KeyH') && !e.shiftKey) { e.preventDefault(); setFindDlg(e.code === 'KeyH' ? 'replace' : 'find'); return; }
    if (ctrl && (e.code === 'Digit1' || e.code === 'Numpad1') && !e.shiftKey) { e.preventDefault(); setFcDlg(true); return; }
    if (ctrl && e.shiftKey && e.code === 'KeyL') { e.preventDefault(); ribbonApi.toggleFilter(); return; }
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
    if (k === 'Tab') { e.preventDefault(); if (sel.r === sel.r2 && sel.c === sel.c2) setSel(nav(sel, e.shiftKey ? 'tabBack' : 'tab')); else move(e.shiftKey ? 'left' : 'right'); return; }
    if (k === 'Enter') { e.preventDefault(); if (!e.shiftKey && sel.r === sel.r2 && sel.c === sel.c2) setSel(nav(sel, 'down')); else move(e.shiftKey ? 'up' : 'down'); return; }
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
    const take = (v) => {
      if (v == null || v === '') return;
      count++;
      if (typeof v === 'number') { sum += v; nums++; }
    };
    const inSel = (k) => { const i = k.indexOf(','); const r = +k.slice(0, i); const c = +k.slice(i + 1); return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2; };
    for (const [k, cell] of sheet.cells) if (inSel(k)) take(cell.v ?? (cell.f == null && sheet.spill ? sheet.spill.get(k) : null));
    if (sheet.spill) for (const [k, v] of sheet.spill) if (!sheet.cells.has(k) && inSel(k)) take(v);
    return { sum, count, nums };
  }, [sheet.cells, sheet.spill, rg.r1, rg.r2, rg.c1, rg.c2]);

  const activeCell = sheet.cells.get(key(sel.r, sel.c));
  const formulaHidden = isProtected(sheet) && activeCell && activeCell.f != null && activeCell.s && activeCell.s.formulaHidden;
  const spillOwner = (!activeCell || (activeCell.v == null && activeCell.f == null)) ? spillAnchorAt(sheet, sel.r, sel.c) : null;
  const ghost = spillOwner && !(spillOwner.r === sel.r && spillOwner.c === sel.c);
  const barText = editing && editing.r === sel.r && editing.c === sel.c ? editing.text : formulaHidden ? ''
    : ghost ? inputOf(sheet.cells.get(key(spillOwner.r, spillOwner.c))) : inputOf(activeCell);

  return (
    <div className="app sheet-app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <TitleBar name={doc.name} dirty={dirty}>
        <QuickButton icon="save" title="Сохранить (Ctrl+S)" onClick={() => save()} />
        <QuickButton icon="undo" title="Отменить (Ctrl+Z)" disabled={!history.current.undo.length} onClick={undo} />
        <QuickButton icon="redo" title="Повторить (Ctrl+Y)" disabled={!history.current.redo.length} onClick={redo} />
      </TitleBar>
      <SheetRibbon api={ribbonApi} />
      <div className="formula-bar">
        <input className="name-box" aria-label="Адрес" value={rangeName(rg).replace(/^(\w+):\1$/, '$1')} onChange={() => {}}
          onKeyDown={(e) => {
            if (e.key !== 'Enter') return;
            const text = e.currentTarget.value.trim();
            const p = parseRange(text.toUpperCase());
            if (p) setSel({ r: p.r1, c: p.c1, r2: p.r2, c2: p.c2 });
            else {
              const t = nameTarget(wb, text, sheet.name);
              if (t) {
                const i = wb.sheets.findIndex((x) => x.name.toLowerCase() === t.sheet.toLowerCase());
                if (i >= 0 && i !== si) setWb((w) => ({ ...w, active: i }));
                setSel({ r: t.r1, c: t.c1, r2: t.r2, c2: t.c2 });
              } else if (isValidName(text)) {
                // новое имя для выделенного диапазона, как в Excel
                const ref = `${sheetPrefix(sheet.name)}${absRange(rg)}`;
                change((w) => setNames(w, [...(w.names || []), { name: text, text: ref, sheet: null }]));
                flash(`Имя «${text}» присвоено диапазону ${absRange(rg).replace(/\$/g, '')}`);
              } else flash('Неверная ссылка или имя');
            }
            gridRef.current && gridRef.current.focus();
          }}
          onFocus={(e) => e.target.select()} />
        <button type="button" className="fx-btn" title="Вставить функцию (Shift+F3)" onMouseDown={(e) => e.preventDefault()} onClick={() => setFnDlg({})}>fx</button>
        <FormulaBarInput className={`fx-input${ghost && !(editing && editing.r === sel.r && editing.c === sel.c) ? ' ghost' : ''}`} value={barText} names={wb.names || []}
          title={ghost ? 'Значение рассчитано формулой массива — изменить можно в её первой ячейке' : undefined}
          readOnly={isProtected(sheet) && isLocked(activeCell)}
          onFocus={() => { if (!editing && !(isProtected(sheet) && isLocked(activeCell))) { lastRef.current = null; setEditingState({ r: sel.r, c: sel.c, text: ghost ? '' : inputOf(activeCell), fromBar: true }); } }}
          onChange={(text) => setEditingState({ r: sel.r, c: sel.c, text, fromBar: true })}
          onCommit={() => commit(editing ? editing.text : barText, 'down')}
          onCancel={() => { setEditingState(null); gridRef.current && gridRef.current.focus(); }} />
      </div>
      <div className="sheet-main" ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={(e) => onCopy(e, false)} onCut={(e) => onCopy(e, true)} onPaste={onPaste}>
        <Grid sheet={sheet} sel={sel} setSel={setSel} editing={editing} setEditing={setEditing} onCommit={commit}
          onColWidth={(c, px) => change((w) => setColWidth(w, si, c, px))} onRowHeight={(r, px) => change((w) => setRowHeight(w, si, r, px))}
          onRefInsert={insertRef} gridRef={gridRef} onOutline={outlineAction} names={wb.names || []}
          refMarks={refMarks} showFormulas={showFormulas} arrows={arrows}
          onErrorInfo={(r, c, rect) => setErrPop({ r, c, x: rect.right + 4, y: rect.top })}
          cf={cfMap}
          pageBreaks={pageBreaks}
          dvList={(() => { const rule = dvAt(sheet, sel.r, sel.c); return rule && rule.type === 'list' && rule.inCellDropdown !== false && !editing; })()}
          dvPrompt={(() => { const rule = dvAt(sheet, sel.r, sel.c); return rule && rule.showPrompt && (rule.prompt || rule.promptTitle) && !editing ? rule : null; })()}
          onDvList={(rect) => openDvList(sel.r, sel.c, rect)}
          onFilterMenu={(j, rect) => setFilterMenu({ j, x: rect.left, y: rect.bottom + 2 })}
          onFill={(src, dst, ctrl) => { change((w) => fillSeries(w, si, src, dst, { ctrl })); setSel({ r: dst.r1, c: dst.c1, r2: dst.r2, c2: dst.c2 }); }} />
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
            {isProtected(s) && <span className="sheet-lock" title="Лист защищён" aria-label="Лист защищён">🔒︎</span>}
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
        {notice && <span className="warn" role="status">{notice}</span>}
        {sheet.autoFilter && Object.keys(sheet.autoFilter.cols).length > 0 && (() => { const n = filterCount(sheet, sheet.autoFilter); return <span>Найдено записей: {n.shown} из {n.total}</span>; })()}
        {!notice && isProtected(sheet) && <span className="kept" title="Заблокированные ячейки этого листа нельзя изменить">Лист защищён</span>}
        {wb.calcPending && <button type="button" className="status-btn" title="Режим вычислений «вручную»: есть несосчитанные изменения. Пересчитать (F9)" onClick={ribbonApi.recalc}>Вычислить</button>}
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
      <PrintPages sheet={sheet} cf={cfMap} />
      {evalAt && (
        <EvalDialog cellName={`${sheetPrefix(sheet.name)}${addr(evalAt.r, evalAt.c)}`} formula={(sheet.cells.get(key(evalAt.r, evalAt.c)) || {}).f || ''}
          ctx={viewContext(wb, si, evalAt.r, evalAt.c)} onClose={() => { setEvalAt(null); gridRef.current && gridRef.current.focus(); }} />
      )}
      {fnDlg && (
        <InsertFunctionDialog ctx={viewContext(wb, si, sel.r, sel.c)} initial={fnDlg.initial}
          onClose={() => { setFnDlg(null); gridRef.current && gridRef.current.focus(); }}
          onInsert={(text) => { setFnDlg(null); setEditingState(null); change((w) => setInput(w, si, sel.r, sel.c, text)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {errPop && (() => {
        const v = valueAt(sheet, errPop.r, errPop.c);
        const hasFormula = (sheet.cells.get(key(errPop.r, errPop.c)) || {}).f != null;
        return (
          <div className="modal-backdrop transparent" onMouseDown={() => setErrPop(null)}>
            <div className="err-pop" style={{ left: errPop.x, top: errPop.y }} onMouseDown={(e) => e.stopPropagation()}>
              <div className="err-title">{isErr(v) ? errText(v.error) : ''}</div>
              <div className="err-text">{isErr(v) ? errorHint(v.error) : ''}</div>
              {hasFormula && <button type="button" onClick={() => { setErrPop(null); setEvalAt({ r: errPop.r, c: errPop.c }); }}>Показать этапы вычисления…</button>}
              <button type="button" onClick={() => setErrPop(null)}>Закрыть</button>
            </div>
          </div>
        );
      })()}
      {namesDlg && (
        <NamesDialog wb={wb} refOfSelection={localize(`${sheetPrefix(sheet.name)}${absRange(rg)}`)}
          onClose={() => { setNamesDlg(false); gridRef.current && gridRef.current.focus(); }}
          onSave={(list) => { setNamesDlg(false); change((w) => setNames(w, list)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {filterMenu && sheet.autoFilter && (
        <FilterMenu sheet={sheet} j={filterMenu.j} x={filterMenu.x} y={filterMenu.y}
          onClose={() => { setFilterMenu(null); gridRef.current && gridRef.current.focus(); }}
          onApply={(crit) => { const { j } = filterMenu; setFilterMenu(null); change((w) => setColumnFilter(w, si, j, crit)); gridRef.current && gridRef.current.focus(); }}
          onSort={(desc) => {
            const { j } = filterMenu; setFilterMenu(null);
            change((w) => { const af = w.sheets[si].autoFilter; return sortRows(w, si, af.ref, [{ col: af.ref.c1 + j, desc }], { header: true }); });
            gridRef.current && gridRef.current.focus();
          }}
          onCustom={(kind, preset) => { const { j } = filterMenu; setFilterMenu(null); setCustomFilter({ j, kind, preset }); }} />
      )}
      {customFilter && sheet.autoFilter && (
        <CustomFilterDialog kind={customFilter.kind} preset={customFilter.preset} initial={sheet.autoFilter.cols[customFilter.j]}
          head={shownText(sheet.cells.get(key(sheet.autoFilter.ref.r1, sheet.autoFilter.ref.c1 + customFilter.j))) || `Столбец ${colName(sheet.autoFilter.ref.c1 + customFilter.j)}`}
          onClose={() => { setCustomFilter(null); gridRef.current && gridRef.current.focus(); }}
          onApply={(crit) => { const { j } = customFilter; setCustomFilter(null); change((w) => setColumnFilter(w, si, j, crit)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {preview && (
        <PrintPreview sheet={sheet} cf={cfMap} onClose={() => { setPreview(false); gridRef.current && gridRef.current.focus(); }}
          onPage={(patch) => change((w) => setPage(w, si, patch))}
          onPrint={() => fileApi.print()}
          onPdf={() => fileApi.exportPdf(doc.name.replace(/\.(xlsx|xls|csv)$/i, '') + '.pdf')} />
      )}
      {dvPop && (
        <div className="modal-backdrop clear" onMouseDown={() => setDvPop(null)}>
          <div className="dv-pop" style={{ left: dvPop.x, top: dvPop.y, minWidth: dvPop.w }} role="listbox" onMouseDown={(e) => e.stopPropagation()}>
            {dvPop.items.map((it, i) => (
              <button key={i} type="button" className="dv-pop-item" onClick={() => { const { r, c } = dvPop; setDvPop(null); change((w) => setInput(w, si, r, c, it)); gridRef.current && gridRef.current.focus(); }}>{it}</button>
            ))}
            {!dvPop.items.length && <div className="fm-empty">Список пуст</div>}
          </div>
        </div>
      )}
      {dvErr && (
        <DvErrorDialog rule={dvErr.rule}
          onCommit={() => { const { ed, text, dir } = dvErr; setDvErr(null); change((w) => setInput(w, si, ed.r, ed.c, text)); if (dir) setSelState(nav(ed, dir)); gridRef.current && gridRef.current.focus(); }}
          onRetry={() => { const { ed, text } = dvErr; setDvErr(null); setSelState({ r: ed.r, c: ed.c, r2: ed.r, c2: ed.c }); setEditingState({ r: ed.r, c: ed.c, text }); }}
          onCancel={() => { setDvErr(null); gridRef.current && gridRef.current.focus(); }} />
      )}
      {dvDlg && (
        <DvDialog rule={dvAt(sheet, rg.r1, rg.c1)} onClose={() => { setDvDlg(false); gridRef.current && gridRef.current.focus(); }}
          onClear={() => { setDvDlg(false); change((w) => setValidation(w, si, rg, null)); gridRef.current && gridRef.current.focus(); }}
          onApply={(rule) => { setDvDlg(false); change((w) => setValidation(w, si, rg, rule)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {findDlg && (
        <FindDialog key={findDlg} wb={wb} si={si} sel={sel} mode={findDlg} onClose={() => { setFindDlg(null); gridRef.current && gridRef.current.focus(); }}
          onMessage={flash}
          onGo={(x) => { if (x.si !== si) setWb((w) => ({ ...w, active: x.si })); setSelState({ r: x.r, c: x.c, r2: x.r, c2: x.c }); }}
          onChange={(fn) => change(fn)} />
      )}
      {fcDlg && (
        <FormatCellsDialog cell={sheet.cells.get(key(rg.r1, rg.c1))} rg={rg} merged={!!mergeAt(sheet, rg.r1, rg.c1)}
          sample={(() => { const v = valueAt(sheet, rg.r1, rg.c1); return v == null || v === '' ? undefined : v; })()}
          onClose={() => { setFcDlg(false); gridRef.current && gridRef.current.focus(); }}
          onApply={(res) => { setFcDlg(false); change((w) => applyCellFormat(w, si, rg, res)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {cfDlg && (
        <CfDialog kind={cfDlg} rg={rg} onClose={() => { setCfDlg(null); gridRef.current && gridRef.current.focus(); }}
          onApply={(rule) => { setCfDlg(null); change((w) => addCfRule(w, si, [rg], rule)); gridRef.current && gridRef.current.focus(); }} />
      )}
      {sortDlg && (
        <SortDialog sheet={sheet} rg={sortDlg} onClose={() => { setSortDlg(null); gridRef.current && gridRef.current.focus(); }}
          onApply={(levels, header) => { const r = sortDlg; setSortDlg(null); change((w) => sortRows(w, si, r, levels, { header })); gridRef.current && gridRef.current.focus(); }} />
      )}
      {protectDlg && (
        <ProtectDialog mode={protectDlg} protection={sheet.protection} onClose={() => { setProtectDlg(null); gridRef.current && gridRef.current.focus(); }}
          onDone={(p) => { setProtectDlg(null); change((w) => setProtection(w, si, p)); gridRef.current && gridRef.current.focus(); }} />
      )}
    </div>
  );
}

function errorHint(code) {
  return {
    '#DIV/0!': 'Деление на ноль', '#VALUE!': 'Неверный тип значения', '#REF!': 'Ссылка на удалённую ячейку или лист',
    '#NAME?': 'Неизвестная функция или имя', '#N/A': 'Значение не найдено', '#NUM!': 'Недопустимое число', '#CIRC!': 'Циклическая ссылка',
    '#SPILL!': 'Результат-массив не помещается: ячейки рядом заняты', '#CALC!': 'Пустой результат (например, ФИЛЬТР ничего не нашёл)',
  }[code] || code;
}

// Строка формул с подсказками функций и F4 (те же, что в редакторе ячейки)
function FormulaBarInput({ value, names, onChange, onCommit, onCancel, ...rest }) {
  const ref = useRef(null);
  const assist = useFormulaAssist(value, names);
  const apply = ({ text, caret }) => {
    onChange(text);
    assist.setCaret(caret);
    requestAnimationFrame(() => { const el = ref.current; if (el) { el.focus(); el.setSelectionRange(caret, caret); } });
  };
  const box = ref.current ? { left: ref.current.offsetLeft, top: ref.current.offsetTop + ref.current.offsetHeight + 2 } : { left: 0, top: 30 };
  return (
    <>
      <input ref={ref} aria-label="Строка формул" value={value} {...rest}
        onChange={(e) => { onChange(e.target.value); assist.setCaret(e.target.selectionStart); }}
        onSelect={(e) => assist.setCaret(e.currentTarget.selectionStart)}
        onKeyDown={(e) => {
          if (assist.onKey(e, apply)) return;
          if (e.key === 'F4') { e.preventDefault(); const r = cycleRef(value, e.currentTarget.selectionStart); if (r) apply(r); return; }
          if (e.key === 'Enter') { e.preventDefault(); onCommit(); }
          if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
        }} />
      {document.activeElement === ref.current && <AssistPopup assist={assist} style={box} onPick={(i) => { const r = assist.accept(i); if (r) apply(r); }} />}
    </>
  );
}
