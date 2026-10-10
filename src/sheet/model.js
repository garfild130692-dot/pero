// Модель книги и операции над ней. Все операции возвращают новую книгу (неизменённые листы/ячейки — общие),
// поэтому отмена — просто стек предыдущих книг.
//
// Книга: { sheets: Sheet[], active }
// Лист:  { name, cells: Map<"r,c", Cell>, colW: { c: px }, rowH: { r: px }, merges: Range[],
//          rowInfo/colInfo: { n: { hidden, level, collapsed, xf } } — скрытие, группировка, стиль строки/столбца (xf файла),
//          src: { part, rid, name, shifts } — лист из файла (экспорт переписывает его поверх исходного XML) }
// Ячейка: { v — значение (у формулы — вычисленное), f — формула без «=» (хранимый вид), s — стиль }
// Стиль: { bold, italic, underline, strike, color, fill, fontSize, fontFamily, align, valign, wrap, numFmt }
import { key, unkey } from './address.js';
import { ERR, err, fromUser, localize, evaluate, isErr, offsetFormula, parse, rewriteRefs, shiftFormula } from './formula.js';
import { parseInput } from './format.js';

export const DEFAULT_COL_W = 72;  // px (≈ 8,43 символа Calibri 11, как в Excel)
export const DEFAULT_ROW_H = 20;  // px (15 pt)

export function newSheet(name) {
  return { name, cells: new Map(), colW: {}, rowH: {}, merges: [], rowInfo: {}, colInfo: {} };
}
export function newWorkbook() {
  return { sheets: [newSheet('Лист1')], active: 0 };
}

const withSheet = (wb, i, sheet) => ({ ...wb, sheets: wb.sheets.map((s, k) => (k === i ? sheet : s)) });
const sheetByName = (wb, name) => wb.sheets.find((s) => s.name.toLowerCase() === String(name).toLowerCase());

// ---------- Пересчёт ----------
// Все формулы книги: memo по (лист, ячейка), циклические ссылки → #CIRC!
export function recalc(wb) {
  const memo = new Map();
  const visiting = new Set();
  const astCache = new Map();
  const astOf = (f) => {
    if (!astCache.has(f)) { try { astCache.set(f, parse(f)); } catch { astCache.set(f, null); } }
    return astCache.get(f);
  };
  const get = (sheetName, r, c) => {
    const sheet = sheetByName(wb, sheetName);
    if (!sheet) return err(ERR.REF);
    const k = key(r, c);
    const cell = sheet.cells.get(k);
    if (!cell) return null;
    if (cell.f == null) return cell.v ?? null;
    const mk = `${sheet.name}!${k}`;
    if (memo.has(mk)) return memo.get(mk);
    if (visiting.has(mk)) return err(ERR.CIRC);
    visiting.add(mk);
    const ast = astOf(cell.f);
    let v = ast ? evaluate(ast, { get, sheet: sheet.name }) : err(ERR.NAME);
    if (Array.isArray(v)) v = v[0] ? v[0][0] : null;
    if (v == null) v = 0; // пустая ссылка в формуле — 0, как в Excel
    visiting.delete(mk);
    memo.set(mk, v);
    return v;
  };
  const sheets = wb.sheets.map((sheet) => {
    let changed = false;
    const cells = new Map(sheet.cells);
    for (const [k, cell] of sheet.cells) {
      if (cell.f == null) continue;
      const [r, c] = unkey(k);
      const v = get(sheet.name, r, c);
      if (!sameValue(v, cell.v)) { cells.set(k, { ...cell, v }); changed = true; }
    }
    return changed ? { ...sheet, cells } : sheet;
  });
  return { ...wb, sheets };
}
const sameValue = (a, b) => (isErr(a) && isErr(b) ? a.error === b.error : a === b);

// ---------- Ввод ----------
// Ввод пользователя в ячейку: «=…» — формула; иначе число/дата/проценты/текст (формат подставляется, если был «Общий»)
export function setInput(wb, si, r, c, raw) {
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  const k = key(r, c);
  const old = cells.get(k) || {};
  const s = String(raw ?? '');
  let cell;
  if (s.startsWith('=') && s.length > 1) {
    cell = { ...old, f: fromUser(s.slice(1)), v: null };
  } else {
    const p = parseInput(s);
    cell = { ...old, v: p.value };
    delete cell.f;
    if (p.format && (!old.s || !old.s.numFmt || old.s.numFmt === 'General')) cell.s = { ...(old.s || {}), numFmt: p.format };
  }
  if (cell.v == null && cell.f == null && !cell.s) cells.delete(k);
  else cells.set(k, cell);
  return recalc(withSheet(wb, si, { ...sheet, cells }));
}

// Текст для строки формул / редактирования
export function inputOf(cell) {
  if (!cell) return '';
  if (cell.f != null) return `=${localize(cell.f)}`;
  const v = cell.v;
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'number') {
    const fmt = cell.s && cell.s.numFmt;
    if (fmt && /%/.test(fmt)) return `${String(Math.round(v * 100 * 1e10) / 1e10).replace('.', ',')}%`;
    if (fmt && /[dy]/i.test(fmt.replace(/"[^"]*"/g, ''))) {
      // дата — как её удобно вводить
      return formatForEdit(v, fmt);
    }
    return String(Math.round(v * 1e10) / 1e10).replace('.', ',');
  }
  if (isErr(v)) return v.error;
  return String(v);
}
function formatForEdit(v, fmt) {
  const EPOCH = Date.UTC(1899, 11, 30);
  const d = new Date(EPOCH + Math.round(v * 86400000));
  const p = (n) => String(n).padStart(2, '0');
  const date = `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
  return /h/i.test(fmt) ? `${date} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}` : date;
}

// ---------- Стили ----------
export function setStyle(wb, si, rg, patch) {
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  for (let r = rg.r1; r <= rg.r2; r++) {
    for (let c = rg.c1; c <= rg.c2; c++) {
      const k = key(r, c);
      const old = cells.get(k) || { v: null };
      const s = { ...(old.s || {}), ...patch };
      for (const [n, v] of Object.entries(s)) if (v == null || v === false || v === '') delete s[n];
      const cell = { ...old, s };
      if (!Object.keys(s).length) delete cell.s;
      if (cell.v == null && cell.f == null && !cell.s) cells.delete(k); else cells.set(k, cell);
    }
  }
  return withSheet(wb, si, { ...sheet, cells });
}

// Очистить: what — 'all' | 'contents' | 'formats'
export function clearRange(wb, si, rg, what = 'contents') {
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  for (const [k, cell] of sheet.cells) {
    const [r, c] = unkey(k);
    if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) continue;
    if (what === 'all') { cells.delete(k); continue; }
    if (what === 'formats') { const n = { ...cell }; delete n.s; if (n.v == null && n.f == null) cells.delete(k); else cells.set(k, n); continue; }
    if (cell.s) cells.set(k, { s: cell.s, v: null }); else cells.delete(k);
  }
  return recalc(withSheet(wb, si, { ...sheet, cells }));
}

// ---------- Строки и столбцы ----------
// axis 'row' | 'col'; count > 0 — вставить перед at, < 0 — удалить |count| начиная с at.
// Формулы всех листов, ссылающиеся на этот лист, переписываются (удалённые ссылки — #REF!).
export function shiftCells(wb, si, axis, at, count) {
  const name = wb.sheets[si].name;
  const sheets = wb.sheets.map((sheet, i) => {
    let cells = sheet.cells;
    if (i === si) {
      cells = new Map();
      for (const [k, cell] of sheet.cells) {
        let [r, c] = unkey(k);
        const v = axis === 'row' ? r : c;
        if (count < 0 && v >= at && v < at - count) continue; // удалена
        const nv = v >= at ? v + count : v;
        if (axis === 'row') r = nv; else c = nv;
        cells.set(key(r, c), cell);
      }
    }
    // формулы: ссылки на изменённый лист
    let changed = cells !== sheet.cells;
    const out = new Map(cells);
    for (const [k, cell] of cells) {
      if (cell.f == null) continue;
      const f = shiftFormula(cell.f, { sheet: name, current: sheet.name, axis, at, count });
      if (f !== cell.f) { out.set(k, { ...cell, f }); changed = true; }
    }
    if (!changed) return sheet;
    const next = { ...sheet, cells: out };
    if (i === si) {
      for (const dim of axis === 'row' ? ['rowH', 'rowInfo'] : ['colW', 'colInfo']) {
        const sizes = {};
        for (const [n, x] of Object.entries(sheet[dim] || {})) {
          const v = +n;
          if (count < 0 && v >= at && v < at - count) continue;
          sizes[v >= at ? v + count : v] = x;
        }
        next[dim] = sizes;
      }
      next.merges = sheet.merges.map((m) => shiftMerge(m, axis, at, count)).filter(Boolean);
      // лист из файла: сдвиг запоминается — экспорт сдвинет диапазоны условного форматирования, проверки, ссылок
      if (sheet.src) next.src = { ...sheet.src, shifts: [...(sheet.src.shifts || []), { axis, at, count }] };
    }
    return next;
  });
  return recalc({ ...wb, sheets });
}
function shiftMerge(m, axis, at, count) {
  const [a, b] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
  const out = { ...m };
  if (count > 0) {
    if (m[a] >= at) { out[a] += count; out[b] += count; }
    else if (m[b] >= at) out[b] += count;
  } else {
    const n = -count;
    const end = at + n - 1;
    if (m[a] > end) { out[a] -= n; out[b] -= n; }
    else if (m[b] >= at) {
      const keepBefore = Math.max(0, at - m[a]);
      const keepAfter = Math.max(0, m[b] - end);
      const len = keepBefore + keepAfter;
      if (len <= 0) return null;
      out[a] = Math.min(m[a], at);
      out[b] = out[a] + len - 1;
    }
  }
  return out[a] === out[b] && out[axis === 'row' ? 'c1' : 'r1'] === out[axis === 'row' ? 'c2' : 'r2'] ? null : out;
}

export const setColWidth = (wb, si, c, px) => withSheet(wb, si, { ...wb.sheets[si], colW: { ...wb.sheets[si].colW, [c]: Math.max(8, Math.round(px)) } });
export const setRowHeight = (wb, si, r, px) => withSheet(wb, si, { ...wb.sheets[si], rowH: { ...wb.sheets[si].rowH, [r]: Math.max(8, Math.round(px)) } });

// ---------- Объединение ячеек ----------
export function mergeRange(wb, si, rg) {
  const sheet = wb.sheets[si];
  const inter = (m) => !(m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2);
  const merges = sheet.merges.filter((m) => !inter(m));
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return withSheet(wb, si, { ...sheet, merges });
  merges.push({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 });
  // как Excel: остаётся значение левой верхней ячейки
  const cells = new Map(sheet.cells);
  for (const [k] of sheet.cells) {
    const [r, c] = unkey(k);
    if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2 && !(r === rg.r1 && c === rg.c1)) {
      const cell = cells.get(k);
      if (cell.s) cells.set(k, { s: cell.s, v: null }); else cells.delete(k);
    }
  }
  return recalc(withSheet(wb, si, { ...sheet, merges, cells }));
}
export function unmergeRange(wb, si, rg) {
  const sheet = wb.sheets[si];
  const merges = sheet.merges.filter((m) => m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2);
  return withSheet(wb, si, { ...sheet, merges });
}
export const mergeAt = (sheet, r, c) => sheet.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;

// ---------- Сортировка ----------
// Строки диапазона упорядочиваются по столбцу col; пустые — в конце; формулы сдвигаются вместе со строкой
export function sortRange(wb, si, rg, col, asc = true) {
  const sheet = wb.sheets[si];
  const rows = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const cells = [];
    for (let c = rg.c1; c <= rg.c2; c++) cells.push(sheet.cells.get(key(r, c)) || null);
    rows.push({ r, cells, k: (sheet.cells.get(key(r, col)) || {}).v ?? null });
  }
  const rank = (v) => (v == null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  rows.sort((a, b) => {
    const ra = rank(a.k); const rb = rank(b.k);
    if (ra === 3 || rb === 3) return ra - rb;
    if (ra !== rb) return asc ? ra - rb : rb - ra;
    const x = typeof a.k === 'string' ? a.k.toLowerCase() : +a.k;
    const y = typeof b.k === 'string' ? b.k.toLowerCase() : +b.k;
    const d = typeof x === 'string' ? x.localeCompare(y, 'ru') : x - y;
    return asc ? d : -d;
  });
  const cells = new Map(sheet.cells);
  rows.forEach((row, i) => {
    const r = rg.r1 + i;
    row.cells.forEach((cell, j) => {
      const k = key(r, rg.c1 + j);
      if (!cell) { cells.delete(k); return; }
      cells.set(k, cell.f != null ? { ...cell, f: offsetFormula(cell.f, r - row.r, 0) } : cell);
    });
  });
  return recalc(withSheet(wb, si, { ...sheet, cells }));
}

// ---------- Буфер обмена ----------
// Блок ячеек: { rows, cols, cells: [[cell|null]], origin: { r, c } }
export function copyBlock(sheet, rg) {
  const cells = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    for (let c = rg.c1; c <= rg.c2; c++) row.push(sheet.cells.get(key(r, c)) || null);
    cells.push(row);
  }
  return { rows: rg.r2 - rg.r1 + 1, cols: rg.c2 - rg.c1 + 1, cells, origin: { r: rg.r1, c: rg.c1 } };
}
// Вставка блока в (r, c): формулы сдвигаются относительно исходного места, стили — как у источника
export function pasteBlock(wb, si, r0, c0, block) {
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  block.cells.forEach((row, i) => row.forEach((cell, j) => {
    const k = key(r0 + i, c0 + j);
    if (!cell) { cells.delete(k); return; }
    cells.set(k, cell.f != null ? { ...cell, f: offsetFormula(cell.f, r0 - block.origin.r, c0 - block.origin.c) } : cell);
  }));
  return recalc(withSheet(wb, si, { ...sheet, cells }));
}
// Текст из буфера (TSV, как копирует Excel) → ввод по ячейкам
export function pasteText(wb, si, r0, c0, text) {
  const rows = String(text).replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
  let out = wb;
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  rows.forEach((row, i) => row.forEach((raw, j) => {
    const k = key(r0 + i, c0 + j);
    const old = cells.get(k) || {};
    if (raw.startsWith('=') && raw.length > 1) cells.set(k, { ...old, f: fromUser(raw.slice(1)), v: null });
    else {
      const p = parseInput(raw);
      const cell = { ...old, v: p.value };
      delete cell.f;
      if (p.format && (!old.s || !old.s.numFmt)) cell.s = { ...(old.s || {}), numFmt: p.format };
      if (cell.v == null && !cell.s) cells.delete(k); else cells.set(k, cell);
    }
  }));
  out = withSheet(out, si, { ...sheet, cells });
  return recalc(out);
}
export function blockToText(block, display) {
  return block.cells.map((row, i) => row.map((cell, j) => display(cell, i, j)).join('\t')).join('\r\n');
}

// Заполнить вниз/вправо из первой строки/столбца диапазона (Ctrl+D / Ctrl+R)
export function fill(wb, si, rg, dir) {
  const sheet = wb.sheets[si];
  const cells = new Map(sheet.cells);
  if (dir === 'down') {
    for (let c = rg.c1; c <= rg.c2; c++) {
      const src = sheet.cells.get(key(rg.r1, c));
      for (let r = rg.r1 + 1; r <= rg.r2; r++) {
        if (!src) cells.delete(key(r, c));
        else cells.set(key(r, c), src.f != null ? { ...src, f: offsetFormula(src.f, r - rg.r1, 0) } : src);
      }
    }
  } else {
    for (let r = rg.r1; r <= rg.r2; r++) {
      const src = sheet.cells.get(key(r, rg.c1));
      for (let c = rg.c1 + 1; c <= rg.c2; c++) {
        if (!src) cells.delete(key(r, c));
        else cells.set(key(r, c), src.f != null ? { ...src, f: offsetFormula(src.f, 0, c - rg.c1) } : src);
      }
    }
  }
  return recalc(withSheet(wb, si, { ...sheet, cells }));
}

// Автосумма для ячейки (r, c): ближайший непрерывный блок чисел сверху (иначе слева)
export function autoSumFormula(sheet, r, c) {
  const isNum = (rr, cc) => typeof (sheet.cells.get(key(rr, cc)) || {}).v === 'number';
  let r1 = r - 1;
  while (r1 >= 0 && !isNum(r1, c) && r1 > r - 2) r1--;
  if (r1 >= 0 && isNum(r1, c)) {
    let top = r1;
    while (top - 1 >= 0 && isNum(top - 1, c)) top--;
    return `SUM(${colLetter(c)}${top + 1}:${colLetter(c)}${r1 + 1})`;
  }
  let c1 = c - 1;
  if (c1 >= 0 && isNum(r, c1)) {
    let left = c1;
    while (left - 1 >= 0 && isNum(r, left - 1)) left--;
    return `SUM(${colLetter(left)}${r + 1}:${colLetter(c1)}${r + 1})`;
  }
  return 'SUM()';
}
const colLetter = (c) => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

// ---------- Листы ----------
export function addSheet(wb) {
  let n = wb.sheets.length + 1;
  while (sheetByName(wb, `Лист${n}`)) n++;
  return { ...wb, sheets: [...wb.sheets, newSheet(`Лист${n}`)], active: wb.sheets.length };
}
// Переименование: ссылки «Старое!A1» во всех формулах — на новое имя
export function renameSheet(wb, si, name) {
  const clean = String(name).trim().replace(/[\\/?*[\]:]/g, '').slice(0, 31);
  if (!clean || wb.sheets.some((s, i) => i !== si && s.name.toLowerCase() === clean.toLowerCase())) return wb;
  const old = wb.sheets[si].name;
  const sheets = wb.sheets.map((sheet, i) => {
    const cells = new Map(sheet.cells);
    let changed = false;
    for (const [k, cell] of sheet.cells) {
      if (cell.f == null || !cell.f.includes('!')) continue;
      const f = renameInFormula(cell.f, old, clean);
      if (f !== cell.f) { cells.set(k, { ...cell, f }); changed = true; }
    }
    const next = changed ? { ...sheet, cells } : sheet;
    return i === si ? { ...next, name: clean } : next;
  });
  return recalc({ ...wb, sheets });
}
function renameInFormula(f, old, name) {
  const same = (n) => n && n.toLowerCase() === old.toLowerCase();
  return rewriteRefs(f, (t) => ({ r: t.r, c: t.c, sheet: same(t.sheet) ? name : t.sheet }));
}
export function deleteSheet(wb, si) {
  if (wb.sheets.length <= 1) return wb;
  const sheets = wb.sheets.filter((_, i) => i !== si);
  const active = Math.max(0, Math.min(wb.active > si ? wb.active - 1 : wb.active, sheets.length - 1));
  return recalc({ ...wb, sheets, active }); // ссылки на удалённый лист станут #REF! при пересчёте
}

// Границы использованных данных листа
export function usedRange(sheet) {
  let r2 = -1; let c2 = -1;
  for (const k of sheet.cells.keys()) { const [r, c] = unkey(k); if (r > r2) r2 = r; if (c > c2) c2 = c; }
  for (const m of sheet.merges) { r2 = Math.max(r2, m.r2); c2 = Math.max(c2, m.c2); }
  return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
}
