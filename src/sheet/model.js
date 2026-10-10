// Модель книги и операции над ней. Все операции возвращают новую книгу (неизменённые листы/ячейки — общие),
// поэтому отмена — просто стек предыдущих книг.
//
// Книга: { sheets: Sheet[], active, names? }
// names: [{ name, text — формула имени в хранимом виде (Данные!$C$2:$C$11, 0.2), sheet — имя листа (локальное) | null }]
// Лист:  { name, cells: Map<"r,c", Cell>, colW: { c: px }, rowH: { r: px }, merges: Range[],
//          rowInfo/colInfo: { n: { hidden, level, collapsed, xf } } — скрытие, группировка, стиль строки/столбца (xf файла),
//          src: { part, rid, name, shifts } — лист из файла (экспорт переписывает его поверх исходного XML) }
// Ячейка: { v — значение (у формулы — вычисленное), f — формула без «=» (хранимый вид), s — стиль,
//          arr — результат-массив динамической формулы («разлив»), ax — формула с поэлементными операциями (в файл —
//          как формула массива), legacy — обычная формула из файла (результат-массив → неявное пересечение),
//          cse — диапазон формулы массива Ctrl+Shift+Enter из файла }
// Лист после пересчёта: spill — Map(ячейка → значение «разлива»), spillAt — Map(якорь → диапазон), spillBlocked — Set якорей с #ПЕРЕНОС!
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
// Разобранные формулы — общий кеш для всех пересчётов (формула → AST; ошибка разбора → null)
const astCache = new Map();
function astOf(f) {
  let a = astCache.get(f);
  if (a === undefined) {
    try { a = parse(f); } catch { a = null; }
    if (astCache.size > 100_000) astCache.clear();
    astCache.set(f, a);
  }
  return a;
}
// Все формулы книги: memo по (лист, ячейка), циклические ссылки → #CIRC!
export function recalc(wb) {
  // лист по имени — словарём (обращений к ячейкам много)
  const byName = new Map();
  for (const sh of wb.sheets) { byName.set(sh.name, sh); byName.set(sh.name.toLowerCase(), sh); }
  const sheetByName = (_wb, name) => byName.get(name) || byName.get(String(name).toLowerCase());
  const memo = new Map();
  const visiting = new Set();

  // занятая часть листа (для ссылок на целые столбцы и строки)
  const extents = new Map();
  const extent = (name) => {
    const k = String(name).toLowerCase();
    if (!extents.has(k)) {
      const sh = sheetByName(wb, name);
      let r = -1; let c = -1;
      if (sh) for (const key0 of sh.cells.keys()) { const [rr, cc] = unkey(key0); if (rr > r) r = rr; if (cc > c) c = cc; }
      if (sh) for (const rg of spillOfSheet(sh.name).at.values()) { if (rg.r2 > r) r = rg.r2; if (rg.c2 > c) c = rg.c2; }
      extents.set(k, { r, c });
    }
    return extents.get(k);
  };
  // листы трёхмерной ссылки: от первого до последнего по порядку в книге
  const sheetsBetween = (a, b) => {
    const i = wb.sheets.findIndex((x) => x.name.toLowerCase() === String(a).toLowerCase());
    const j = wb.sheets.findIndex((x) => x.name.toLowerCase() === String(b).toLowerCase());
    if (i < 0 || j < 0) return null;
    return wb.sheets.slice(Math.min(i, j), Math.max(i, j) + 1).map((x) => x.name);
  };
  // именованные диапазоны и константы: сначала имя листа (локальное), потом книги
  const namesBusy = new Set();
  const nameValue = (sheetName) => (n) => {
    const list = wb.names || [];
    const lc = n.toLowerCase();
    const def = list.find((x) => x.name.toLowerCase() === lc && x.sheet && x.sheet.toLowerCase() === String(sheetName).toLowerCase())
      || list.find((x) => x.name.toLowerCase() === lc && !x.sheet);
    if (!def) return err(ERR.NAME);
    const key0 = `${def.sheet || ''}!${lc}`;
    if (namesBusy.has(key0)) return err(ERR.CIRC);
    const ast = astOf(def.text);
    if (!ast) return err(ERR.NAME);
    namesBusy.add(key0);
    const v = evaluate(ast, ctxOf(sheetName));
    namesBusy.delete(key0);
    return v;
  };
  const ctxCache = new Map();
  const ctxOf = (sheetName) => {
    if (!ctxCache.has(sheetName)) {
      ctxCache.set(sheetName, {
        get, sheet: sheetName, extent, sheetsBetween, name: nameValue(sheetName), spill: spillRef, markArray, cur: null,
        formulaAt: (sh, r, c) => { const s0 = sheetByName(wb, sh); const cell = s0 && s0.cells.get(key(r, c)); return cell && cell.f != null ? cell.f : null; },
        rowHidden: (sh, r) => { const s0 = sheetByName(wb, sh); return !!(s0 && s0.rowInfo && s0.rowInfo[r] && s0.rowInfo[r].hidden); },
        sheetIndex: (name) => wb.sheets.findIndex((x) => x.name.toLowerCase() === String(name).toLowerCase()),
        sheetCount: () => wb.sheets.length,
        nameDef: (n) => {
          const lc = String(n).toLowerCase();
          const d = (wb.names || []).find((x) => x.name.toLowerCase() === lc && x.sheet && x.sheet.toLowerCase() === String(sheetName).toLowerCase())
            || (wb.names || []).find((x) => x.name.toLowerCase() === lc && !x.sheet);
          return d ? d.text : null;
        },
      });
    }
    return ctxCache.get(sheetName);
  };
  // «Разлив» динамических массивов: по листу — значения соседних ячеек (vals), диапазоны якорей (at), заблокированные (blocked).
  // Пока не посчитано — догадка из прошлого пересчёта; после прохода раскладка уточняется, проход повторяется (не больше 4 раз).
  const emptySpill = () => ({ vals: new Map(), at: new Map(), blocked: new Set() });
  let spill = new Map(wb.sheets.map((sh) => [sh.name.toLowerCase(), sh.spill ? { vals: sh.spill, at: sh.spillAt || new Map(), blocked: sh.spillBlocked || new Set() } : emptySpill()]));
  const spillOfSheet = (name) => spill.get(String(name).toLowerCase()) || emptySpill();
  let arrFlag = false; // при вычислении текущей ячейки были поэлементные операции над массивами
  const markArray = () => { arrFlag = true; };
  const evalCell = (sheet, k, cell) => {
    const mk = `${sheet.name}!${k}`;
    if (memo.has(mk)) return memo.get(mk);
    if (visiting.has(mk)) return { top: err(ERR.CIRC) };
    visiting.add(mk);
    const ast = astOf(cell.f);
    const saved = arrFlag;
    arrFlag = false;
    const ctx = ctxOf(sheet.name);
    const prevCur = ctx.cur;
    const [cr, cc] = unkey(k);
    ctx.cur = { r: cr, c: cc }; // СТРОКА(), СТОЛБЕЦ(), ЯЧЕЙКА() без ссылки — текущая ячейка
    let v = ast ? evaluate(ast, ctx) : err(ERR.NAME);
    ctx.cur = prevCur;
    const ax = arrFlag;
    arrFlag = saved;
    let full = null;
    if (Array.isArray(v)) {
      const one = v.length === 1 && v[0] && v[0].length === 1;
      if (!v.length || !v[0] || !v[0].length) v = err(ERR.CALC);
      else if (cell.legacy || one) v = v[0][0]; // старая формула — неявное пересечение (упрощённо: левый верхний)
      else { full = v; v = spillOfSheet(sheet.name).blocked.has(k) && !cell.cse ? err(ERR.SPILL) : v[0][0]; }
    }
    if (v == null) v = 0; // пустая ссылка в формуле — 0, как в Excel
    const res = { top: v, full, ax };
    visiting.delete(mk);
    memo.set(mk, res);
    return res;
  };
  const get = (sheetName, r, c) => {
    const sheet = sheetByName(wb, sheetName);
    if (!sheet) return err(ERR.REF);
    const k = key(r, c);
    const cell = sheet.cells.get(k);
    if (!cell || (cell.f == null && cell.v == null)) {
      const g = spillOfSheet(sheet.name).vals.get(k);
      return g === undefined ? null : g;
    }
    if (cell.f == null) return cell.v ?? null;
    return evalCell(sheet, k, cell).top;
  };
  // A1# — весь «разлив» якоря A1
  const spillRef = (sheetName, r, c) => {
    const sheet = sheetByName(wb, sheetName);
    const cell = sheet && sheet.cells.get(key(r, c));
    if (!cell || cell.f == null) return err(ERR.REF);
    const res = evalCell(sheet, key(r, c), cell);
    if (isErr(res.top) && res.top.error === ERR.SPILL) return res.top;
    return res.full || [[res.top]];
  };
  // раскладка «разливов» после прохода
  const resolve = () => {
    const out = new Map();
    for (const sheet of wb.sheets) {
      const sp = emptySpill();
      const anchors = [];
      for (const [k, cell] of sheet.cells) {
        if (cell.f == null) continue;
        const res = memo.get(`${sheet.name}!${k}`);
        if (res && res.full) anchors.push([k, cell, res.full]);
      }
      anchors.sort((a, b) => { const [ra, ca] = unkey(a[0]); const [rb, cb] = unkey(b[0]); return ra - rb || ca - cb; });
      const claimed = new Set();
      for (const [k, cell, full] of anchors) {
        const [r0, c0] = unkey(k);
        // формула массива Ctrl+Shift+Enter — свой фиксированный диапазон (лишнее — #Н/Д); динамическая — по размеру результата
        let rows = full.length; let cols = full[0].length;
        if (cell.cse) { const rg = parseCse(cell.cse); if (rg) { rows = rg.r2 - rg.r1 + 1; cols = rg.c2 - rg.c1 + 1; } }
        const keys = [];
        let blocked = false;
        for (let i = 0; i < rows && !blocked; i++) {
          for (let j = 0; j < cols; j++) {
            if (!i && !j) continue;
            const kk = key(r0 + i, c0 + j);
            const other = sheet.cells.get(kk);
            if (!cell.cse && ((other && (other.v != null || other.f != null)) || claimed.has(kk) || (sheet.merges.length && mergeAt(sheet, r0 + i, c0 + j)))) { blocked = true; break; }
            keys.push([kk, i, j]);
          }
        }
        if (blocked) { sp.blocked.add(k); continue; }
        claimed.add(k);
        for (const [kk, i, j] of keys) {
          claimed.add(kk);
          const v = full[i] && full[i][j] !== undefined ? full[i][j] : err(ERR.NA);
          sp.vals.set(kk, v == null ? 0 : v);
        }
        sp.at.set(k, { r1: r0, c1: c0, r2: r0 + rows - 1, c2: c0 + cols - 1 });
      }
      out.set(sheet.name.toLowerCase(), sp);
    }
    return out;
  };
  const sameSpill = (a, b) => {
    for (const [name, x] of a) {
      const y = b.get(name);
      if (!y || x.vals.size !== y.vals.size || x.blocked.size !== y.blocked.size || x.at.size !== y.at.size) return false;
      for (const k of x.blocked) if (!y.blocked.has(k)) return false;
      for (const [k, v] of x.vals) if (!y.vals.has(k) || !sameValue(v, y.vals.get(k))) return false;
      for (const [k, rg] of x.at) { const o = y.at.get(k); if (!o || o.r2 !== rg.r2 || o.c2 !== rg.c2) return false; }
    }
    return true;
  };
  for (let pass = 0; pass < 4; pass++) {
    memo.clear();
    extents.clear();
    for (const sheet of wb.sheets) for (const [k, cell] of sheet.cells) if (cell.f != null) evalCell(sheet, k, cell);
    const next = resolve();
    const done = sameSpill(next, spill) && sameSpill(spill, next);
    spill = next;
    if (done) break; // раскладка не изменилась — значения последнего прохода верны
  }
  const sheets = wb.sheets.map((sheet) => {
    let changed = false;
    const cells = new Map(sheet.cells);
    for (const [k, cell] of sheet.cells) {
      if (cell.f == null) continue;
      const res = evalCell(sheet, k, cell);
      const arr = res.full && !cell.legacy ? res.full : undefined;
      const ax = res.ax && !cell.legacy ? true : undefined;
      if (!sameValue(res.top, cell.v) || !sameArr(arr, cell.arr) || ax !== cell.ax) {
        const n = { ...cell, v: res.top };
        if (arr) n.arr = arr; else delete n.arr;
        if (ax) n.ax = true; else delete n.ax;
        cells.set(k, n);
        changed = true;
      }
    }
    const sp = spillOfSheet(sheet.name);
    const before = sheet.spill ? { vals: sheet.spill, at: sheet.spillAt || new Map(), blocked: sheet.spillBlocked || new Set() } : emptySpill();
    const spillChanged = !(sameSpill(new Map([['x', sp]]), new Map([['x', before]])) && sameSpill(new Map([['x', before]]), new Map([['x', sp]])));
    if (!changed && !spillChanged) return sheet;
    const next = { ...sheet, cells };
    if (sp.vals.size || sp.at.size || sp.blocked.size) { next.spill = sp.vals; next.spillAt = sp.at; next.spillBlocked = sp.blocked; }
    else { delete next.spill; delete next.spillAt; delete next.spillBlocked; }
    return next;
  });
  return { ...wb, sheets };
}
const sameArr = (a, b) => (a === b || (!!a && !!b && JSON.stringify(a) === JSON.stringify(b)));
// диапазон формулы массива Ctrl+Shift+Enter: «A1:B3»
const parseCse = (ref) => {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(String(ref));
  if (!m) return null;
  const col = (t) => [...t].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { r1: +m[2] - 1, c1: col(m[1]), r2: +(m[4] || m[2]) - 1, c2: col(m[3] || m[1]) };
};
// Значение ячейки с учётом «разлива» (для показа, копирования, итогов в строке состояния)
export function valueAt(sheet, r, c) {
  const cell = sheet.cells.get(key(r, c));
  if (cell && (cell.v != null || cell.f != null)) return cell.v ?? null;
  return sheet.spill ? (sheet.spill.get(key(r, c)) ?? null) : null;
}
// Якорь «разлива», которому принадлежит ячейка: { r, c, range } или null
export function spillAnchorAt(sheet, r, c) {
  if (!sheet.spillAt) return null;
  for (const [k, rg] of sheet.spillAt) if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { const [ar, ac] = unkey(k); return { r: ar, c: ac, range: rg }; }
  return null;
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
    for (const n of ['arr', 'ax', 'legacy', 'cse']) delete cell[n];
  } else {
    const p = parseInput(s);
    cell = { ...old, v: p.value };
    for (const n of ['f', 'arr', 'ax', 'legacy', 'cse']) delete cell[n];
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
      // область печати и сквозные строки/столбцы
      const [k1, k2] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
      if (sheet.printArea) {
        const span = shiftSpan(sheet.printArea[k1], sheet.printArea[k2], at, count);
        next.printArea = span ? { ...sheet.printArea, [k1]: span[0], [k2]: span[1] } : null;
        if (!span) delete next.printArea;
      }
      const tk = axis === 'row' ? 'rows' : 'cols';
      if (sheet.printTitles && sheet.printTitles[tk]) {
        const span = shiftSpan(sheet.printTitles[tk][0], sheet.printTitles[tk][1], at, count);
        next.printTitles = { ...sheet.printTitles, [tk]: span };
      }
      // лист из файла: сдвиг запоминается — экспорт сдвинет диапазоны условного форматирования, проверки, ссылок
      if (sheet.src) next.src = { ...sheet.src, shifts: [...(sheet.src.shifts || []), { axis, at, count }] };
    }
    return next;
  });
  const names = wb.names && wb.names.map((n) => {
    const text = shiftFormula(n.text, { sheet: name, current: null, axis, at, count });
    return text === n.text ? n : { ...n, text };
  });
  return recalc({ ...wb, sheets, ...(names ? { names } : {}) });
}
// Отрезок a..b после вставки (count > 0) / удаления строк или столбцов; удалён целиком — null
function shiftSpan(a, b, at, count) {
  if (count > 0) return [a >= at ? a + count : a, b >= at ? b + count : b];
  const n = -count; const end = at + n - 1;
  if (a >= at && b <= end) return null;
  return [a > end ? a - n : a >= at ? at : a, b > end ? b - n : b >= at ? at - 1 : b];
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

// ---------- Закрепление, скрытие, группировка ----------
// Закрепление областей: { rows, cols } — число закреплённых строк сверху и столбцов слева; null — снять
export function setFreeze(wb, si, freeze) {
  const f = freeze && (freeze.rows > 0 || freeze.cols > 0) ? { rows: Math.max(0, freeze.rows | 0), cols: Math.max(0, freeze.cols | 0) } : null;
  const sheet = { ...wb.sheets[si], freeze: f };
  if (!f) delete sheet.freeze;
  return withSheet(wb, si, sheet);
}

const infoKey = (axis) => (axis === 'row' ? 'rowInfo' : 'colInfo');
// Изменить свойства строк/столбцов a..b: fn(info) → новый info (пустой — удаляется)
function mapLines(wb, si, axis, a, b, fn) {
  const sheet = wb.sheets[si];
  const k = infoKey(axis);
  const info = { ...(sheet[k] || {}) };
  for (let i = a; i <= b; i++) {
    const next = fn({ ...(info[i] || {}) }, i);
    for (const n of Object.keys(next)) if (next[n] == null || next[n] === false || next[n] === 0) delete next[n];
    if (Object.keys(next).length) info[i] = next; else delete info[i];
  }
  return withSheet(wb, si, { ...sheet, [k]: info });
}
export const isHidden = (sheet, axis, i) => !!((sheet[infoKey(axis)] || {})[i] || {}).hidden;
export const levelOf = (sheet, axis, i) => ((sheet[infoKey(axis)] || {})[i] || {}).level || 0;

// Скрыть / отобразить строки или столбцы a..b
export const setHidden = (wb, si, axis, a, b, hidden) => mapLines(wb, si, axis, a, b, (x) => ({ ...x, hidden }));

// Группировать (delta = 1) / разгруппировать (delta = -1); уровней не больше 7, как в Excel
export function groupLines(wb, si, axis, a, b, delta) {
  let out = mapLines(wb, si, axis, a, b, (x) => {
    const level = Math.max(0, Math.min(7, (x.level || 0) + delta));
    // разгруппированная строка показывается (иначе её уже не развернуть)
    return { ...x, level, hidden: delta < 0 && level < (x.level || 0) && !level ? false : x.hidden };
  });
  // признак «свёрнуто» у итоговой строки после группы
  if (delta < 0) out = mapLines(out, si, axis, b + 1, b + 1, (x) => ({ ...x, collapsed: false }));
  return out;
}

// Группы уровня level: [{ a, b }] — сплошные участки строк с уровнем ≥ level; итоговая строка — b + 1
export function outlineRuns(sheet, axis, level) {
  const info = sheet[infoKey(axis)] || {};
  const idx = Object.keys(info).map(Number).filter((i) => (info[i].level || 0) >= level).sort((x, y) => x - y);
  const runs = [];
  for (const i of idx) {
    const last = runs[runs.length - 1];
    if (last && last.b === i - 1) last.b = i; else runs.push({ a: i, b: i });
  }
  return runs;
}
export function maxLevel(sheet, axis) {
  let m = 0;
  for (const x of Object.values(sheet[infoKey(axis)] || {})) if (x.level > m) m = x.level;
  return m;
}

// Свернуть группу a..b уровня level: строки скрываются, итоговая помечается «свёрнуто»
export function collapseGroup(wb, si, axis, a, b) {
  const out = setHidden(wb, si, axis, a, b, true);
  return mapLines(out, si, axis, b + 1, b + 1, (x) => ({ ...x, collapsed: true }));
}
// Развернуть: показываются строки этого уровня, вложенные группы — если сами не свёрнуты
export function expandGroup(wb, si, axis, a, b, level) {
  let out = mapLines(wb, si, axis, b + 1, b + 1, (x) => ({ ...x, collapsed: false }));
  const show = (w, x, y, lvl) => {
    let o = mapLines(w, si, axis, x, y, (info) => ((info.level || 0) === lvl ? { ...info, hidden: false } : info));
    for (const run of outlineRuns(o.sheets[si], axis, lvl + 1)) {
      if (run.a < x || run.b > y) continue;
      const summary = ((o.sheets[si][infoKey(axis)] || {})[run.b + 1] || {});
      if (!summary.collapsed) o = show(o, run.a, run.b, lvl + 1);
    }
    return o;
  };
  out = show(out, a, b, level);
  return out;
}
// Кнопки уровней структуры (1, 2, 3…): показать уровни меньше n, глубже — свернуть
export function showLevels(wb, si, axis, n) {
  let out = wb;
  const max = maxLevel(wb.sheets[si], axis);
  for (let lvl = max; lvl >= 1; lvl--) {
    for (const run of outlineRuns(out.sheets[si], axis, lvl)) {
      out = lvl >= n ? collapseGroup(out, si, axis, run.a, run.b) : mapLines(out, si, axis, run.b + 1, run.b + 1, (x) => ({ ...x, collapsed: false }));
    }
  }
  // строки уровней < n — видимы
  const info = out.sheets[si][infoKey(axis)] || {};
  for (const i of Object.keys(info).map(Number)) {
    const lvl = info[i].level || 0;
    if (lvl && lvl < n) out = mapLines(out, si, axis, i, i, (x) => ({ ...x, hidden: false }));
  }
  return out;
}
// Группа, содержащая строку i, самого глубокого уровня: { a, b, level } или null
export function groupAt(sheet, axis, i) {
  for (let lvl = levelOf(sheet, axis, i); lvl >= 1; lvl--) {
    const run = outlineRuns(sheet, axis, lvl).find((x) => i >= x.a && i <= x.b);
    if (run) return { ...run, level: lvl };
  }
  return null;
}

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
    for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = sheet.cells.get(key(r, c)) || null;
      // ячейка «разлива» копируется значением (формула — только у якоря), как в Excel
      const g = (!cell || (cell.v == null && cell.f == null)) && sheet.spill ? sheet.spill.get(key(r, c)) : undefined;
      row.push(g !== undefined ? { ...(cell || {}), v: g } : cell);
    }
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

// ---------- Имена ----------
// Допустимое имя (как в Excel): буква, «_» или «\» в начале, без пробелов, не адрес ячейки (A1, R1C1) и не ИСТИНА/ЛОЖЬ
export function isValidName(name) {
  const n = String(name || '');
  if (!/^[A-Za-zА-Яа-яЁё_\\][\wА-Яа-яЁё.]*$/.test(n) || n.length > 255) return false;
  if (/^[A-Za-z]{1,3}\d{1,7}$/.test(n) || /^[RrCc]\d*$/.test(n) || /^R\d*C\d*$/i.test(n)) return false;
  return !/^(TRUE|FALSE|ИСТИНА|ЛОЖЬ)$/i.test(n);
}
// Заменить список имён книги (диспетчер имён) — с пересчётом
export function setNames(wb, names) {
  const next = { ...wb, names };
  if (!names || !names.length) delete next.names;
  return recalc(next);
}
// Имя → диапазон для перехода (поле имени): { sheet, r1, c1, r2, c2 } или null
export function nameTarget(wb, name, sheetName) {
  const lc = String(name).toLowerCase();
  const def = (wb.names || []).find((x) => x.name.toLowerCase() === lc && x.sheet && x.sheet.toLowerCase() === String(sheetName).toLowerCase())
    || (wb.names || []).find((x) => x.name.toLowerCase() === lc && !x.sheet);
  if (!def) return null;
  let ast;
  try { ast = parse(def.text); } catch { return null; }
  if (ast.t === 'ref') return { sheet: ast.sheet || sheetName, r1: ast.r, c1: ast.c, r2: ast.r, c2: ast.c };
  if (ast.t === 'range' && !ast.sheet2) return { sheet: ast.sheet || sheetName, r1: ast.r1, c1: ast.c1, r2: Math.min(ast.r2, 1048575), c2: Math.min(ast.c2, 16383) };
  return null;
}

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
  const names = wb.names && wb.names.map((n) => {
    const text = n.text.includes('!') ? renameInFormula(n.text, old, clean) : n.text;
    const sheet = n.sheet && n.sheet.toLowerCase() === old.toLowerCase() ? clean : n.sheet;
    return text === n.text && sheet === n.sheet ? n : { ...n, text, sheet };
  });
  return recalc({ ...wb, sheets, ...(names ? { names } : {}) });
}
function renameInFormula(f, old, name) {
  const same = (n) => n && n.toLowerCase() === old.toLowerCase();
  return rewriteRefs(f, (t) => ({ r: t.r, c: t.c, sheet: same(t.sheet) ? name : t.sheet }));
}
export function deleteSheet(wb, si) {
  if (wb.sheets.length <= 1) return wb;
  const sheets = wb.sheets.filter((_, i) => i !== si);
  const active = Math.max(0, Math.min(wb.active > si ? wb.active - 1 : wb.active, sheets.length - 1));
  // локальные имена удалённого листа исчезают; ссылки на него станут #REF! при пересчёте
  const gone = wb.sheets[si].name.toLowerCase();
  const names = wb.names && wb.names.filter((n) => !(n.sheet && n.sheet.toLowerCase() === gone));
  return recalc({ ...wb, sheets, active, ...(names ? { names } : {}) });
}

// Границы использованных данных листа
export function usedRange(sheet) {
  let r2 = -1; let c2 = -1;
  for (const k of sheet.cells.keys()) { const [r, c] = unkey(k); if (r > r2) r2 = r; if (c > c2) c2 = c; }
  for (const m of sheet.merges) { r2 = Math.max(r2, m.r2); c2 = Math.max(c2, m.c2); }
  return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
}
