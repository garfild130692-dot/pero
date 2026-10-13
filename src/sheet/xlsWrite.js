// Модель книги → .xls (Excel 97–2003, BIFF8) — своим кодом по спецификации [MS-XLS] (этап 19.3).
// Глобальная часть (шрифты, форматы, XF, палитра под используемые цвета, листы, внешние ссылки, имена, общие строки
// SST + EXTSST) и часть каждого листа (параметры страницы, столбцы, строки, ячейки, формулы, окно, объединения).
// Формулы: текст модели → AST (parse) → байт-код Ptg в обратной польской записи; функции из таблицы Ftab — штатно,
// функции новее Excel 2003 — как пишет сам Excel: скрытое имя «_xlfn.X» и вызов функции 255; функции пакета анализа —
// внешние имена надстройки. Пределы BIFF8: 65 536 строк, 256 столбцов, 56 цветов палитры.
import { writeCfb } from '../formats/cfb.js';
import { key } from './address.js';
import { BUILTIN_FMT } from './format.js';
import { FUTURE, XLWS, parse } from './formula.js';
import { FTAB_BY_NAME, UDF } from './xlsFtab.js';
import { PARAM_CLASS } from './xlsClasses.js';
import { DEFAULT_COL_W, DEFAULT_ROW_H } from './model.js';
import { INDEXED, pxToWidth } from './xlsx.js';
import { EXCEL_RU } from './fnnames.js';

const MAXR = 65536; const MAXC = 256;
const ERR_CODE = { '#NULL!': 0x00, '#DIV/0!': 0x07, '#VALUE!': 0x0F, '#REF!': 0x17, '#NAME?': 0x1D, '#NUM!': 0x24, '#N/A': 0x2A, '#GETTING_DATA': 0x2B };
const BORDER_DG = { thin: 1, medium: 2, dashed: 3, dotted: 4, thick: 5, double: 6, hair: 7, mediumDashed: 8, dashDot: 9, mediumDashDot: 10, dashDotDot: 11, mediumDashDotDot: 12, slantDashDot: 13 };

// ---------- Байты ----------
class Buf {
  constructor() { this.a = []; }
  u8(v) { this.a.push(v & 0xFF); return this; }
  u16(v) { this.a.push(v & 0xFF, (v >> 8) & 0xFF); return this; }
  u32(v) { this.u16(v & 0xFFFF); this.u16((v >>> 16) & 0xFFFF); return this; }
  f64(v) { const b = new Uint8Array(8); new DataView(b.buffer).setFloat64(0, v, true); for (const x of b) this.a.push(x); return this; }
  bytes(b) { for (const x of b) this.a.push(x); return this; }
  get length() { return this.a.length; }
  out() { return Uint8Array.from(this.a); }
}
const wide = (s) => [...s].some((ch) => ch.charCodeAt(0) > 0xFF);
const charBytes = (s, high) => { const b = new Buf(); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); if (high) b.u16(c); else b.u8(c); } return b.a; };
// XLUnicodeString: длина u16, флаги, символы
const xlString = (s) => { const h = wide(s); return new Buf().u16(s.length).u8(h ? 1 : 0).bytes(charBytes(s, h)).a; };
// ShortXLUnicodeString: длина u8, флаги, символы
const shortString = (s) => { const t = s.slice(0, 255); const h = wide(t); return new Buf().u8(t.length).u8(h ? 1 : 0).bytes(charBytes(t, h)).a; };

// Поток записей; длинные данные — продолжением CONTINUE (не больше 8224 байт в записи)
class Stream {
  constructor() { this.parts = []; this.len = 0; }
  rec(type, data) {
    const d = data instanceof Uint8Array ? data : Uint8Array.from(data instanceof Buf ? data.a : data);
    for (let o = 0, first = true; first || o < d.length; first = false) {
      const n = Math.min(8224, d.length - o);
      const h = new Uint8Array(4); h[0] = (first ? type : 0x3C) & 0xFF; h[1] = (first ? type : 0x3C) >> 8; h[2] = n & 0xFF; h[3] = n >> 8;
      this.parts.push(h, d.subarray(o, o + n)); this.len += 4 + n; o += n;
    }
  }
  out() { const b = new Uint8Array(this.len); let o = 0; for (const p of this.parts) { b.set(p, o); o += p.length; } return b; }
}

// ---------- Формулы: AST → Ptg ----------
const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
const BINOP = { '+': 0x03, '-': 0x04, '*': 0x05, '/': 0x06, '^': 0x07, '&': 0x08, '<': 0x09, '<=': 0x0A, '=': 0x0B, '>=': 0x0C, '>': 0x0D, '<>': 0x0E };
const REF_FN = new Set(['OFFSET', 'INDIRECT', 'INDEX', 'CHOOSE', 'IF']); // могут вернуть ссылку — там, где ждут ссылку, класс R
// функции, возвращающие массив (в Excel 365 — «разлив»; в .xls такую формулу Excel пишет формулой массива)
const DYN_FN = new Set(['FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'TRANSPOSE', 'MMULT', 'MINVERSE', 'MUNIT',
  'FREQUENCY', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'CHOOSEROWS', 'CHOOSECOLS',
  'EXPAND', 'LINEST', 'LOGEST', 'TREND', 'GROWTH', 'MAP', 'SCAN', 'MAKEARRAY', 'BYROW', 'BYCOL', 'GROUPBY', 'PIVOTBY']);
class Unsupported extends Error {}

// ctx: { sheetIndex(name) → номер листа | -1, own (номер текущего листа), xti(first, last) → номер XTI,
//        nameIndex(имя) → номер имени (1…) | 0, futureName(имя функции) → номер скрытого имени _xlfn.,
//        addin(имя функции) → { ixti, index } — внешнее имя надстройки; three — ссылки всегда трёхмерные (имена) }
function compile(ast, ctx) {
  const ce = new Buf(); const cb = new Buf();
  // формуле нужна запись формулой массива: диапазон там, где ждут значение (иначе неявное пересечение), или функция,
  // возвращающая массив, вне параметра-массива — так Excel 365 решает, «динамическая» ли формула
  let needArray = false; // диапазон на месте значения (у формулы legacy так задумано — неявное пересечение)
  let needDyn = false; // функция-массив — формула «динамическая» в любом случае
  const arrayCtx = ctx.mode === 'A'; // формула массива: всё, кроме параметров-ссылок, — массивы
  const cls = (base, mode) => base + (mode === 'R' ? 0 : mode === 'A' ? 0x40 : 0x20);
  const loc = (r, c, absR, absC) => { ce.u16(r); ce.u16((c & 0x3FFF) | (absC ? 0 : 0x4000) | (absR ? 0 : 0x8000)); };
  const sheetOf = (name) => {
    if (!name) return ctx.own;
    const i = ctx.sheetIndex(name);
    if (i < 0) throw new Unsupported('ссылка на другую книгу');
    return i;
  };
  // mode — класс, который ждёт родитель: V (значение), R (ссылка), A (массив), null — неизвестно (решает сам операнд)
  const emit = (n, mode) => {
    switch (n.t) {
      case 'num': if (Number.isInteger(n.v) && n.v >= 0 && n.v <= 0xFFFF) ce.u8(0x1E).u16(n.v); else ce.u8(0x1F).f64(n.v); return;
      case 'str': if (n.v.length > 255) throw new Unsupported('строка длиннее 255'); ce.u8(0x17).bytes(shortString(n.v)); return;
      case 'bool': ce.u8(0x1D).u8(n.v ? 1 : 0); return;
      case 'err': ce.u8(0x1C).u8(ERR_CODE[n.v] ?? 0x2A); return;
      case 'empty': ce.u8(0x16); return;
      case 'ref': case 'spill': {
        const m = mode || 'V';
        if (n.r >= MAXR || n.c >= MAXC) { ce.u8(cls(0x2A, m)).u32(0); return; } // за пределами листа BIFF8 — #ССЫЛКА!
        const [absR, absC] = n.abs || [false, false];
        if (n.sheet || ctx.three) { const s = sheetOf(n.sheet); ce.u8(cls(0x3A, m)).u16(ctx.xti(s, s)); } else ce.u8(cls(0x24, m));
        loc(n.r, n.c, absR, absC);
        return;
      }
      case 'range': {
        const r2 = Math.min(n.r2, MAXR - 1); const c2 = Math.min(n.c2, MAXC - 1);
        const m = mode || 'R';
        if (m === 'V' && !(n.r1 === n.r2 && n.c1 === n.c2)) needArray = true;
        if (n.r1 >= MAXR || n.c1 >= MAXC) { ce.u8(cls(0x2B, m)).u32(0).u32(0); return; }
        const [aR1, aC1, aR2, aC2] = n.abs || [false, false, false, false];
        if (n.sheet2 && n.r1 === n.r2 && n.c1 === n.c2 && !n.whole) { // Лист1:Лист3!A1 — трёхмерная ссылка на ячейку
          const s1 = sheetOf(n.sheet); const s2 = sheetOf(n.sheet2);
          ce.u8(cls(0x3A, m)).u16(ctx.xti(Math.min(s1, s2), Math.max(s1, s2))); loc(n.r1, n.c1, aR1, aC1);
          return;
        }
        if (n.sheet || n.sheet2 || ctx.three) {
          const s1 = sheetOf(n.sheet); const s2 = n.sheet2 ? sheetOf(n.sheet2) : s1;
          ce.u8(cls(0x3B, m)).u16(ctx.xti(Math.min(s1, s2), Math.max(s1, s2)));
        } else ce.u8(cls(0x25, m));
        ce.u16(n.r1); ce.u16(r2);
        ce.u16((n.c1 & 0x3FFF) | (aC1 ? 0 : 0x4000) | (aR1 ? 0 : 0x8000)); ce.u16((c2 & 0x3FFF) | (aC2 ? 0 : 0x4000) | (aR2 ? 0 : 0x8000));
        return;
      }
      case 'name': {
        const ix = ctx.nameIndex(n.v);
        if (!ix) throw new Unsupported(`имя ${n.v}`);
        ce.u8(cls(0x23, mode || 'V')).u32(ix);
        return;
      }
      case 'array': {
        ce.u8(0x60).bytes([0, 0, 0, 0, 0, 0, 0]);
        cb.u8(n.v[0].length - 1).u16(n.v.length - 1);
        for (const row of n.v) for (const v of row) {
          if (typeof v === 'number') cb.u8(0x01).f64(v);
          else if (typeof v === 'string') cb.u8(0x02).bytes(xlString(v));
          else if (typeof v === 'boolean') cb.u8(0x04).u8(v ? 1 : 0).bytes([0, 0, 0, 0, 0, 0, 0]);
          else if (v && v.error) cb.u8(0x10).u8(ERR_CODE[v.error] ?? 0x2A).bytes([0, 0, 0, 0, 0, 0, 0]);
          else cb.u8(0x00).bytes([0, 0, 0, 0, 0, 0, 0, 0]);
        }
        return;
      }
      case 'un': emit(n.a, arrayCtx || mode === 'A' ? 'A' : 'V'); if (paren(n.a, 6)) ce.u8(0x15); ce.u8(n.op === '-' ? 0x13 : 0x12); return;
      case 'pct': emit(n.a, arrayCtx || mode === 'A' ? 'A' : 'V'); if (paren(n.a, 6)) ce.u8(0x15); ce.u8(0x14); return;
      case 'bin': {
        if (!BINOP[n.op]) throw new Unsupported(n.op);
        const p = PREC[n.op];
        const om = arrayCtx || mode === 'A' ? 'A' : 'V'; // операнды операторов — значения (в формуле массива — массивы)
        emit(n.a, om); if (paren(n.a, p)) ce.u8(0x15);
        emit(n.b, om); if (paren(n.b, p + (n.op === '^' ? 0 : 1))) ce.u8(0x15);
        ce.u8(BINOP[n.op]);
        return;
      }
      case 'fn': {
        const name = n.name;
        if (name === 'LET' || name === 'LAMBDA') throw new Unsupported(name); // в .xls их нет (Excel пишет #ЗНАЧ!)
        const f = FTAB_BY_NAME.get(name);
        // класс аргумента — как пишет сам Excel (xlsClasses.js); в формуле массива значения — массивами
        const pc = PARAM_CLASS[name];
        const argMode = (i) => { const c = pc ? pc[Math.min(i, pc.length - 1)] : null; return (arrayCtx || mode === 'A') && c !== 'R' ? 'A' : c; };
        // токен функции (как пишет Excel): где ждут ссылку и функция может её вернуть — R; аргумент формулы массива и
        // функция, возвращающая массив, — A; иначе V. Функция-массив там, где ждут значение, — формула массива (ARRAY)
        const fnMode = mode === 'R' && REF_FN.has(name) ? 'R' : mode === 'A' || DYN_FN.has(name) || (arrayCtx && mode === 'R') ? 'A' : 'V';
        if (DYN_FN.has(name)) needDyn = true;
        if ((name === 'ROW' || name === 'COLUMN') && n.args[0] && n.args[0].t === 'range' && !(n.args[0].r1 === n.args[0].r2 && n.args[0].c1 === n.args[0].c2)) needArray = true;
        if (f && f.iftab !== UDF) {
          n.args.forEach((a, i) => emit(a, argMode(i)));
          // СУММ одного аргумента Excel пишет tAttrSum
          if (name === 'SUM' && n.args.length === 1 && fnMode !== 'R') { ce.u8(0x19).u8(0x10).u16(0); return; }
          if (f.argc != null && f.argc === n.args.length) ce.u8(cls(0x21, fnMode)).u16(f.iftab);
          else ce.u8(cls(0x22, fnMode)).u8(n.args.length).u16(f.iftab);
          return;
        }
        // функция по имени: новее Excel 2003 — скрытое имя _xlfn.X; пакета анализа — внешнее имя надстройки
        // имя функции — токен класса R (так пишет Excel)
        if (ctx.atp(name) === 'future') ce.u8(0x23).u32(ctx.futureName(`_xlfn.${XLWS.has(name) ? '_xlws.' : ''}${name}`));
        else { const x = ctx.addin(name); ce.u8(0x39).u16(x.ixti).u16(x.index).u16(0); }
        n.args.forEach((a, i) => emit(a, argMode(i)));
        ce.u8(cls(0x22, fnMode)).u8(n.args.length + 1).u16(UDF);
        return;
      }
      default: throw new Unsupported(n.t);
    }
  };
  // скобки: у дочернего выражения приоритет ниже нужного — PtgParen (для показа формулы в Excel)
  const prec = (n) => (n.t === 'bin' ? PREC[n.op] : 7);
  const paren = (child, need) => prec(child) < need;
  // верхняя функция формулы массива — класса V (так пишет Excel), её аргументы — массивы
  emit(ast, ctx.mode === 'A' && ast.t === 'fn' ? 'V' : ctx.mode || 'V');
  return { rgce: ce.out(), rgcb: cb.out(), needArray, needDyn };
}

// ---------- Книга ----------
// → { bytes, lost: { что: сколько } } — то, что в .xls не помещается или не записывается
export function exportXls(wb, opts = {}) {
  const lost = {};
  const lose = (k) => { lost[k] = (lost[k] || 0) + 1; };
  const sheets = wb.sheets;
  const lcIndex = new Map(sheets.map((s, i) => [s.name.toLowerCase(), i]));

  // --- внешние ссылки и имена (заполняются при компиляции формул) ---
  const xtis = []; const xtiKey = new Map();
  const xti = (book, first, last) => { const k = `${book}:${first}:${last}`; if (!xtiKey.has(k)) { xtiKey.set(k, xtis.length); xtis.push([book, first, last]); } return xtiKey.get(k); };
  const SELF = 0; const ADDIN = 1; let addinUsed = false;
  const addinNames = []; const addinIndex = new Map();
  const userNames = (wb.names || []).filter((n) => /^[A-Za-z_Ѐ-ӿ\\][\w.Ѐ-ӿ]*$/.test(n.name));
  const nameRecs = []; // { name, builtin (код), itab, formula? (AST), hidden, fn }
  for (const n of userNames) nameRecs.push({ name: n.name, itab: n.sheet ? (lcIndex.get(n.sheet.toLowerCase()) ?? -1) + 1 : 0, text: n.text });
  const futureIdx = new Map();
  const ctxBase = {
    sheetIndex: (name) => lcIndex.get(String(name).toLowerCase()) ?? -1,
    xti: (a, b) => xti(SELF, a, b),
    nameIndex: (name) => { const lc = String(name).toLowerCase(); const i = nameRecs.findIndex((x) => !x.builtin && !x.fn && x.name.toLowerCase() === lc); return i + 1; },
    futureName: (name) => {
      if (!futureIdx.has(name)) { nameRecs.push({ name, fn: true, hidden: true, itab: 0 }); futureIdx.set(name, nameRecs.length); }
      return futureIdx.get(name);
    },
    // функция не из таблицы Ftab: 'future' — скрытое имя _xlfn.X, иначе — имя надстройки (пакет анализа)
    // проверено Excel (COM): скрытое имя _xlfn.X понимают все функции не из Ftab, включая пакет анализа, в любом языке
    // Excel; имена надстройки — только локализованные (русский Excel пишет «ДАТАМЕС») — оставлены для opts.atp 'ru'/'en'
    atp: (name) => (!opts.atp || opts.atp === 'future' || FUTURE.has(name) ? 'future' : 'addin'),
    addin: (name0) => {
      const name = opts.atp === 'ru' ? EXCEL_RU[name0] || name0 : name0;
      addinUsed = true;
      if (!addinIndex.has(name)) { addinNames.push(name); addinIndex.set(name, addinNames.length); }
      return { ixti: xti(ADDIN, -2, -2), index: addinIndex.get(name) };
    },
  };
  const tryCompile = (text, extra) => {
    let ast;
    try { ast = parse(text); } catch { return null; }
    if (!ast) return null;
    try { return compile(ast, { ...ctxBase, ...extra }); } catch (e) { if (e instanceof Unsupported) return null; throw e; }
  };
  // область печати и сквозные строки — встроенные имена листа (Print_Area = 6, Print_Titles = 7)
  sheets.forEach((s, i) => {
    const q = `'${s.name.replace(/'/g, "''")}'`;
    const col = (c) => { let n = c + 1; let r = ''; while (n) { const m = (n - 1) % 26; r = String.fromCharCode(65 + m) + r; n = Math.floor((n - 1) / 26); } return r; };
    if (s.printArea) { const a = s.printArea; nameRecs.push({ builtin: 6, itab: i + 1, text: `${q}!$${col(a.c1)}$${a.r1 + 1}:$${col(a.c2)}$${a.r2 + 1}` }); }
    if (s.printTitles && (s.printTitles.rows || s.printTitles.cols)) {
      const parts = [];
      if (s.printTitles.cols) parts.push(`${q}!$${col(s.printTitles.cols[0])}:$${col(s.printTitles.cols[1])}`);
      if (s.printTitles.rows) parts.push(`${q}!$${s.printTitles.rows[0] + 1}:$${s.printTitles.rows[1] + 1}`);
      nameRecs.push({ builtin: 7, itab: i + 1, text: parts.length > 1 ? `(${parts.join(',')})` : parts[0] });
    }
  });

  // --- стили: шрифты, форматы, палитра, XF ---
  const palette = INDEXED.slice(8, 64); // 56 цветов, номера 8…63
  const used = new Map(); // цвет → номер
  const free = []; // номера, которые можно занять своим цветом (с конца)
  for (let i = 55; i >= 0; i--) free.push(i);
  let paletteChanged = false;
  const rgb = (hex) => [1, 3, 5].map((k) => parseInt(hex.slice(k, k + 2), 16));
  const icv = (c) => {
    if (!c) return null;
    const hex = c.slice(1).toUpperCase();
    if (used.has(hex)) return used.get(hex);
    let i = palette.indexOf(hex);
    if (i < 0) {
      // свободный номер палитры (не занятый другим использованным цветом) — переписать; иначе ближайший
      let slot;
      while (free.length && [...used.values()].includes(free[0] + 8)) free.shift();
      if (free.length) { slot = free.shift(); palette[slot] = hex; paletteChanged = true; i = slot; } else {
        const [r, g, b] = rgb(c); let best = 0; let bd = Infinity;
        palette.forEach((p, k) => { const [r2, g2, b2] = rgb(`#${p}`); const dd = (r - r2) ** 2 + (g - g2) ** 2 + (b - b2) ** 2; if (dd < bd) { bd = dd; best = k; } });
        i = best; lose('цвета сверх 56 (заменены ближайшими)');
      }
    }
    used.set(hex, i + 8);
    return i + 8;
  };
  const fonts = []; const fontKey = new Map();
  const fontOf = (s) => {
    const f = { name: s.fontFamily || 'Calibri', size: s.fontSize || 11, bold: !!s.bold, italic: !!s.italic, underline: !!s.underline, strike: !!s.strike, color: s.color || null };
    const k = JSON.stringify(f);
    if (!fontKey.has(k)) { fonts.push(f); fontKey.set(k, fonts.length - 1); }
    const i = fontKey.get(k);
    return i >= 4 ? i + 1 : i; // шрифта с номером 4 в BIFF нет
  };
  const base = fontOf({}); for (let i = 1; i < 4; i++) { fonts.push(fonts[base]); } // шрифты 0–3 — по умолчанию
  const builtinId = new Map(Object.entries(BUILTIN_FMT).map(([id, code]) => [code, +id]));
  const formats = []; const fmtId = new Map();
  const fmtOf = (code) => {
    if (!code || code === 'General') return 0;
    if (builtinId.has(code)) return builtinId.get(code);
    if (!fmtId.has(code)) { fmtId.set(code, 164 + formats.length); formats.push(code); }
    return fmtId.get(code);
  };
  const xfs = []; const xfKey = new Map();
  const xfOf = (s) => {
    if (!s) return 15;
    const k = JSON.stringify(s);
    if (xfKey.has(k)) return xfKey.get(k);
    const b = new Buf();
    b.u16(fontOf(s)).u16(fmtOf(s.numFmt));
    b.u16((s.unlocked ? 0 : 1) | (s.formulaHidden ? 2 : 0)); // ячейка, родитель — стиль 0
    const alc = { left: 1, center: 2, right: 3, justify: 5 }[s.align] || 0;
    const alcV = s.valign === 'top' ? 0 : s.valign === 'center' ? 1 : 2;
    b.u8(alc | (s.wrap ? 0x08 : 0) | (alcV << 4)).u8(0).u8(0).u8(0xFC);
    const bd = s.border || {};
    const dg = (side) => (bd[side] ? BORDER_DG[bd[side].style] || 1 : 0);
    const bc = (side) => (bd[side] ? icv(bd[side].color) ?? 64 : 0);
    b.u32((dg('left') | (dg('right') << 4) | (dg('top') << 8) | (dg('bottom') << 12) | (bc('left') << 16) | (bc('right') << 23)) >>> 0);
    const fill = s.fill ? icv(s.fill) : null;
    b.u32((bc('top') | (bc('bottom') << 7) | ((fill != null ? 1 : 0) << 26)) >>> 0);
    b.u16((fill ?? 64) | (65 << 7));
    xfs.push(b.out());
    xfKey.set(k, 15 + xfs.length);
    return 15 + xfs.length;
  };

  // --- листы ---
  const sst = []; const sstIndex = new Map(); let sstTotal = 0;
  const strIx = (s) => { sstTotal++; if (!sstIndex.has(s)) { sstIndex.set(s, sst.length); sst.push(s); } return sstIndex.get(s); };
  const colLetters = (c) => { let n = c + 1; let r = ''; while (n) { const m = (n - 1) % 26; r = String.fromCharCode(65 + m) + r; n = Math.floor((n - 1) / 26); } return r; };
  const sheetStreams = sheets.map((sheet, si) => {
    const st = new Stream();
    st.rec(0x0809, new Buf().u16(0x0600).u16(0x0010).u16(0x0DBB).u16(0x07CC).u32(0).u32(0x06)); // BOF листа
    st.rec(0x000D, new Buf().u16(wb.calcMode === 'manual' ? 0 : 1)); // CALCMODE
    st.rec(0x000C, new Buf().u16(100)); // CALCCOUNT
    st.rec(0x000F, new Buf().u16(1)); // REFMODE: A1
    st.rec(0x0011, new Buf().u16(0)); // ITERATION
    st.rec(0x0010, new Buf().f64(0.001)); // DELTA
    st.rec(0x005F, new Buf().u16(1)); // SAVERECALC
    st.rec(0x002A, new Buf().u16(0)); // PRINTHEADERS
    const pg = sheet.page || {};
    st.rec(0x002B, new Buf().u16(pg.gridLines ? 1 : 0)); // PRINTGRIDLINES
    st.rec(0x0082, new Buf().u16(1)); // GRIDSET
    st.rec(0x0225, new Buf().u16(0).u16(300)); // DEFAULTROWHEIGHT: 15 pt
    const fit = pg.fitWidth != null || pg.fitHeight != null;
    st.rec(0x0081, new Buf().u16(0x04C1 | (fit ? 0x0100 : 0))); // WSBOOL
    st.rec(0x0014, new Buf().bytes(xlString(''))); st.rec(0x0015, new Buf().bytes(xlString(''))); // HEADER/FOOTER
    st.rec(0x0083, new Buf().u16(pg.centerH ? 1 : 0)); st.rec(0x0084, new Buf().u16(0)); // HCENTER/VCENTER
    const m = pg.margins || {};
    st.rec(0x0026, new Buf().f64(m.left ?? 0.7)); st.rec(0x0027, new Buf().f64(m.right ?? 0.7));
    st.rec(0x0028, new Buf().f64(m.top ?? 0.75)); st.rec(0x0029, new Buf().f64(m.bottom ?? 0.75));
    st.rec(0x00A1, new Buf().u16(pg.paper || 9).u16(pg.scale || 100).u16(1).u16(pg.fitWidth ?? 1).u16(pg.fitHeight ?? 1)
      .u16((pg.orientation === 'landscape' ? 0 : 0x02) | (pg.paper || pg.orientation || pg.scale ? 0 : 0x04)).u16(600).u16(600).f64(m.header ?? 0.3).f64(m.footer ?? 0.3).u16(1)); // SETUP
    if (sheet.protection) {
      st.rec(0x0012, new Buf().u16(1)); // PROTECT
      const pw = sheet.protection.attrs && sheet.protection.attrs.password;
      if (pw && /^[0-9A-F]{1,4}$/i.test(pw)) st.rec(0x0013, new Buf().u16(parseInt(pw, 16)));
      else if (sheet.protection.attrs && sheet.protection.attrs.hashValue) lose('пароль защиты листа (в .xls — только старый формат)');
    }
    st.rec(0x0055, new Buf().u16(8)); // DEFCOLWIDTH
    // столбцы
    const colKeys = new Set([...Object.keys(sheet.colW || {}), ...Object.keys(sheet.colInfo || {})].map(Number).filter((c) => c < MAXC));
    for (const c of [...colKeys].sort((a, b) => a - b)) {
      const px = sheet.colW && sheet.colW[c] != null ? sheet.colW[c] : DEFAULT_COL_W;
      const li = (sheet.colInfo && sheet.colInfo[c]) || {};
      const fl = (li.hidden ? 1 : 0) | 0x02 | (li.bestFit ? 4 : 0) | ((li.level || 0) << 8) | (li.collapsed ? 0x1000 : 0);
      st.rec(0x007D, new Buf().u16(c).u16(c).u16(Math.round(pxToWidth(px) * 256)).u16(15).u16(fl).u16(0));
    }
    // ячейки: свои + значения «разлива» (динамический массив — формулой массива на свой диапазон, как пишет Excel)
    const cells = new Map();
    let maxR = -1; let maxC = -1; let minR = Infinity; let minC = Infinity;
    const put = (r, c, x) => {
      if (r >= MAXR || c >= MAXC) { lose('ячейки за пределами 65 536 строк / 256 столбцов'); return; }
      cells.set(key(r, c), { r, c, ...x });
      if (r > maxR) maxR = r; if (c > maxC) maxC = c; if (r < minR) minR = r; if (c < minC) minC = c;
    };
    const arrays = new Map(); // ключ якоря → { r1, c1, r2, c2 }
    for (const [k, rg] of (sheet.spillAt || new Map())) arrays.set(k, rg);
    for (const [k, cell] of sheet.cells) {
      const [r, c] = k.split(',').map(Number);
      if (cell.cse) {
        const mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(cell.cse);
        if (mm) {
          const cn = (t) => [...t].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
          arrays.set(k, { r1: +mm[2] - 1, c1: cn(mm[1]), r2: +mm[4] - 1, c2: cn(mm[3]) });
        }
      }
      // формула с поэлементными операциями над массивами (СУММ(ДЛСТР(A1:A9)) и т. п.) — формулой массива, как пишет Excel
      if (cell.f != null && !cell.cse && (cell.ax || cell.arr) && !arrays.has(k)) arrays.set(k, { r1: r, c1: c, r2: r, c2: c });
      put(r, c, { cell });
    }
    if (sheet.spill) for (const [k, v] of sheet.spill) if (!cells.has(k)) { const [r, c] = k.split(',').map(Number); put(r, c, { cell: { v }, ghost: true }); }
    const arrayOf = new Map(); // ключ ячейки → ключ якоря её формулы массива
    for (const [ak, rg] of arrays) for (let r = rg.r1; r <= Math.min(rg.r2, MAXR - 1); r++) for (let c = rg.c1; c <= Math.min(rg.c2, MAXC - 1); c++) arrayOf.set(key(r, c), ak);
    const rowKeys = new Set([...cells.values()].map((x) => x.r));
    for (const r of Object.keys(sheet.rowH || {})) if (+r < MAXR) rowKeys.add(+r);
    for (const r of Object.keys(sheet.rowInfo || {})) if (+r < MAXR) rowKeys.add(+r);
    if (rowKeys.size) { maxR = Math.max(maxR, ...rowKeys); minR = Math.min(minR, ...rowKeys); }
    st.rec(0x0200, new Buf().u32(maxR < 0 ? 0 : minR).u32(maxR + 1).u16(maxC < 0 ? 0 : minC).u16(maxC + 1).u16(0)); // DIMENSIONS
    const byRow = new Map();
    for (const x of cells.values()) { if (!byRow.has(x.r)) byRow.set(x.r, []); byRow.get(x.r).push(x); }
    const rows = [...rowKeys].sort((a, b) => a - b);
    for (let i = 0; i < rows.length; i += 32) {
      const block = rows.slice(i, i + 32);
      for (const r of block) { // ROW
        const list = byRow.get(r) || [];
        const cMin = list.length ? Math.min(...list.map((x) => x.c)) : 0; const cMax = list.length ? Math.max(...list.map((x) => x.c)) + 1 : 0;
        const hpx = sheet.rowH && sheet.rowH[r];
        const li = (sheet.rowInfo && sheet.rowInfo[r]) || {};
        const custom = hpx != null && hpx !== DEFAULT_ROW_H;
        const fl = (li.level || 0) | (li.collapsed ? 0x10 : 0) | (li.hidden ? 0x20 : 0) | (custom ? 0x40 : 0) | 0x100;
        st.rec(0x0208, new Buf().u16(r).u16(cMin).u16(cMax).u16(custom ? Math.round(hpx * 0.75 * 20) : 300).u16(0).u16(0).u16(fl).u16(15));
      }
      for (const r of block) {
        for (const x of (byRow.get(r) || []).sort((a, b) => a.c - b.c)) {
          const { cell, c } = x;
          const xf = xfOf(cell.s);
          let ak = arrayOf.get(key(r, c));
          if (cell.f != null || ak) {
            let rgce; let rgcb = new Uint8Array(0);
            if (ak) {
              const [ar, ac] = ak.split(',').map(Number);
              rgce = Uint8Array.from([0x01, ar & 0xFF, ar >> 8, ac & 0xFF, ac >> 8]); // PtgExp → формула массива якоря
            } else {
              const res = tryCompile(cell.f, { own: si });
              if (!res) { lose('формулы, которых нет в Excel 97–2003 (записаны значением)'); writeValue(st, r, c, xf, cell.v); continue; }
              ({ rgce, rgcb } = res);
              if (res.needDyn || (res.needArray && !cell.legacy)) { // «динамическая» формула — формулой массива на свою ячейку, как пишет Excel
                ak = key(r, c);
                arrays.set(ak, { r1: r, c1: c, r2: r, c2: c });
                rgce = Uint8Array.from([0x01, r & 0xFF, r >> 8, c & 0xFF, c >> 8]); rgcb = new Uint8Array(0);
              }
            }
            const v = x.ghost ? cell.v : cell.v;
            const val = new Buf();
            let str = null;
            if (typeof v === 'number') val.f64(v);
            else if (typeof v === 'string') { if (v === '') val.bytes([3, 0, 0, 0, 0, 0, 0xFF, 0xFF]); else { val.bytes([0, 0, 0, 0, 0, 0, 0xFF, 0xFF]); str = v; } }
            else if (typeof v === 'boolean') val.bytes([1, 0, v ? 1 : 0, 0, 0, 0, 0xFF, 0xFF]);
            else if (v && v.error) val.bytes([2, 0, ERR_CODE[v.error] ?? 0x2A, 0, 0, 0, 0xFF, 0xFF]);
            else val.f64(0);
            st.rec(0x0006, new Buf().u16(r).u16(c).u16(xf).bytes(val.a).u16(0x0002).u32(0).u16(rgce.length).bytes(rgce).bytes(rgcb));
            if (ak === key(r, c)) { // запись ARRAY — сразу после формулы якоря
              const rg = arrays.get(ak);
              const res = tryCompile(cell.f, { own: si, mode: 'A' });
              if (res) st.rec(0x0221, new Buf().u16(rg.r1).u16(Math.min(rg.r2, MAXR - 1)).u8(rg.c1).u8(Math.min(rg.c2, MAXC - 1)).u16(0x0001).u32(0).u16(res.rgce.length).bytes(res.rgce).bytes(res.rgcb));
            }
            if (str != null) st.rec(0x0207, new Buf().bytes(xlString(str)));
          } else writeValue(st, r, c, xf, cell.v);
        }
      }
    }
    // окно: выбранный лист, закрепление
    const active = si === (wb.active || 0);
    const fr = sheet.freeze && (sheet.freeze.rows || sheet.freeze.cols) ? sheet.freeze : null;
    st.rec(0x023E, new Buf().u16(0x02 | 0x04 | 0x10 | 0x20 | 0x80 | (fr ? 0x108 : 0) | (active ? 0x600 : 0)).u16(0).u16(0).u16(64).u16(0).u16(0).u16(0).u32(0));
    if (fr) st.rec(0x0041, new Buf().u16(fr.cols || 0).u16(fr.rows || 0).u16(fr.rows || 0).u16(fr.cols || 0).u8(fr.rows && fr.cols ? 0 : fr.rows ? 2 : 1).u8(0));
    // объединения — не больше 1026 в записи
    const merges = (sheet.merges || []).filter((g) => g.r1 < MAXR && g.c1 < MAXC);
    for (let i = 0; i < merges.length; i += 1026) {
      const part = merges.slice(i, i + 1026);
      const b = new Buf().u16(part.length);
      for (const g of part) b.u16(g.r1).u16(Math.min(g.r2, MAXR - 1)).u16(g.c1).u16(Math.min(g.c2, MAXC - 1));
      st.rec(0x00E5, b);
    }
    if (sheet.autoFilter) lose('автофильтр');
    if (sheet.cf && sheet.cf.length) lose('условное форматирование');
    if (sheet.dv && sheet.dv.length) lose('проверка данных');
    st.rec(0x000A, new Uint8Array(0)); // EOF
    return st;
  });
  function writeValue(st, r, c, xf, v) {
    if (typeof v === 'number') st.rec(0x0203, new Buf().u16(r).u16(c).u16(xf).f64(v));
    else if (typeof v === 'string') st.rec(0x00FD, new Buf().u16(r).u16(c).u16(xf).u32(strIx(v)));
    else if (typeof v === 'boolean') st.rec(0x0205, new Buf().u16(r).u16(c).u16(xf).u8(v ? 1 : 0).u8(0));
    else if (v && v.error) st.rec(0x0205, new Buf().u16(r).u16(c).u16(xf).u8(ERR_CODE[v.error] ?? 0x2A).u8(1));
    else st.rec(0x0201, new Buf().u16(r).u16(c).u16(xf));
  }
  // формулы имён (после ячеек: функции _xlfn. из ячеек уже записаны в nameRecs)
  const nameBodies = nameRecs.map((n) => {
    if (n.fn) return { rgce: Uint8Array.from([0x1C, 0x17]), rgcb: new Uint8Array(0) }; // заглушка #ССЫЛКА!, как у Excel
    const res = tryCompile(n.text, { own: n.itab ? n.itab - 1 : 0, three: true, mode: 'R' });
    if (!res) lose('имена, которые не записались');
    return res || { rgce: Uint8Array.from([0x1C, 0x17]), rgcb: new Uint8Array(0) };
  });

  // --- глобальная часть ---
  const g = new Stream();
  g.rec(0x0809, new Buf().u16(0x0600).u16(0x0005).u16(0x0DBB).u16(0x07CC).u32(0).u32(0x06)); // BOF книги
  g.rec(0x00E1, new Buf().u16(1200)); g.rec(0x00C1, new Buf().u16(0)); g.rec(0x00E2, new Uint8Array(0)); // INTERFACEHDR, MMS, INTERFACEEND
  g.rec(0x005C, new Buf().bytes(xlString('Перо')).bytes(new Array(Math.max(0, 112 - 3 - 8)).fill(0x20)).out().subarray(0, 112)); // WRITEACCESS
  g.rec(0x0042, new Buf().u16(1200)); // CODEPAGE: UTF-16
  g.rec(0x0161, new Buf().u16(0)); // DSF
  g.rec(0x013D, new Buf().bytes(sheets.flatMap((_, i) => [(i + 1) & 0xFF, (i + 1) >> 8]))); // TABID
  g.rec(0x009C, new Buf().u16(14)); // FNGROUPCOUNT
  g.rec(0x0019, new Buf().u16(0)); g.rec(0x0012, new Buf().u16(0)); g.rec(0x0013, new Buf().u16(0)); // WINDOWPROTECT, PROTECT, PASSWORD
  g.rec(0x01AF, new Buf().u16(0)); g.rec(0x01BC, new Buf().u16(0)); // PROT4REV, PROT4REVPASS
  g.rec(0x0040, new Buf().u16(0)); g.rec(0x008D, new Buf().u16(0)); // BACKUP, HIDEOBJ
  g.rec(0x003D, new Buf().u16(0x0168).u16(0x0168).u16(0x4000).u16(0x2800).u16(0x0038).u16(wb.active || 0).u16(0).u16(1).u16(0x0258)); // WINDOW1
  g.rec(0x0022, new Buf().u16(0)); g.rec(0x000E, new Buf().u16(1)); g.rec(0x01B7, new Buf().u16(0)); g.rec(0x00DA, new Buf().u16(0)); // DATEMODE, PRECISION, REFRESHALL, BOOKBOOL
  // шрифты (палитра для цветов текста уже собрана при XF; цвета шрифтов — здесь)
  const fontRecs = fonts.map((f) => new Buf().u16(Math.round(f.size * 20)).u16((f.italic ? 2 : 0) | (f.strike ? 8 : 0)).u16(f.color ? icv(f.color) : 0x7FFF)
    .u16(f.bold ? 700 : 400).u16(0).u8(f.underline ? 1 : 0).u8(0).u8(1).u8(0).bytes(shortString(f.name)));
  for (const fr of fontRecs) g.rec(0x0031, fr);
  for (const [k, code] of formats.entries()) g.rec(0x041E, new Buf().u16(164 + k).bytes(xlString(code)));
  // XF: 0–14 — стили, 15 — ячейка по умолчанию, дальше — стили ячеек
  const styleXf = (fnt) => new Buf().u16(fnt).u16(0).u16(0xFFF5).u8(0x20).u8(0).u8(0).u8(fnt ? 0xF4 : 0).u32(0).u32(0).u16(64 | (65 << 7));
  for (let i = 0; i < 15; i++) g.rec(0x00E0, styleXf(i === 0 ? 0 : i < 3 ? 1 : i < 5 ? 2 : 0));
  g.rec(0x00E0, new Buf().u16(0).u16(0).u16(0x0001).u8(0x20).u8(0).u8(0).u8(0).u32(0).u32(0).u16(64 | (65 << 7)));
  for (const x of xfs) g.rec(0x00E0, x);
  g.rec(0x0293, new Buf().u16(0x8000).u8(0).u8(0xFF)); // STYLE «Обычный»
  if (paletteChanged) { const b = new Buf().u16(56); for (const p of palette) b.bytes([...rgb(`#${p}`), 0]); g.rec(0x0092, b); }
  g.rec(0x0160, new Buf().u16(0)); // USESELFS
  const boundPos = [];
  for (const s of sheets) { boundPos.push(g.len + 4); g.rec(0x0085, new Buf().u32(0).u8(0).u8(0).bytes(shortString(s.name))); }
  g.rec(0x008C, new Buf().u16(7).u16(7)); // COUNTRY: Россия
  if (xtis.length || nameRecs.length) {
    g.rec(0x01AE, new Buf().u16(sheets.length).u16(0x0401)); // SUPBOOK: эта книга
    if (addinUsed) {
      g.rec(0x01AE, new Buf().u16(1).u16(0x3A01)); // SUPBOOK: надстройка (функции пакета анализа)
      for (const n of addinNames) g.rec(0x0023, new Buf().u16(0).u32(0).u8(n.length).u8(wide(n) ? 1 : 0).bytes(charBytes(n, wide(n))).u16(2).bytes([0x1C, 0x17]));
    }
    const b = new Buf().u16(xtis.length);
    for (const [book, a, c] of xtis) b.u16(book).u16(a & 0xFFFF).u16(c & 0xFFFF);
    g.rec(0x0017, b); // EXTERNSHEET
  }
  nameRecs.forEach((n, i) => {
    const body = nameBodies[i];
    const nm = n.builtin ? String.fromCharCode(n.builtin) : n.name;
    const h = wide(nm);
    const fl = (n.hidden ? 0x01 : 0) | (n.fn ? 0x02 | 0x08 : 0) | (n.builtin ? 0x20 : 0);
    g.rec(0x0018, new Buf().u16(fl).u8(0).u8(nm.length).u16(body.rgce.length).u16(0).u16(n.itab || 0).u8(0).u8(0).u8(0).u8(0)
      .u8(h ? 1 : 0).bytes(charBytes(nm, h)).bytes(body.rgce).bytes(body.rgcb));
  });
  // SST — строки разрываются на границах записей (в продолжении — свой байт флагов); EXTSST — каждая 8-я строка
  const sstStart = g.len;
  const extsst = [];
  {
    const recs = []; let cur = new Buf().u32(sstTotal).u32(sst.length);
    const MAX = 8224;
    const flush = () => { recs.push(cur.out()); cur = new Buf(); };
    sst.forEach((s, i) => {
      const h = wide(s); const per = h ? 2 : 1;
      if (cur.length + 3 + per > MAX) flush();
      if (i % 8 === 0) extsst.push({ rec: recs.length, off: cur.length });
      cur.u16(s.length).u8(h ? 1 : 0);
      let p = 0;
      while (p < s.length) {
        const n = Math.min(s.length - p, Math.floor((MAX - cur.length) / per));
        if (n <= 0) { flush(); cur.u8(h ? 1 : 0); continue; }
        cur.bytes(charBytes(s.slice(p, p + n), h)); p += n;
      }
    });
    recs.push(cur.out());
    let pos = sstStart; const recPos = [];
    recs.forEach((d, k) => { recPos.push(pos); g.rec(k === 0 ? 0x00FC : 0x003C, d); pos += 4 + d.length; });
    const b = new Buf().u16(8);
    for (const e of extsst) b.u32(recPos[e.rec] + 4 + e.off).u16(4 + e.off).u16(0);
    g.rec(0x00FF, b);
  }
  g.rec(0x000A, new Uint8Array(0)); // EOF
  // положение каждого листа в потоке — в BOUNDSHEET8
  const globals = g.out();
  let pos = globals.length;
  for (let i = 0; i < sheets.length; i++) {
    const p = boundPos[i];
    globals[p] = pos & 0xFF; globals[p + 1] = (pos >> 8) & 0xFF; globals[p + 2] = (pos >> 16) & 0xFF; globals[p + 3] = (pos >>> 24) & 0xFF;
    pos += sheetStreams[i].len;
  }
  const all = new Uint8Array(pos);
  all.set(globals, 0);
  let o = globals.length;
  for (const s of sheetStreams) { all.set(s.out(), o); o += s.len; }
  return { bytes: writeCfb([{ path: 'Workbook', data: all }]), lost };
}
