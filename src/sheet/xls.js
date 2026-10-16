// Книга Excel 97–2003 (.xls, BIFF8) → модель книги — своим кодом по спецификации [MS-XLS] (этап 19).
// Файл — составной (cfb.js), книга — поток «Workbook»: записи [тип u16][длина u16][данные]. Глобальная часть (шрифты,
// форматы, XF, общие строки SST, листы BOUNDSHEET8, имена, внешние ссылки), затем часть каждого листа (ячейки, формулы,
// строки, столбцы, объединения, окно, параметры страницы). Формулы — байт-код Ptg в обратной польской записи →
// текст формулы в канонической записи модели (как из .xlsx).
import { readCfb } from '../formats/cfb.js';
import { colName } from './address.js';
import { BUILTIN_FMT } from './format.js';
import { canonical } from './formula.js';
import { FTAB, UDF } from './xlsFtab.js';
import { newSheet, recalcOpen, RECOMPUTE, DEFAULT_COL_W } from './model.js';
import { parsePrintArea, parsePrintTitles } from './pageSetup.js';
import JSZip from 'jszip';
import { INDEXED, applyTint, widthToPx } from './xlsx.js';

export const isXls = (bytes) => { try { const c = readCfb(bytes); return !!(c.stream('Workbook') || c.stream('Book')); } catch { return false; } };

// ---------- Записи ----------
const R = {
  BOF: 0x0809, EOF: 0x000A, CONTINUE: 0x003C, CODEPAGE: 0x0042, DATEMODE: 0x0022, FONT: 0x0031, FORMAT: 0x041E, XF: 0x00E0,
  PALETTE: 0x0092, BOUNDSHEET: 0x0085, SST: 0x00FC, NAME: 0x0018, SUPBOOK: 0x01AE, EXTERNNAME: 0x0023, EXTERNSHEET: 0x0017,
  WINDOW1: 0x003D, CALCMODE: 0x000D,
  DIMENSIONS: 0x0200, ROW: 0x0208, COLINFO: 0x007D, BLANK: 0x0201, MULBLANK: 0x00BE, NUMBER: 0x0203, RK: 0x027E, MULRK: 0x00BD,
  LABELSST: 0x00FD, LABEL: 0x0204, RSTRING: 0x00D6, BOOLERR: 0x0205, FORMULA: 0x0006, STRING: 0x0207, SHRFMLA: 0x04BC,
  ARRAY: 0x0221, MERGECELLS: 0x00E5, WINDOW2: 0x023E, PANE: 0x0041, SETUP: 0x00A1, LEFTMARGIN: 0x0026, RIGHTMARGIN: 0x0027,
  TOPMARGIN: 0x0028, BOTTOMMARGIN: 0x0029, PRINTGRIDLINES: 0x002B, HCENTER: 0x0083, WSBOOL: 0x0081, PROTECT: 0x0012,
  PASSWORD: 0x0013, HORIZONTALPAGEBREAKS: 0x001B, VERTICALPAGEBREAKS: 0x001A, MSODRAWING: 0x00EC, OBJ: 0x005D, NOTE: 0x001C,
  HLINK: 0x01B8, CF: 0x01B1, DV: 0x01BE, AUTOFILTER: 0x009E, CHART_BOF: 0x0809, XFEXT: 0x087D, THEME: 0x0896, CONTINUEFRT12: 0x087F,
};
// Записи, которые Перо не разбирает, но переносит при сохранении в .xls как есть (зона — место в потоке листа)
const KEEP_PRE = new Set([0x009B, 0x009D, 0x009E, 0x087E, 0x0090]); // автофильтр, сортировка — до DIMENSIONS
const KEEP_OBJ = new Set([0x00EC, 0x005D, 0x01B6, 0x001C, 0x00ED]); // рисунки, объекты, надписи, примечания — до WINDOW2
const KEEP_PROT = new Set([0x00DD, 0x0063]); // защита сценариев и объектов — после PROTECT
const KEEP_TAIL = new Set([0x01B0, 0x01B1, 0x0879, 0x087A, 0x087B, 0x01B8, 0x0800, 0x01B2, 0x01BE, 0x0862, 0x0867, 0x0868,
  0x0871, 0x0872, 0x0877]); // условное форматирование, ссылки, проверка, таблицы, защита (FEAT) — в конце листа
const KEEP_LABEL = { 0x009E: 'автофильтр', 0x001C: 'примечания к ячейкам', 0x00EC: 'рисунки и фигуры', 0x01B1: 'условное форматирование',
  0x087A: 'условное форматирование', 0x01B8: 'гиперссылки', 0x01BE: 'проверка данных', 0x0872: 'таблицы', 0x004D: 'параметры принтера',
  0x0014: 'колонтитулы', 0x0015: 'колонтитулы' };

// цвета темы Office по умолчанию в порядке индексов Excel: lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink
const THEME_DEF = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

function records(bytes) {
  const out = [];
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let p = 0; p + 4 <= bytes.length;) {
    const type = dv.getUint16(p, true); const len = dv.getUint16(p + 2, true);
    out.push({ type, pos: p, data: bytes.subarray(p + 4, p + 4 + len) });
    p += 4 + len;
  }
  return out;
}

// Чтение подряд по нескольким частям (запись + CONTINUE): строки, разорванные границей записи, в новой части
// начинаются с байта флагов (однобайтные или двухбайтные символы)
class Seg {
  constructor(parts) { this.parts = parts; this.i = 0; this.p = 0; }
  get cur() { return this.parts[this.i]; }
  avail() { while (this.cur && this.p >= this.cur.length) { this.i++; this.p = 0; } return this.cur ? this.cur.length - this.p : 0; }
  u8() { this.avail(); return this.cur[this.p++]; }
  u16() { return this.u8() | (this.u8() << 8); }
  u32() { return (this.u16() | (this.u16() << 16)) >>> 0; }
  skip(n) { while (n > 0) { const a = this.avail(); if (!a) return; const k = Math.min(a, n); this.p += k; n -= k; } }
  chars(n, high) {
    let s = '';
    while (n > 0) {
      if (!this.avail()) break;
      const per = high ? 2 : 1;
      const k = Math.min(n, Math.floor((this.cur.length - this.p) / per));
      for (let j = 0; j < k; j++) { s += String.fromCharCode(high ? this.cur[this.p] | (this.cur[this.p + 1] << 8) : this.cur[this.p]); this.p += per; }
      n -= k;
      if (n > 0) { this.i++; this.p = 0; if (!this.cur) break; high = (this.cur[this.p++] & 1) === 1; } // продолжение — свой флаг
    }
    return s;
  }
  // XLUnicodeRichExtendedString (общие строки)
  richString() {
    const cch = this.u16(); const fl = this.u8();
    const cRun = fl & 0x08 ? this.u16() : 0;
    const cbExt = fl & 0x04 ? this.u32() : 0;
    const s = this.chars(cch, (fl & 1) === 1);
    this.skip(cRun * 4 + cbExt);
    return s;
  }
}

// строки внутри одной записи
const u16 = (d, o) => d[o] | (d[o + 1] << 8);
const i16 = (d, o) => { const v = u16(d, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (d, o) => (u16(d, o) | (u16(d, o + 2) << 16)) >>> 0;
const f64 = (d, o) => new DataView(d.buffer, d.byteOffset + o, 8).getFloat64(0, true);
function charsAt(d, o, n, high) {
  let s = '';
  for (let j = 0; j < n; j++) s += String.fromCharCode(high ? u16(d, o + j * 2) : d[o + j]);
  return s;
}
// XLUnicodeString: cch u16, флаги u8, символы → [строка, длина в байтах]
function xlString(d, o) {
  const cch = u16(d, o); const fl = d[o + 2];
  let p = o + 3;
  const cRun = fl & 0x08 ? u16(d, p) : 0; if (fl & 0x08) p += 2;
  const cbExt = fl & 0x04 ? u32(d, p) : 0; if (fl & 0x04) p += 4;
  const high = (fl & 1) === 1;
  const s = charsAt(d, p, cch, high);
  p += cch * (high ? 2 : 1) + cRun * 4 + cbExt;
  return [s, p - o];
}
// ShortXLUnicodeString: cch u8, флаги u8
function shortString(d, o) {
  const cch = d[o]; const high = (d[o + 1] & 1) === 1;
  return [charsAt(d, o + 2, cch, high), 2 + cch * (high ? 2 : 1)];
}

// ---------- Формулы: Ptg → текст ----------
const ERRS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#GETTING_DATA' };
const BINOP = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>', 0x0F: ' ', 0x10: ',', 0x11: ':' };
const MAXR = 65535; const MAXC = 255;

export const quoteSheet = (name) => (/^[A-Za-z_Ѐ-ӿ][\w.Ѐ-ӿ]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R\d*C\d*|TRUE|FALSE)$/i.test(name)
  ? name : `'${String(name).replace(/'/g, "''")}'`);
const numText = (n) => { const s = String(n); return /e/i.test(s) ? s.toUpperCase() : s; };
const strText = (s) => `"${s.replace(/"/g, '""')}"`;

// ссылка на ячейку: строка/столбец и признаки относительности (в BIFF8 — старшие биты поля столбца)
function cellText(row, colField, base) {
  const rowRel = (colField & 0x8000) !== 0; const colRel = (colField & 0x4000) !== 0;
  let r = row; let c = colField & 0x3FFF;
  if (base) { // PtgRefN/PtgAreaN: относительные — смещения от ячейки (по модулю размеров листа BIFF8)
    if (rowRel) r = (base.r + (row & 0x8000 ? row - 0x10000 : row)) & 0xFFFF;
    if (colRel) c = (base.c + ((c & 0xFF) >= 0x80 ? (c & 0xFF) - 0x100 : (c & 0xFF))) & 0xFF;
  }
  return { r, c, rowRel, colRel };
}
const refText = (x) => `${x.colRel ? '' : '$'}${colName(x.c)}${x.rowRel ? '' : '$'}${x.r + 1}`;
function areaText(a, b) {
  if (a.r === 0 && b.r === MAXR) return `${a.colRel ? '' : '$'}${colName(a.c)}:${b.colRel ? '' : '$'}${colName(b.c)}`; // целые столбцы
  if (a.c === 0 && b.c === MAXC) return `${a.rowRel ? '' : '$'}${a.r + 1}:${b.rowRel ? '' : '$'}${b.r + 1}`; // целые строки
  return `${refText(a)}:${refText(b)}`;
}

// rgce — байт-код, rgcb — дополнительные данные (массивы-константы); ctx: { base (для RefN), sheetRef(ixti), name(i),
// externName(ixti, i) }. Возвращает текст формулы (без «=») или null, если встретилось неизвестное.
export function decompile(rgce, rgcb, ctx) {
  const st = [];
  let p = 0; let q = 0; // позиция в rgce и в rgcb
  const pop = () => st.pop() ?? '';
  const arg = (n) => st.splice(st.length - n, n);
  while (p < rgce.length) {
    const t = rgce[p];
    const base = t & 0x1F | (t >= 0x20 ? 0x20 : 0); // класс операнда не важен для текста
    if (t >= 0x03 && t <= 0x11) { const b = pop(); const a = pop(); st.push(t === 0x10 ? `${a},${b}` : `${a}${BINOP[t]}${b}`); p++; continue; }
    switch (t) {
      case 0x01: case 0x02: return null; // PtgExp/PtgTbl — разворачивает вызывающий
      case 0x12: st.push(`+${pop()}`); p++; continue;
      case 0x13: st.push(`-${pop()}`); p++; continue;
      case 0x14: st.push(`${pop()}%`); p++; continue;
      case 0x15: st.push(`(${pop()})`); p++; continue;
      case 0x16: st.push(''); p++; continue; // пропущенный аргумент
      case 0x17: { const [s, n] = shortString(rgce, p + 1); st.push(strText(s)); p += 1 + n; continue; }
      case 0x18: return null; // PtgExtend (ссылки на структурированные таблицы и т. п.) — не поддерживается
      case 0x19: {
        const g = rgce[p + 1];
        if (g & 0x04) { const w = u16(rgce, p + 2); p += 4 + 2 * (w + 1); continue; } // tAttrChoose: таблица переходов
        if (g & 0x10) st.push(`SUM(${pop()})`); // tAttrSum — СУММ одного аргумента
        p += 4; continue; // tAttrSemi/If/Goto/Space — для текста не нужны
      }
      case 0x1C: st.push(ERRS[rgce[p + 1]] || '#N/A'); p += 2; continue;
      case 0x1D: st.push(rgce[p + 1] ? 'TRUE' : 'FALSE'); p += 2; continue;
      case 0x1E: st.push(String(u16(rgce, p + 1))); p += 3; continue;
      case 0x1F: st.push(numText(f64(rgce, p + 1))); p += 9; continue;
      default: break;
    }
    switch (base) {
      case 0x20: { // PtgArray — значения в rgcb
        const cols = rgcb[q] + 1; const rows = u16(rgcb, q + 1) + 1; q += 3;
        const rowsOut = [];
        for (let r = 0; r < rows; r++) {
          const row = [];
          for (let c = 0; c < cols; c++) {
            const k = rgcb[q];
            if (k === 0x01) { row.push(numText(f64(rgcb, q + 1))); q += 9; }
            else if (k === 0x02) { const [s, n] = xlString(rgcb, q + 1); row.push(strText(s)); q += 1 + n; }
            else if (k === 0x04) { row.push(rgcb[q + 1] ? 'TRUE' : 'FALSE'); q += 9; }
            else if (k === 0x10) { row.push(ERRS[rgcb[q + 1]] || '#N/A'); q += 9; }
            else { row.push(''); q += 9; }
          }
          rowsOut.push(row.join(','));
        }
        st.push(`{${rowsOut.join(';')}}`); p += 8; continue;
      }
      case 0x21: { // PtgFunc — постоянное число аргументов
        const f = FTAB.get(u16(rgce, p + 1));
        if (!f) return null;
        st.push(`${f.name}(${arg(f.argc).join(',')})`); p += 3; continue;
      }
      case 0x22: { // PtgFuncVar
        const n = rgce[p + 1] & 0x7F; const tab = u16(rgce, p + 2);
        if (tab & 0x8000) return null; // команда макроса
        const args = arg(n);
        if (tab === UDF) { const name = args.shift().replace(/^_xlfn\.(_xlws\.)?/, ''); st.push(`${name}(${args.join(',')})`); }
        else { const f = FTAB.get(tab); if (!f) return null; st.push(`${f.name}(${args.join(',')})`); }
        p += 4; continue;
      }
      case 0x23: st.push(ctx.name(u32(rgce, p + 1))); p += 5; continue; // PtgName
      case 0x24: st.push(refText(cellText(u16(rgce, p + 1), u16(rgce, p + 3), null))); p += 5; continue;
      case 0x25: st.push(areaText(cellText(u16(rgce, p + 1), u16(rgce, p + 5), null), cellText(u16(rgce, p + 3), u16(rgce, p + 7), null))); p += 9; continue;
      case 0x26: { const n = u16(rgcb, q); q += 2 + n * 8; p += 7; continue; } // PtgMemArea: выражение следует за ним
      case 0x27: case 0x28: p += 7; continue;
      case 0x29: p += 3; continue;
      case 0x2A: st.push('#REF!'); p += 5; continue;
      case 0x2B: st.push('#REF!'); p += 9; continue;
      case 0x2C: st.push(refText(cellText(u16(rgce, p + 1), u16(rgce, p + 3), ctx.base))); p += 5; continue;
      case 0x2D: st.push(areaText(cellText(u16(rgce, p + 1), u16(rgce, p + 5), ctx.base), cellText(u16(rgce, p + 3), u16(rgce, p + 7), ctx.base))); p += 9; continue;
      case 0x39: st.push(ctx.externName(u16(rgce, p + 1), u16(rgce, p + 3))); p += 7; continue; // PtgNameX
      case 0x3A: { const sh = ctx.sheetRef(u16(rgce, p + 1)); st.push(sh == null ? '#REF!' : `${sh}!${refText(cellText(u16(rgce, p + 3), u16(rgce, p + 5), ctx.base3d))}`); p += 7; continue; }
      case 0x3B: {
        const sh = ctx.sheetRef(u16(rgce, p + 1));
        st.push(sh == null ? '#REF!' : `${sh}!${areaText(cellText(u16(rgce, p + 3), u16(rgce, p + 7), ctx.base3d), cellText(u16(rgce, p + 5), u16(rgce, p + 9), ctx.base3d))}`);
        p += 11; continue;
      }
      case 0x3C: st.push('#REF!'); p += 7; continue;
      case 0x3D: st.push('#REF!'); p += 11; continue;
      default: return null;
    }
  }
  return st.length === 1 ? st[0] : st.length ? st.join(',') : '';
}

// ---------- Стили ----------
const BORDER = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];

// ---------- Книга ----------
export async function importXls(data, { recalc: doRecalc = true } = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const cfb = readCfb(bytes);
  const stream = cfb.stream('Workbook') || cfb.stream('Book');
  if (!stream) throw new Error('Это не книга Excel: внутри нет потока Workbook');
  const recs = records(stream);
  const kept = {}; // предупреждения: при сохранении потеряется
  const keep = (k) => { kept[k] = (kept[k] || 0) + 1; };
  const saved = {}; // переносится при сохранении в .xls как есть
  const rawRec = (r) => stream.subarray(r.pos, r.pos + 4 + r.data.length);
  const drawingGroup = []; // MSODRAWINGGROUP (+ CONTINUE) — общая часть рисунков всех листов
  const fontRaw = []; const fmtRaw = []; // записи FONT и FORMAT как в файле: на их номера ссылаются диаграммы и надписи
  if (cfb.entries.some((e) => /^_VBA_PROJECT_CUR/i.test(e.path))) keep('макросы VBA');

  // глобальная часть
  const fonts = []; const formats = {}; const xfs = []; const palette = INDEXED.slice(); const xfExt = {}; let themeZip = null;
  const sheetsInfo = []; let sst = []; const names = []; const supbooks = []; const xti = [];
  let date1904 = false; let activeTab = 0; let manual = false;
  let i = 0;
  if (!recs.length || recs[0].type !== R.BOF || u16(recs[0].data, 0) !== 0x0600) throw new Error('Поддерживается формат Excel 97–2003 (BIFF8); этот файл старше');
  for (i = 1; i < recs.length; i++) {
    const { type, data: d } = recs[i];
    if (type === R.EOF) break;
    switch (type) {
      case R.DATEMODE: date1904 = u16(d, 0) === 1; break;
      case R.WINDOW1: activeTab = u16(d, 10); break;
      case R.CALCMODE: manual = i16(d, 0) === 0; break;
      case R.FONT: {
        const [name] = shortString(d, 14);
        fonts.push({ size: u16(d, 0) / 20, italic: !!(u16(d, 2) & 0x02), strike: !!(u16(d, 2) & 0x08), icv: u16(d, 4), bold: u16(d, 6) >= 700, sss: u16(d, 8), u: d[10], name });
        if (fonts.length === 4) fonts.push(null); // шрифта с номером 4 в BIFF нет
        fontRaw.push(rawRec(recs[i]));
        break;
      }
      case R.FORMAT: formats[u16(d, 0)] = xlString(d, 2)[0]; fmtRaw.push({ id: u16(d, 0), code: formats[u16(d, 0)], raw: rawRec(recs[i]) }); break;
      case R.XF: xfs.push(d.slice()); break;
      case R.XFEXT: { // точные цвета стиля (Excel 2007+): заливка, текст, границы — RGB или цвет темы с оттенком
        const ix = u16(d, 14); const n = u16(d, 18); let o = 20; const ext = {};
        for (let k = 0; k < n && o + 4 <= d.length; k++) {
          const t = u16(d, o); const cb = u16(d, o + 2);
          if ([4, 5, 7, 8, 9, 10, 13].includes(t)) ext[t] = { kind: u16(d, o + 4), tint: i16(d, o + 6) / 32767, val: d.subarray(o + 8, o + 12) };
          o += cb;
        }
        xfExt[ix] = ext;
        break;
      }
      case R.THEME: { // тема — ZIP-пакет с theme1.xml (может продолжаться записями CONTINUEFRT12)
        const parts = [d.subarray(16)];
        while (recs[i + 1] && recs[i + 1].type === R.CONTINUEFRT12) parts.push(recs[++i].data.subarray(12));
        const len = parts.reduce((s, x) => s + x.length, 0);
        if (len > 4) { const buf = new Uint8Array(len); let o = 0; for (const x of parts) { buf.set(x, o); o += x.length; } themeZip = buf; }
        break;
      }
      case R.PALETTE: { const n = u16(d, 0); for (let k = 0; k < n && 8 + k < 64; k++) palette[8 + k] = [d[2 + k * 4], d[3 + k * 4], d[4 + k * 4]].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase(); break; }
      case R.BOUNDSHEET: sheetsInfo.push({ pos: u32(d, 0), hidden: d[4] & 3, kind: d[5], name: shortString(d, 6)[0] }); break;
      case R.SST: {
        const parts = [d.subarray(8)];
        while (recs[i + 1] && recs[i + 1].type === R.CONTINUE) parts.push(recs[++i].data);
        const seg = new Seg(parts); const n = u32(d, 4);
        sst = new Array(n);
        for (let k = 0; k < n; k++) sst[k] = seg.richString();
        break;
      }
      case 0x00EB: drawingGroup.push(rawRec(recs[i])); while (recs[i + 1] && recs[i + 1].type === R.CONTINUE) drawingGroup.push(rawRec(recs[++i])); break;
      case R.SUPBOOK: {
        const ctab = u16(d, 0); const cch = u16(d, 2);
        supbooks.push({ self: cch === 0x0401, addin: cch === 0x3A01, ctab, names: [] });
        break;
      }
      case R.EXTERNNAME: {
        const sb = supbooks[supbooks.length - 1];
        if (!sb) break;
        // ExternName: флаги u16, затем (надстройка/функция) 4 байта, длина u8, имя без длины (флаги + символы)
        const cch = d[6]; const high = (d[7] & 1) === 1;
        sb.names.push(charsAt(d, 8, cch, high));
        break;
      }
      case R.EXTERNSHEET: { const n = u16(d, 0); for (let k = 0; k < n; k++) xti.push({ book: u16(d, 2 + k * 6), first: i16(d, 4 + k * 6), last: i16(d, 6 + k * 6) }); break; }
      case R.NAME: {
        const fl = u16(d, 0); const cch = d[3]; const cce = u16(d, 4); const itab = u16(d, 8);
        const high = (d[14] & 1) === 1;
        const raw = charsAt(d, 15, cch, high);
        const fstart = 15 + cch * (high ? 2 : 1);
        names.push({ builtin: !!(fl & 0x20), hidden: !!(fl & 0x01), fn: !!(fl & 0x02), name: fl & 0x20 ? `\u0001${raw.charCodeAt(0)}` : raw, itab, rgce: d.subarray(fstart, fstart + cce), rgcb: d.subarray(fstart + cce), data: d.slice() });
        break;
      }
      default: break;
    }
  }
  const sheetNames = sheetsInfo.map((s) => s.name);
  let theme = THEME_DEF;
  if (themeZip) {
    try {
      const z = await JSZip.loadAsync(themeZip);
      const f = Object.keys(z.files).find((p) => /theme\d*\.xml$/.test(p));
      const xml = f ? await z.file(f).async('string') : '';
      const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml);
      if (scheme) {
        const val = (n) => { const m = new RegExp(`<a:${n}>\\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr\\b[^>]*lastClr="([0-9A-Fa-f]{6})")`).exec(scheme[0]); return m ? (m[1] || m[2]).toUpperCase() : null; };
        theme = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map((n, k) => val(n) || THEME_DEF[k]);
      }
    } catch { /* тема не прочиталась — цвета Office */ }
  }
  // цвет из XFEXT: { kind 1 индекс, 2 RGB, 3 тема; tint } → '#rrggbb' | null (авто); undefined — не задан
  const extColor = (e) => {
    if (!e) return undefined;
    if (e.kind === 2) return `#${applyTint([e.val[0], e.val[1], e.val[2]].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase(), e.tint).toLowerCase()}`;
    if (e.kind === 3) { const t = theme[u32(e.val, 0)]; return t ? `#${applyTint(t, e.tint).toLowerCase()}` : undefined; }
    if (e.kind === 1) { const v = u32(e.val, 0); return v >= 64 ? null : palette[v] ? `#${applyTint(palette[v], e.tint).toLowerCase()}` : undefined; }
    if (e.kind === 0) return null;
    return undefined;
  };
  const color = (icv) => (icv >= 64 || icv == null ? null : palette[icv] ? `#${palette[icv].toLowerCase()}` : null);
  const nameText = (k) => { const n = names[k - 1]; return !n ? '#NAME?' : n.builtin ? `_xlnm.${BUILTIN_NAMES[n.name.slice(1)] || n.name}` : n.name; };
  const ctx = {
    name: nameText,
    sheetRef: (ix) => {
      const x = xti[ix];
      if (!x) return null;
      const sb = supbooks[x.book];
      if (!sb || !sb.self || x.first < 0) return null;
      const a = sheetNames[x.first]; const b = sheetNames[x.last];
      if (a == null) return null;
      return x.first === x.last ? quoteSheet(a) : `${quoteSheet(a)}:${quoteSheet(b)}`;
    },
    externName: (ix, k) => { const x = xti[ix]; const sb = x && supbooks[x.book]; const n = sb && sb.names[k - 1]; return n || '#NAME?'; },
  };
  // стили ячеек (XF): шрифт, заливка, границы, формат, выравнивание, защита — как у .xlsx
  const styles = xfs.map((d, ix) => {
    const s = {};
    const ext = xfExt[ix] || {};
    const ifnt = u16(d, 0); const f = fonts[ifnt];
    if (f) {
      if (f.bold) s.bold = true;
      if (f.italic) s.italic = true;
      if (f.u) s.underline = true;
      if (f.strike) s.strike = true;
      if (f.size !== 11) s.fontSize = f.size;
      if (f.name !== 'Calibri') s.fontFamily = f.name;
      // цвет текста: тема «Текст 1» без оттенка и чёрный — «автоматический» (как у .xlsx)
      const e = ext[13];
      const c = e && e.kind === 3 && u32(e.val, 0) === 1 && !e.tint ? null : extColor(e) !== undefined ? extColor(e) : color(f.icv);
      if (c && c !== '#000000') s.color = c;
    }
    const ifmt = u16(d, 2);
    const code = formats[ifmt] ?? BUILTIN_FMT[ifmt];
    if (code && code !== 'General') s.numFmt = code === 'mm-dd-yy' ? 'dd.mm.yyyy' : code;
    const prot = u16(d, 4);
    if (!(prot & 1)) s.unlocked = true;
    if (prot & 2) s.formulaHidden = true;
    const al = d[6];
    const h = al & 7;
    if (h === 1) s.align = 'left'; else if (h === 2 || h === 6) s.align = 'center'; else if (h === 3) s.align = 'right'; else if (h === 5) s.align = 'justify';
    if (al & 0x08) s.wrap = true;
    const v = (al >> 4) & 7;
    if (v === 0) s.valign = 'top'; else if (v === 1) s.valign = 'center';
    const b1 = u32(d, 10); const b2 = u32(d, 14);
    const border = {};
    const side = (name, dg, icv, et) => { if (BORDER[dg]) border[name] = { style: BORDER[dg], color: extColor(ext[et]) !== undefined ? extColor(ext[et]) : color(icv) }; };
    side('left', b1 & 0xF, (b1 >>> 16) & 0x7F, 9); side('right', (b1 >>> 4) & 0xF, (b1 >>> 23) & 0x7F, 10);
    side('top', (b1 >>> 8) & 0xF, b2 & 0x7F, 7); side('bottom', (b1 >>> 12) & 0xF, (b2 >>> 7) & 0x7F, 8);
    if (Object.keys(border).length) s.border = border;
    const fls = (b2 >>> 26) & 0x3F;
    if (fls === 1) { const ec = extColor(ext[4]); const c = ec !== undefined ? ec : color(u16(d, 18) & 0x7F); if (c) s.fill = c; }
    return Object.keys(s).length ? s : null;
  });

  // листы
  const posToIndex = new Map(recs.map((r, k) => [r.pos, k]));
  const sheets = [];
  // область печати и сквозные строки: встроенные имена или (так пишет Excel 2007+) обычные Print_Area/Print_Titles листа
  const printKind = (n) => (n.builtin ? n.name : /^print_area$/i.test(n.name) ? '\u00016' : /^print_titles$/i.test(n.name) ? '\u00017' : null);
  const printNames = names.filter((n) => n.itab && printKind(n));
  for (const [si, info] of sheetsInfo.entries()) {
    const sheet = newSheet(info.name);
    sheets.push(sheet);
    if (info.kind !== 0) { keep('листы диаграмм'); continue; }
    let k = posToIndex.get(info.pos);
    if (k == null) continue;
    const shared = []; // { r1, r2, c1, c2, rgce, rgcb } — общие формулы
    const arrays = [];
    let pendingString = null; // ячейка формулы, чей текстовый результат — в следующей записи STRING
    const page = {}; let fitToPage = false; let pass = null; let locked = false;
    const setCell = (r, c, cell) => { if (cell.v == null && cell.f == null && !cell.s) return; sheet.cells.set(`${r},${c}`, cell); };
    const styleOf = (ixfe) => styles[ixfe] || null;
    const withStyle = (cell, ixfe) => { const st = styleOf(ixfe); if (st) cell.s = st; return cell; };
    const numVal = (v, ixfe) => { const st = styleOf(ixfe); return date1904 && st && st.numFmt && /[dy]/i.test(st.numFmt) ? v + 1462 : v; };
    const formulaText = (rgce, rgcb, r, c) => {
      if (rgce[0] === 0x01 && rgce.length === 5) { // PtgExp — общая формула или формула массива
        const mr = u16(rgce, 1); const mc = u16(rgce, 3);
        const sh = shared.find((x) => x.r1 === mr && x.c1 === mc) || shared.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2);
        if (sh) return decompile(sh.rgce, sh.rgcb, { ...ctx, base: { r, c }, base3d: { r, c } });
        return undefined; // формула массива — у якоря (ARRAY)
      }
      return decompile(rgce, rgcb, { ...ctx, base: { r, c }, base3d: null });
    };
    const formulaCells = [];
    // записи для переноса при сохранении в .xls: зона → байты записей (с заголовками) в исходном порядке
    const kz = { pre: [], obj: [], prot: [], pls: [], hf: [], tail: [], labels: new Set(), chart: false };
    let lastZone = null;
    for (k += 1; k < recs.length; k++) {
      const { type, data: d } = recs[k];
      if (type === R.EOF) break;
      if (type === R.BOF) { // вложенная часть (диаграмма на листе) — переносится целиком
        let depth = 1;
        kz.obj.push(rawRec(recs[k]));
        while (depth && ++k < recs.length) { kz.obj.push(rawRec(recs[k])); if (recs[k].type === R.BOF) depth++; else if (recs[k].type === R.EOF) depth--; }
        kz.chart = true; kz.labels.add('диаграммы'); lastZone = null;
        continue;
      }
      // зона записи для переноса; CONTINUE — вслед за своей записью
      const zone = type === R.CONTINUE ? lastZone : KEEP_PRE.has(type) ? 'pre' : KEEP_OBJ.has(type) ? 'obj' : KEEP_PROT.has(type) ? 'prot'
        : KEEP_TAIL.has(type) ? 'tail' : type === 0x004D ? 'pls' : ((type === 0x0014 || type === 0x0015) && d.length && u16(d, 0)) || type === 0x089C ? 'hf' : null;
      if (zone) { kz[zone].push(rawRec(recs[k])); if (KEEP_LABEL[type]) kz.labels.add(KEEP_LABEL[type]); }
      lastZone = zone;
      if (type === 0x00B0) keep('сводные таблицы');
      switch (type) {
        case R.COLINFO: {
          const c1 = u16(d, 0); const c2 = Math.min(u16(d, 2), 255); const w = u16(d, 4) / 256; const fl = u16(d, 8);
          const px = widthToPx(w);
          for (let c = c1; c <= c2; c++) {
            if (px !== DEFAULT_COL_W && !((fl & 1) && !w)) sheet.colW[c] = px; // скрытый — ширина 0 в файле, не настоящая
            const li = {};
            if (fl & 1) li.hidden = true;
            if (fl & 4) li.bestFit = true;
            if ((fl >> 8) & 7) li.level = (fl >> 8) & 7;
            if (fl & 0x1000) li.collapsed = true;
            if (Object.keys(li).length) sheet.colInfo[c] = li;
          }
          break;
        }
        case R.ROW: {
          const r = u16(d, 0); const ht = u16(d, 6) & 0x7FFF; const fl = u16(d, 12);
          if (fl & 0x40) sheet.rowH[r] = Math.round(ht / 20 * 4 / 3);
          const li = {};
          if (fl & 0x20) li.hidden = true;
          if (fl & 7) li.level = fl & 7;
          if (fl & 0x10) li.collapsed = true;
          if (Object.keys(li).length) sheet.rowInfo[r] = li;
          break;
        }
        case R.BLANK: setCell(u16(d, 0), u16(d, 2), withStyle({ v: null }, u16(d, 4))); break;
        case R.MULBLANK: { const r = u16(d, 0); const c1 = u16(d, 2); const n = (d.length - 6) / 2; for (let j = 0; j < n; j++) setCell(r, c1 + j, withStyle({ v: null }, u16(d, 4 + j * 2))); break; }
        case R.NUMBER: setCell(u16(d, 0), u16(d, 2), withStyle({ v: numVal(f64(d, 6), u16(d, 4)) }, u16(d, 4))); break;
        case R.RK: setCell(u16(d, 0), u16(d, 2), withStyle({ v: numVal(rk(u32(d, 6)), u16(d, 4)) }, u16(d, 4))); break;
        case R.MULRK: {
          const r = u16(d, 0); const c1 = u16(d, 2); const n = (d.length - 6) / 6;
          for (let j = 0; j < n; j++) { const ix = u16(d, 4 + j * 6); setCell(r, c1 + j, withStyle({ v: numVal(rk(u32(d, 6 + j * 6)), ix) }, ix)); }
          break;
        }
        case R.LABELSST: setCell(u16(d, 0), u16(d, 2), withStyle({ v: sst[u32(d, 6)] ?? '' }, u16(d, 4))); break;
        case R.LABEL: case R.RSTRING: setCell(u16(d, 0), u16(d, 2), withStyle({ v: xlString(d, 6)[0] }, u16(d, 4))); break;
        case R.BOOLERR: setCell(u16(d, 0), u16(d, 2), withStyle({ v: d[7] ? { error: ERRS[d[6]] || '#N/A' } : d[6] === 1 }, u16(d, 4))); break;
        case R.SHRFMLA: {
          const cce = u16(d, 8);
          shared.push({ r1: u16(d, 0), r2: u16(d, 2), c1: d[4], c2: d[5], rgce: d.subarray(10, 10 + cce), rgcb: d.subarray(10 + cce) });
          break;
        }
        case R.ARRAY: {
          const cce = u16(d, 12);
          arrays.push({ r1: u16(d, 0), r2: u16(d, 2), c1: d[4], c2: d[5], rgce: d.subarray(14, 14 + cce), rgcb: d.subarray(14 + cce) });
          break;
        }
        case R.FORMULA: {
          const r = u16(d, 0); const c = u16(d, 2); const ix = u16(d, 4);
          const cell = withStyle({ v: null }, ix);
          if (u16(d, 12) === 0xFFFF) {
            const kind = d[6];
            if (kind === 0) pendingString = cell;
            else if (kind === 1) cell.v = d[8] === 1;
            else if (kind === 2) cell.v = { error: ERRS[d[8]] || '#N/A' };
            else if (kind === 3) cell.v = '';
          } else cell.v = numVal(f64(d, 6), ix);
          const cce = u16(d, 20);
          formulaCells.push({ r, c, cell, rgce: d.subarray(22, 22 + cce), rgcb: d.subarray(22 + cce) });
          sheet.cells.set(`${r},${c}`, cell);
          break;
        }
        case R.STRING: if (pendingString) { pendingString.v = xlString(d, 0)[0]; pendingString = null; } break;
        case R.MERGECELLS: { const n = u16(d, 0); for (let j = 0; j < n; j++) { const o = 2 + j * 8; sheet.merges.push({ r1: u16(d, o), r2: u16(d, o + 2), c1: u16(d, o + 4), c2: u16(d, o + 6) }); } break; }
        case R.WINDOW2: sheet.__frozen = (u16(d, 0) & 0x08) !== 0; break;
        case R.PANE: if (sheet.__frozen) { const cols = u16(d, 0); const rows = u16(d, 2); if (rows || cols) sheet.freeze = { rows, cols }; } break;
        case R.SETUP: {
          const paper = u16(d, 0); const scale = u16(d, 2); const fw = u16(d, 6); const fh = u16(d, 8); const fl = u16(d, 10);
          if (!(fl & 0x04)) { // fNoPls — параметры принтера не заданы
            if (paper) page.paper = paper;
            if (!(fl & 0x02)) page.orientation = 'landscape';
            if (scale && scale !== 100) page.scale = scale;
          }
          page.__fit = { fw, fh };
          if (d.length >= 32) { (page.margins ||= {}).header = f64(d, 16); page.margins.footer = f64(d, 24); }
          break;
        }
        case R.WSBOOL: fitToPage = (u16(d, 0) & 0x0100) !== 0; break;
        case R.LEFTMARGIN: (page.margins ||= {}).left = f64(d, 0); break;
        case R.RIGHTMARGIN: (page.margins ||= {}).right = f64(d, 0); break;
        case R.TOPMARGIN: (page.margins ||= {}).top = f64(d, 0); break;
        case R.BOTTOMMARGIN: (page.margins ||= {}).bottom = f64(d, 0); break;
        case R.PRINTGRIDLINES: if (u16(d, 0)) page.gridLines = true; break;
        case R.HCENTER: if (u16(d, 0)) page.centerH = true; break;
        case R.PROTECT: locked = u16(d, 0) === 1; break;
        case R.PASSWORD: { const h = u16(d, 0); if (h) pass = h.toString(16).toUpperCase(); break; }
        default: break;
      }
    }
    delete sheet.__frozen;
    // формулы — после всех SHRFMLA/ARRAY листа (запись общей формулы идёт после первой её ячейки)
    for (const { r, c, cell, rgce, rgcb } of formulaCells) {
      const arr = rgce[0] === 0x01 && arrays.find((x) => x.r1 === u16(rgce, 1) && x.c1 === u16(rgce, 3));
      if (arr) {
        if (r === arr.r1 && c === arr.c1) {
          const f = canon(decompile(arr.rgce, arr.rgcb, { ...ctx, base: { r, c }, base3d: null }));
          if (f != null) {
            cell.f = f; cell.cse = `${colName(arr.c1)}${arr.r1 + 1}:${colName(arr.c2)}${arr.r2 + 1}`;
            Object.defineProperty(cell, RECOMPUTE, { value: true }); // массив вычислить при открытии (в файле — только значения ячеек)
          }
        } else if (!cell.s) sheet.cells.delete(`${r},${c}`); // значение внутри формулы массива — её посчитает якорь
        else cell.v = null;
        continue;
      }
      const f = canon(formulaText(rgce, rgcb, r, c));
      if (f == null) { keep('формулы, которые Перо не прочитало'); continue; }
      cell.f = f;
      cell.legacy = true; // обычная формула: результат-массив Excel сводит к одному значению (как у .xlsx)
    }
    if (fitToPage && page.__fit) { page.fitWidth = page.__fit.fw; page.fitHeight = page.__fit.fh; }
    delete page.__fit;
    if (Object.keys(page).length) sheet.page = page;
    // перенос при сохранении в .xls: index — место листа в исходной книге (ссылки диаграмм — по нему), page — для PLS
    if ([kz.pre, kz.obj, kz.prot, kz.pls, kz.hf, kz.tail].some((z) => z.length)) {
      sheet.xlsKeep = { index: si, pre: kz.pre, obj: kz.obj, prot: kz.prot, pls: kz.pls, hf: kz.hf, tail: kz.tail, chart: kz.chart, labels: [...kz.labels], page: JSON.stringify(sheet.page || null) };
      for (const l of kz.labels) saved[l] = (saved[l] || 0) + 1;
    }
    if (locked) sheet.protection = { attrs: { sheet: '1', objects: '1', scenarios: '1', ...(pass ? { password: pass } : {}) } };
    // область печати и сквозные строки (встроенные имена листа)
    for (const n of printNames.filter((x) => x.itab === si + 1)) {
      const text = decompile(n.rgce, n.rgcb, ctx);
      if (text == null) continue;
      if (printKind(n) === '\u00016') sheet.printArea = parsePrintArea(text);
      if (printKind(n) === '\u00017') sheet.printTitles = parsePrintTitles(text);
    }
  }
  // имена книги (без встроенных и скрытых служебных)
  // служебные имена: функции новее Excel 2003 (_xlfn.), параметры LET/LAMBDA (_xlpm.), пустые заглушки
  const wbNames = names.filter((n) => !n.builtin && !n.fn && !printKind(n) && n.rgce.length && !/^(_FilterDatabase|_xlfn\.|_xlpm\.)/i.test(n.name)).map((n) => {
    const text = decompile(n.rgce, n.rgcb, ctx);
    return text == null ? null : { name: n.name, text: canon(text), sheet: n.itab ? sheetNames[n.itab - 1] ?? null : null };
  }).filter(Boolean);
  if (!sheets.length) sheets.push(newSheet('Лист1'));
  const base = { sheets, active: Math.min(activeTab, sheets.length - 1), ...(wbNames.length ? { names: wbNames } : {}), ...(manual ? { calcMode: 'manual' } : {}) };
  // сохранение в .xls поверх исходной книги: порядок EXTERNSHEET и NAME — как в файле (на их номера ссылаются
  // перенесённые записи: ряды диаграмм, формулы проверки), общая часть рисунков, свойства документа
  if (sheets.some((s) => s.xlsKeep)) {
    base.xlsKeep = {
      books: supbooks.map((b) => (b.self ? 'self' : b.addin ? { addin: b.names } : 'other')),
      xti: xti.map((x) => ({ ...x })),
      names: names.map((n) => ({ name: n.name, code: n.builtin ? +n.name.slice(1) : null, fn: n.fn, itab: n.itab, data: n.data })),
      drawingGroup,
      // шрифты (вид — как у стиля ячейки в записи), форматы, палитра — на прежних номерах
      fonts: fonts.filter((f, k) => k !== 4).map((f, k) => ({ desc: { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: !!f.u, strike: f.strike, color: color(f.icv) }, raw: fontRaw[k] })),
      formats: fmtRaw,
      palette: palette.slice(8, 64),
      streams: cfb.entries.filter((e) => e.type === 'stream' && /SummaryInformation$/.test(e.path)).map((e) => ({ path: e.path, data: e.data })),
    };
  }
  // формулы массива (cse) — вычислить при открытии; остальное — значения из файла, как у .xlsx
  const wb = doRecalc ? recalcOpen(base) : base;
  // warnings — что теряется при сохранении; kept — что переносится в .xls как есть
  return { wb, warnings: kept, kept: saved };
}

// Текст формулы → каноническая запись: функции надстроек (пакет анализа) русский Excel пишет в .xls русскими именами
const canon = (f) => { if (f == null) return f; try { return canonical(f); } catch { return f; } };

// RK: число, сжатое в 30 бит (целое или старшие биты double), возможно ×100
function rk(v) {
  let n;
  if (v & 2) n = (v | 0) >> 2;
  else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xFFFFFFFC, true); b.setUint32(0, 0, true); n = b.getFloat64(0, true); }
  return v & 1 ? n / 100 : n;
}

// встроенные имена (номер символа → имя)
const BUILTIN_NAMES = { 0: 'Consolidate_Area', 1: 'Auto_Open', 2: 'Auto_Close', 3: 'Extract', 4: 'Database', 5: 'Criteria', 6: 'Print_Area', 7: 'Print_Titles', 8: 'Recorder', 9: 'Data_Form', 10: 'Auto_Activate', 11: 'Auto_Deactivate', 12: 'Sheet_Title', 13: '_FilterDatabase' };
