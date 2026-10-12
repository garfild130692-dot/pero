// Импорт/экспорт .xlsx — собственный код (JSZip + xmldom), как и для .docx.
import JSZip from 'jszip';
import { precompress } from '../docx/zipBytes.js';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { addr, colIndex, key, MAX_COLS, parseRange, rangeName, unkey } from './address.js';
import { canonical, fileFormula, isErr, offsetFormula, shiftFormula } from './formula.js';
import { BUILTIN_FMT } from './format.js';
import { maxLevel, newSheet, recalc, recalcOpen, RECOMPUTE, DEFAULT_COL_W, DEFAULT_ROW_H } from './model.js';
import { parsePrintArea, parsePrintTitles, printAreaText, printTitlesText } from './pageSetup.js';
import { filterColumnXml, readFilterColumn } from './filter.js';
import { cfSignature, cfXml, dxfXml, readCf, readDxf } from './cf.js';
import { dvSignature, dvXml, readDv } from './dv.js';
import { parseSheet } from './sheetScan.js';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ln = (n) => n.localName || String(n.nodeName).split(':').pop();
const kids = (el, name) => { const out = []; if (!el) return out; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && (!name || ln(c) === name)) out.push(c); return out; };
const kid = (el, name) => kids(el, name)[0] || null;
const parseXml = (t) => new DOMParser().parseFromString(t, 'text/xml');
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const resolve = (dir, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const out = [];
  for (const x of `${dir}/${target}`.split('/')) { if (x === '..') out.pop(); else if (x && x !== '.') out.push(x); }
  return out.join('/');
};

// Ширина столбца ↔ px по формулам спецификации (ширина цифры Calibri 11 — 7 px; в файле ширина с полями)
export const widthToPx = (w) => Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
const pxToWidth = (px) => Math.trunc((px / 7) * 256) / 256;

// ---------- Цвета ----------
// Стандартная палитра индексных цветов Excel (0–63)
export const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
// Тон (tint) как в Excel: осветление/затемнение по светлоте HSL
export function applyTint(hex, tint) {
  if (!tint) return hex;
  let r = parseInt(hex.slice(0, 2), 16) / 255; let g = parseInt(hex.slice(2, 4), 16) / 255; let b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b); const min = Math.min(r, g, b);
  let h = 0; let s = 0; let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (s === 0) { r = g = b = l; } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
    r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3);
  }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function readColor(el, theme) {
  if (!el) return null;
  const rgb = el.getAttribute('rgb');
  const tint = +(el.getAttribute('tint') || 0);
  let hex = null;
  if (rgb) hex = rgb.slice(-6);
  else if (el.getAttribute('theme') != null) hex = theme[+el.getAttribute('theme')] || null;
  else if (el.getAttribute('indexed') != null) { const i = +el.getAttribute('indexed'); hex = i === 64 || i === 65 ? null : INDEXED[i] || null; }
  return hex ? `#${applyTint(hex.toUpperCase(), tint).toLowerCase()}` : null;
}
// Цвета темы в порядке индексов Excel: lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink
function parseTheme(doc) {
  const def = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  if (!doc) return def;
  const scheme = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'clrScheme')[0];
  if (!scheme) return def;
  const val = (name) => {
    const el = kid(scheme, name);
    const c = el && kids(el)[0];
    return c ? (c.getAttribute('lastClr') || c.getAttribute('val') || null) : null;
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((n, i) => val(n === 'lt1' ? 'lt1' : n) || def[i]).map((x) => (x || '').toUpperCase());
}

// ---------- Стили ----------
function parseStyles(doc, theme) {
  if (!doc) return [];
  const root = doc.documentElement;
  const fmts = { ...BUILTIN_FMT };
  for (const f of kids(kid(root, 'numFmts'), 'numFmt')) fmts[+f.getAttribute('numFmtId')] = f.getAttribute('formatCode');
  const fonts = kids(kid(root, 'fonts'), 'font').map((f) => {
    const v = (n) => kid(f, n);
    const on = (n) => { const e = v(n); return !!e && !['0', 'false'].includes(e.getAttribute('val')); };
    const sz = v('sz') ? +v('sz').getAttribute('val') : null;
    const name = v('name') ? v('name').getAttribute('val') : null;
    return {
      bold: on('b'), italic: on('i'), underline: !!v('u') && v('u').getAttribute('val') !== 'none', strike: on('strike'),
      fontSize: sz && sz !== 11 ? sz : null, fontFamily: name && name !== 'Calibri' ? name : null,
      // цвет текста темы (theme=1 без оттенка) — «автоматический», не задан
      color: v('color') && v('color').getAttribute('theme') === '1' && !v('color').getAttribute('tint') ? null : readColor(v('color'), theme),
    };
  });
  const fills = kids(kid(root, 'fills'), 'fill').map((f) => {
    const p = kid(f, 'patternFill');
    if (!p || !['solid'].includes(p.getAttribute('patternType'))) return null;
    return readColor(kid(p, 'fgColor'), theme);
  });
  const borders = kids(kid(root, 'borders'), 'border').map((b) => {
    const out = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const e = kid(b, side);
      const style = e && e.getAttribute('style');
      if (style) out[side] = { style, color: readColor(kid(e, 'color'), theme) };
    }
    return Object.keys(out).length ? out : null;
  });
  return kids(kid(root, 'cellXfs'), 'xf').map((xf) => {
    const s = {};
    const font = fonts[+xf.getAttribute('fontId') || 0] || {};
    for (const [k, v] of Object.entries(font)) if (v) s[k] = v;
    const fill = fills[+xf.getAttribute('fillId') || 0];
    if (fill) s.fill = fill;
    const border = borders[+xf.getAttribute('borderId') || 0];
    if (border) s.border = border;
    const nf = fmts[+xf.getAttribute('numFmtId') || 0];
    if (nf && nf !== 'General') s.numFmt = nf === 'mm-dd-yy' ? 'dd.mm.yyyy' : nf;
    const al = kid(xf, 'alignment');
    if (al) {
      const h = al.getAttribute('horizontal');
      if (['left', 'center', 'right', 'justify'].includes(h)) s.align = h;
      else if (h === 'centerContinuous') s.align = 'center';
      const v = al.getAttribute('vertical');
      if (v === 'center' || v === 'top') s.valign = v;
      if (['1', 'true'].includes(al.getAttribute('wrapText'))) s.wrap = true;
    }
    const pr = kid(xf, 'protection');
    if (pr && ['0', 'false'].includes(pr.getAttribute('locked'))) s.unlocked = true;
    if (pr && ['1', 'true'].includes(pr.getAttribute('hidden'))) s.formulaHidden = true;
    return Object.keys(s).length ? s : null;
  });
}

// ---------- Чтение пакета (общее для импорта и экспорта поверх исходного файла) ----------
async function readBook(zip) {
  const text = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };
  const read = async (p) => { const t = await text(p); return t != null ? parseXml(t) : null; };
  const rootRels = await read('_rels/.rels');
  let wbPath = 'xl/workbook.xml';
  for (const rel of kids(rootRels && rootRels.documentElement, 'Relationship')) {
    if (/officeDocument$/.test(rel.getAttribute('Type'))) wbPath = resolve('', rel.getAttribute('Target'));
  }
  const wbDoc = await read(wbPath);
  if (!wbDoc) throw new Error('Это не книга Excel: нет xl/workbook.xml');
  const dir = wbPath.slice(0, wbPath.lastIndexOf('/'));
  const relsPath = `${dir}/_rels/${wbPath.slice(dir.length + 1)}.rels`;
  const relsDoc = await read(relsPath);
  const rels = {};
  for (const rel of kids(relsDoc && relsDoc.documentElement, 'Relationship')) rels[rel.getAttribute('Id')] = { type: rel.getAttribute('Type'), target: resolve(dir, rel.getAttribute('Target')) };
  const relOf = (type) => Object.values(rels).find((r) => r.type.endsWith(`/${type}`));

  const theme = parseTheme(relOf('theme') ? await read(relOf('theme').target) : null);
  const stylesPath = relOf('styles') ? relOf('styles').target : null;
  const stylesDoc = stylesPath ? await read(stylesPath) : null;
  const styles = parseStyles(stylesDoc, theme);
  // стили условного форматирования (dxfs)
  const colorOf = (el) => readColor(el, theme);
  const dxfs = stylesDoc ? kids(kid(stylesDoc.documentElement, 'dxfs'), 'dxf').map((d) => readDxf(d, colorOf)) : [];
  const sstPath = relOf('sharedStrings') ? relOf('sharedStrings').target : null;
  const sstDoc = sstPath ? await read(sstPath) : null;
  const sstEls = sstDoc ? kids(sstDoc.documentElement, 'si') : [];
  const sst = sstEls.map((si) => {
    const t = kid(si, 't');
    if (t) return t.textContent;
    return kids(si, 'r').map((r) => (kid(r, 't') || {}).textContent || '').join('');
  });
  // строки без оформления — их можно переиспользовать для новых значений
  const sstPlain = sstEls.map((si) => !!kid(si, 't') && !kid(si, 'r'));
  const date1904 = (() => { const p = kid(wbDoc.documentElement, 'workbookPr'); return !!p && ['1', 'true'].includes(p.getAttribute('date1904')); })();
  const sheets = kids(kid(wbDoc.documentElement, 'sheets'), 'sheet').map((se) => {
    const rid = se.getAttributeNS(R, 'id') || se.getAttribute('r:id');
    const rel = rels[rid];
    return { name: se.getAttribute('name'), rid, target: rel ? rel.target : null, worksheet: !!rel && /\/worksheet$/.test(rel.type) };
  });
  return { text, read, wbPath, wbDoc, dir, relsPath, rels, stylesPath, stylesDoc, styles, dxfs, colorOf, sstPath, sst, sstPlain, date1904, sheets };
}

// Ячейки листа: on.cell(r, c, cell, el, shared) — для каждой непустой <c>; on.row(r, el) — для каждой <row>
// Ячейки из файла: sheet.src.cells — карта ячеек, как её прочитал импорт (объекты ячеек модель не меняет: правка —
// новый объект). Ячейка, которая и сейчас тот же объект по тому же ключу, при сохранении поверх исходного пакета
// пишется исходной разметкой без разбора и сравнения. Общие формулы и формулы массива (src.noRaw) — заново: их
// разметка зависит от соседей.
// Формула из файла → каноническая запись. Записанная Excel (английские имена, «,», без _xlfn. и строчных латинских
// букв) уже каноническая — без разбора на лексемы (проверено на 80 тыс. формул эталонных книг: совпадает)
const canonFile = (t) => (/[a-z_;{]/.test(t) || t.includes('ANCHORARRAY') ? canonical(t) : t);

// Строки листа: из сканера (parseSheet, sheetScan.js) или из DOM — в одном виде { el, cells: [{ el, v, f, is }] }
function sheetRows(root) {
  if (root.__sheetData) return root.__sheetData;
  return kids(kid(root, 'sheetData'), 'row').map((row) => ({
    el: row,
    cells: kids(row, 'c').map((c) => {
      const v = kid(c, 'v'); const f = kid(c, 'f'); const is = kid(c, 'is');
      const ts = is ? kids(is, 't').map((x) => x.textContent) : [];
      return { el: c, v: v ? v.textContent : null, f: f ? { el: f, text: f.textContent } : null, is: is ? { direct: ts.join(''), hasDirect: ts.length > 0, all: is.textContent } : null };
    }),
  }));
}

function readSheetData(root, ctx, on) {
  const { sst, styles, date1904 } = ctx;
  const shared = {}; // si → { f, r, c }
  const ghosts = []; // диапазоны «разлива»/формул массива: значения в них считает формула-якорь
  const inGhost = (r, c) => ghosts.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2 && !(r === g.r1 && c === g.c1));
  for (const { el: row, cells } of sheetRows(root)) {
    const ri = +row.getAttribute('r') - 1;
    if (on.row) on.row(ri, row);
    let cAuto = 0;
    for (const { el: c, v: vt, f: fx, is } of cells) {
      const ref = c.getAttribute('r');
      let r = ri; let ci = cAuto;
      if (ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); ci = colIndex(m[1]); r = +m[2] - 1; }
      cAuto = ci + 1;
      const t = c.getAttribute('t') || 'n';
      const sIdx = +(c.getAttribute('s') || 0);
      const fEl = fx && fx.el;
      const cell = {};
      const st = styles[sIdx];
      if (st) cell.s = st;
      // значение
      if (t === 's') cell.v = sst[+vt] ?? '';
      else if (t === 'inlineStr') cell.v = is ? (is.direct || is.all) : '';
      else if (t === 'str') cell.v = vt ?? '';
      else if (t === 'b') cell.v = vt === '1';
      else if (t === 'e') cell.v = { error: vt || '#N/A' };
      else if (t === 'd') { const d = new Date(vt); cell.v = isNaN(d) ? vt : (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000; }
      else if (vt != null && vt !== '') {
        cell.v = +vt;
        if (date1904 && st && st.numFmt && /[dy]/i.test(st.numFmt)) cell.v += 1462;
      } else cell.v = null;
      // формула (общие формулы — развернуть со смещением от главной ячейки)
      let isShared = false;
      let isArr = false;
      if (fEl) {
        const ft = fEl.getAttribute('t');
        const text = fx.text;
        if (ft === 'shared') {
          isShared = true;
          const si = fEl.getAttribute('si');
          if (text) { shared[si] = { f: canonFile(text), r, c: ci }; cell.f = shared[si].f; }
          else if (shared[si]) cell.f = offsetFormula(shared[si].f, r - shared[si].r, ci - shared[si].c);
        } else if (text) cell.f = canonFile(text);
        if (cell.f != null) {
          if (ft === 'array') {
            isArr = true;
            // t="array": с cm — динамический массив Excel 365, без cm — формула массива Ctrl+Shift+Enter
            const ref = fEl.getAttribute('ref') || addr(r, ci);
            const rg = parseRange(ref);
            if (!c.getAttribute('cm')) cell.cse = ref;
            if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) ghosts.push(rg);
          } else cell.legacy = true; // обычная формула: результат-массив Excel сводит к одному значению
        }
      } else if (inGhost(r, ci)) {
        // значение внутри «разлива» — его посчитает формула; оформление ячейки сохраняется
        if (!cell.s) continue;
        cell.v = null;
      }
      if (cell.v == null && cell.f == null && !cell.s) continue;
      if (cell.v === undefined) cell.v = null;
      on.cell(r, ci, cell, c, isShared, isArr);
    }
  }
}

const isOn = (el, a) => ['1', 'true'].includes(el.getAttribute(a));
// Закрепление областей первого представления листа: { rows, cols } или null (разделение без закрепления — не считаем)
function readFreeze(root) {
  const view = kid(kid(root, 'sheetViews'), 'sheetView');
  const pane = view && kid(view, 'pane');
  if (!pane || !/^frozen/.test(pane.getAttribute('state') || '')) return null;
  const rows = Math.round(+(pane.getAttribute('ySplit') || 0));
  const cols = Math.round(+(pane.getAttribute('xSplit') || 0));
  return rows || cols ? { rows, cols } : null;
}
// Параметры страницы листа (только то, что записано в файле)
function readPage(root) {
  const page = {};
  const ps = kid(root, 'pageSetup');
  if (ps) {
    if (ps.getAttribute('orientation') === 'landscape') page.orientation = 'landscape';
    if (ps.getAttribute('paperSize')) page.paper = +ps.getAttribute('paperSize');
    const sc = +(ps.getAttribute('scale') || 100);
    if (sc !== 100) page.scale = sc;
  }
  const spr = kid(kid(root, 'sheetPr'), 'pageSetUpPr');
  if (spr && isOn(spr, 'fitToPage')) {
    page.fitWidth = ps && ps.getAttribute('fitToWidth') != null ? +ps.getAttribute('fitToWidth') : 1;
    page.fitHeight = ps && ps.getAttribute('fitToHeight') != null ? +ps.getAttribute('fitToHeight') : 1;
  }
  const pm = kid(root, 'pageMargins');
  if (pm) {
    const m = {};
    for (const n of ['left', 'right', 'top', 'bottom', 'header', 'footer']) if (pm.getAttribute(n) != null) m[n] = +pm.getAttribute(n);
    page.margins = m;
  }
  const po = kid(root, 'printOptions');
  if (po && isOn(po, 'gridLines')) page.gridLines = true;
  if (po && isOn(po, 'horizontalCentered')) page.centerH = true;
  return Object.keys(page).length ? page : null;
}
// Защита листа: { attrs } — все атрибуты <sheetProtection> (хеш пароля и разрешения), без sheet="1" — не защищён
function readProtection(root) {
  const el = kid(root, 'sheetProtection');
  if (!el || !isOn(el, 'sheet')) return null;
  const attrs = {};
  for (let i = 0; i < el.attributes.length; i++) attrs[el.attributes[i].name] = el.attributes[i].value;
  return { attrs };
}
// Ручные разрывы страниц: { rows: [номер первой строки новой страницы], cols: [...] } (brk id — с нуля)
function readBreaks(root) {
  const ids = (tag) => kids(kid(root, tag), 'brk').filter((b) => b.getAttribute('man') !== '0').map((b) => +b.getAttribute('id')).filter((n) => n > 0);
  const rows = ids('rowBreaks'); const cols = ids('colBreaks');
  return rows.length || cols.length ? { rows, cols } : null;
}
const breaksXml = (b, tag, max) => {
  const list = (b && b[tag === 'rowBreaks' ? 'rows' : 'cols']) || [];
  return list.length ? `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((id) => `<brk id="${id}" max="${max}" man="1"/>`).join('')}</${tag}>` : '';
};
// Условное форматирование листа (cf.js)
const serializeEl = (el) => new XMLSerializer().serializeToString(el).replace(/\sxmlns(:\w+)?="[^"]*"/g, '');
const readSheetDv = (root) => readDv(root, { parseRange, serialize: serializeEl });
const readSheetCf = (root, book) => readCf(root, { dxfs: book.dxfs, readColor: book.colorOf, parseRange, serialize: serializeEl });
// Автофильтр: { ref, cols, raw } (filter.js); непонятные условия столбцов (top10, по цвету…) — исходным XML
function readAutoFilter(root) {
  const el = kid(root, 'autoFilter');
  const ref = el && parseRange(el.getAttribute('ref') || '');
  if (!ref) return null;
  const af = { ref, cols: {} };
  for (const fc of [...el.childNodes].filter((n) => n.nodeType === 1 && (n.localName || n.nodeName) === 'filterColumn')) {
    const j = +fc.getAttribute('colId');
    const crit = readFilterColumn(fc);
    if (crit) af.cols[j] = crit;
    else (af.raw = af.raw || {})[j] = new XMLSerializer().serializeToString(fc).replace(/\sxmlns(:\w+)?="[^"]*"/g, '');
  }
  return af;
}
function autoFilterXml(af) {
  if (!af) return '';
  const cols = new Map();
  for (const [j, crit] of Object.entries(af.cols || {})) cols.set(+j, filterColumnXml(+j, crit));
  for (const [j, xml] of Object.entries(af.raw || {})) if (!cols.has(+j)) cols.set(+j, xml);
  const body = [...cols.keys()].sort((a, b) => a - b).map((j) => cols.get(j)).join('');
  return `<autoFilter ref="${rangeName(af.ref)}"${body ? `>${body}</autoFilter>` : '/>'}`;
}
const protectionXml = (p) => (p ? `<sheetProtection${Object.entries(p.attrs).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>` : '');

// Именованные диапазоны и константы книги (кроме служебных _xlnm.*): { name, text (хранимый вид), sheet (локальное — имя листа) }
function readNames(wbDoc, bookSheets) {
  return kids(kid(wbDoc.documentElement, 'definedNames'), 'definedName')
    .filter((n) => !/^_xlnm\./.test(n.getAttribute('name') || ''))
    .map((n) => {
      const id = n.getAttribute('localSheetId');
      const sheet = id != null && bookSheets[+id] ? bookSheets[+id].name : null;
      return { name: n.getAttribute('name'), text: canonical(n.textContent), sheet };
    });
}
// <definedName> пользовательских имён модели; номер листа — по порядку листов книги
function namesXml(wb) {
  return (wb.names || []).map((n) => {
    const idx = n.sheet ? wb.sheets.findIndex((x) => x.name.toLowerCase() === n.sheet.toLowerCase()) : -1;
    if (n.sheet && idx < 0) return '';
    return `<definedName name="${esc(n.name)}"${idx >= 0 ? ` localSheetId="${idx}"` : ''}>${xmlEsc(fileFormula(n.text))}</definedName>`;
  }).join('');
}

// Имена печати листа с номером idx (порядок в <sheets>): { printArea, printTitles }
function readPrintNames(wbDoc, idx) {
  const out = { printArea: null, printTitles: null };
  for (const n of kids(kid(wbDoc.documentElement, 'definedNames'), 'definedName')) {
    if (n.getAttribute('localSheetId') !== String(idx)) continue;
    if (n.getAttribute('name') === '_xlnm.Print_Area') out.printArea = parsePrintArea(n.textContent);
    if (n.getAttribute('name') === '_xlnm.Print_Titles') out.printTitles = parsePrintTitles(n.textContent);
  }
  return out;
}
// Атрибуты элемента (строка тега): null — удалить
function setAttrs(tag, attrs) {
  let t = tag;
  for (const [k, v] of Object.entries(attrs)) {
    t = t.replace(new RegExp(`\\s${k}="[^"]*"`), '');
    if (v != null) t = t.replace(/\s*(\/?)>$/, ` ${k}="${v}"$1>`);
  }
  return t;
}
const pageSetupAttrs = (page) => {
  const fit = page.fitWidth != null || page.fitHeight != null;
  return {
    paperSize: page.paper ?? null, scale: page.scale && page.scale !== 100 ? page.scale : null,
    fitToWidth: fit ? page.fitWidth ?? 0 : null, fitToHeight: fit ? page.fitHeight ?? 0 : null,
    orientation: page.orientation === 'landscape' ? 'landscape' : page.orientation ? 'portrait' : null,
  };
};
const marginsXml = (m) => `<pageMargins left="${m.left ?? 0.7}" right="${m.right ?? 0.7}" top="${m.top ?? 0.75}" bottom="${m.bottom ?? 0.75}" header="${m.header ?? 0.3}" footer="${m.footer ?? 0.3}"/>`;
// Параметры страницы поверх исходного XML листа: прочие атрибуты (r:id настроек принтера и т. п.) — как были
function setPageXml(xml, page) {
  let out = xml;
  const fit = page.fitWidth != null || page.fitHeight != null;
  // флаг «вписать» — в sheetPr/pageSetUpPr
  const sp = findEl(out, 'sheetPr');
  if (sp) {
    let el = out.slice(sp.start, sp.end);
    if (/<pageSetUpPr\b/.test(el)) el = el.replace(/<pageSetUpPr\b[^>]*?\/>/, (t) => setAttrs(t, { fitToPage: fit ? 1 : null }));
    else if (fit) el = el.includes('</sheetPr>') ? el.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>') : el.replace(/\s*\/>$/, '><pageSetUpPr fitToPage="1"/></sheetPr>');
    out = out.slice(0, sp.start) + el + out.slice(sp.end);
  } else if (fit) out = setSheetEl(out, 'sheetPr', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
  // печать сетки и центрирование
  const po = findEl(out, 'printOptions');
  const poAttrs = { gridLines: page.gridLines ? 1 : null, horizontalCentered: page.centerH ? 1 : null };
  if (po) {
    const t = setAttrs(out.slice(po.start, po.end), poAttrs);
    out = out.slice(0, po.start) + (/^<printOptions\s*\/>$/.test(t) ? '' : t) + out.slice(po.end);
  } else if (page.gridLines || page.centerH) out = setSheetEl(out, 'printOptions', setAttrs('<printOptions/>', poAttrs));
  if (page.margins) out = setSheetEl(out, 'pageMargins', marginsXml(page.margins));
  const ps = findEl(out, 'pageSetup');
  if (ps) out = out.slice(0, ps.start) + setAttrs(out.slice(ps.start, ps.end), pageSetupAttrs(page)) + out.slice(ps.end);
  else out = setSheetEl(out, 'pageSetup', setAttrs('<pageSetup/>', pageSetupAttrs(page)));
  return out;
}

// <pane> и <selection> для закрепления
function paneXml(f) {
  if (!f) return '';
  const act = f.rows && f.cols ? 'bottomRight' : f.rows ? 'bottomLeft' : 'topRight';
  return `<pane${f.cols ? ` xSplit="${f.cols}"` : ''}${f.rows ? ` ySplit="${f.rows}"` : ''} topLeftCell="${addr(f.rows, f.cols)}" activePane="${act}" state="frozen"/><selection pane="${act}"/>`;
}
// Свойства строки/столбца, которые хранит модель: скрыт, уровень группировки, свёрнут, стиль (номер xf исходного файла)
function readLineInfo(el, styleAttr) {
  const info = {};
  if (isOn(el, 'hidden')) info.hidden = true;
  const lvl = +(el.getAttribute('outlineLevel') || 0);
  if (lvl) info.level = lvl;
  if (isOn(el, 'collapsed')) info.collapsed = true;
  if (styleAttr === 'style' && isOn(el, 'bestFit')) info.bestFit = true;
  const s = el.getAttribute(styleAttr);
  if (s && +s && (styleAttr === 'style' || isOn(el, 'customFormat'))) info.xf = +s;
  return Object.keys(info).length ? info : null;
}

// Листы, открытые из файла и с тех пор не менявшиеся (модель неизменяема: правка — новый объект листа):
// при сохранении поверх исходного пакета их XML остаётся байт в байт
const pristine = new WeakSet();

// ---------- Импорт ----------
/** @returns {Promise<{ wb, warnings, kept }>} warnings — что потеряется при сохранении, kept — что сохраняется как есть */
// recalc: false — значения формул как записал Excel (сверка движка с Excel в тестах)
export async function importXlsx(data, { recalc: doRecalc = true } = {}) {
  const zip = await JSZip.loadAsync(data);
  const book = await readBook(zip);
  const { read, rels, wbDoc } = book;

  const kept = {};
  const keep = (k) => { kept[k] = (kept[k] || 0) + 1; };
  const sheets = [];
  for (const [idx, info] of book.sheets.entries()) {
    const sheet = newSheet(info.name);
    if (!info.worksheet) { keep('листы диаграмм'); continue; }
    const doc = parseSheet(await book.text(info.target)); // <sheetData> — сканером (sheetScan.js)
    if (!doc) continue;
    sheet.src = { part: info.target, rid: info.rid, name: info.name, cells: sheet.cells, noRaw: new Set() };
    const root = doc.documentElement;
    for (const col of kids(kid(root, 'cols'), 'col')) {
      const min = +col.getAttribute('min') - 1;
      const max = Math.min(+col.getAttribute('max') - 1, MAX_COLS - 1);
      const w = +col.getAttribute('width');
      const px = w && ['1', 'true'].includes(col.getAttribute('customWidth') || '1') ? widthToPx(w) : null;
      const li = readLineInfo(col, 'style');
      for (let c = min; c <= max; c++) {
        if (px != null && px !== DEFAULT_COL_W) sheet.colW[c] = px;
        if (li) sheet.colInfo[c] = li;
      }
    }
    readSheetData(root, book, {
      row: (ri, row) => {
        const ht = row.getAttribute('ht');
        if (ht && isOn(row, 'customHeight')) sheet.rowH[ri] = Math.round(+ht * 4 / 3);
        const li = readLineInfo(row, 's');
        if (li) sheet.rowInfo[ri] = li;
      },
      cell: (r, c, cell, el, shared, arr) => {
        if (shared || arr) sheet.src.noRaw.add(key(r, c));
        if (arr && cell.f != null) Object.defineProperty(cell, RECOMPUTE, { value: true }); // массив вычислить при открытии
        sheet.cells.set(key(r, c), cell);
      },
    });
    for (const m of kids(kid(root, 'mergeCells'), 'mergeCell')) {
      const rg = parseRange(m.getAttribute('ref'));
      if (rg) sheet.merges.push(rg);
    }
    const freeze = readFreeze(root);
    if (freeze) sheet.freeze = freeze;
    const page = readPage(root);
    if (page) sheet.page = page;
    const prot = readProtection(root);
    if (prot) sheet.protection = prot;
    const pn = readPrintNames(wbDoc, idx);
    if (pn.printArea) sheet.printArea = pn.printArea;
    if (pn.printTitles) sheet.printTitles = pn.printTitles;
    const cf = readSheetCf(root, book);
    if (cf.length) {
      sheet.cf = cf;
      if (cf.some((b) => b.rules.some((r) => r.unknown))) keep('наборы значков в условном форматировании');
    }
    const dv = readSheetDv(root);
    if (dv.length) sheet.dv = dv;
    const breaks = readBreaks(root);
    if (breaks) sheet.breaks = breaks;
    if (kid(root, 'drawing')) keep('рисунки и диаграммы');
    if (kid(root, 'tableParts')) keep('«умные» таблицы');
    if (kid(root, 'legacyDrawing')) keep('примечания к ячейкам');
    const af = readAutoFilter(root);
    if (af) {
      sheet.autoFilter = af;
      // скрытые строки под заголовком фильтра — скрыты фильтром (так их понимает Excel)
      for (const [r, info] of Object.entries(sheet.rowInfo || {})) if (info.hidden && +r > af.ref.r1 && +r <= af.ref.r2) sheet.rowInfo[r] = { ...info, filtered: true };
      if (af.raw) keep('условия фильтра по цвету, первым N и т. п.');
    }
    if (kid(root, 'hyperlinks')) keep('гиперссылки');
    sheets.push(sheet);
  }
  if (Object.values(rels).some((r) => /pivotCacheDefinition$/.test(r.type))) keep('сводные таблицы');
  // именованные диапазоны (служебные _xlnm.* — область печати, фильтр — не считаем)
  const names = readNames(wbDoc, book.sheets);
  if (!sheets.length) sheets.push(newSheet('Лист1'));
  const view = kid(kid(wbDoc.documentElement, 'bookViews'), 'workbookView');
  const active = Math.min(+(view && view.getAttribute('activeTab')) || 0, sheets.length - 1);
  const calcPr = kid(wbDoc.documentElement, 'calcPr');
  const manual = !!calcPr && calcPr.getAttribute('calcMode') === 'manual';
  const base = { sheets, active, ...(names.length ? { names } : {}), ...(manual ? { calcMode: 'manual' } : {}) };
  // при открытии — полный пересчёт (летучие функции, устаревший кеш), в том числе в режиме «вручную»
  const wb = doRecalc ? recalcOpen(base) : base;
  for (const sh of wb.sheets) if (sh.src) pristine.add(sh);
  if (wb.names) pristine.add(wb.names);
  return { wb, warnings: {}, kept };
}

// ---------- Экспорт ----------
const BUILTIN_BY_CODE = Object.fromEntries(Object.entries(BUILTIN_FMT).map(([id, code]) => [code, +id]));
const DEFAULT_FONT = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';

// Равенство стилей ячеек (порядок ключей не важен)
function sameStyle(a, b) {
  if (a === b) return true;
  if (!a || !b) return false;
  const ka = Object.keys(a); const kb = Object.keys(b);
  if (ka.length !== kb.length) return false;
  return ka.every((k) => (typeof a[k] === 'object' && a[k] ? JSON.stringify(a[k]) === JSON.stringify(b[k]) : a[k] === b[k]));
}
const sameValue = (a, b) => (isErr(a) || isErr(b) ? isErr(a) && isErr(b) && a.error === b.error : a === b);

// Таблица стилей: уникальные шрифты/заливки/границы/форматы → cellXfs.
// base (книга из файла) — стили файла остаются на своих местах, новые дописываются в конец:
// { styles: разобранные cellXfs, counts: { fonts, fills, borders, xfs }, fmts: { код: id } }
function buildStyles(wb, base = null) {
  const fonts = base ? Array(base.counts.fonts).fill(null) : [DEFAULT_FONT];
  const fills = base ? Array(base.counts.fills).fill(null) : ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = base ? Array(base.counts.borders).fill(null) : ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const xfs = base ? Array(base.counts.xfs).fill(null) : ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const fmtIds = base ? { ...base.fmts } : {};
  let nextFmt = Math.max(163, ...Object.values(fmtIds)) + 1;
  const numFmts = []; // новые { id, code }
  // стили условного форматирования: исходные — на местах, новые — в конец
  const dxfs = base ? Array(base.counts.dxfs || 0).fill(null) : [];
  const dxfSrc = new Map(); const dxfIndex = new Map();
  if (base && base.dxfs) base.dxfs.forEach((d, i) => dxfSrc.set(d, i));
  const dxfIdOf = (d) => {
    if (dxfSrc.has(d)) return dxfSrc.get(d);
    const k = JSON.stringify(d);
    if (!dxfIndex.has(k)) { dxfs.push(dxfXml(d)); dxfIndex.set(k, dxfs.length - 1); }
    return dxfIndex.get(k);
  };
  const index = new Map(); // JSON стиля → номер xf
  const srcIndex = new Map(); // стиль файла (объект) → его номер xf
  if (base) base.styles.forEach((s, i) => { if (s && !srcIndex.has(s)) srcIndex.set(s, i); });
  const idxOf = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
  const color = (c) => (c ? `<color rgb="FF${c.slice(1).toUpperCase()}"/>` : '');
  const xfOf = (s) => {
    if (!s) return 0;
    if (srcIndex.has(s)) return srcIndex.get(s);
    const k = JSON.stringify(s);
    if (index.has(k)) return index.get(k);
    // такой же стиль уже есть в файле (ячейка скопирована, стиль собран заново)
    if (base) {
      const i = base.styles.findIndex((x) => x && sameStyle(x, s));
      if (i >= 0) { index.set(k, i); return i; }
    }
    const font = (s.bold || s.italic || s.underline || s.strike || s.color || s.fontSize || s.fontFamily || base)
      ? idxOf(fonts, `<font>${s.bold ? '<b/>' : ''}${s.italic ? '<i/>' : ''}${s.strike ? '<strike/>' : ''}${s.underline ? '<u/>' : ''}<sz val="${s.fontSize || 11}"/>${s.color ? color(s.color) : '<color theme="1"/>'}<name val="${esc(s.fontFamily || 'Calibri')}"/><family val="2"/></font>`)
      : 0;
    const fill = s.fill ? idxOf(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${s.fill.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    let border = 0;
    if (s.border) {
      const side = (n) => (s.border[n] ? `<${n} style="${s.border[n].style}">${color(s.border[n].color) || '<color indexed="64"/>'}</${n}>` : `<${n}/>`);
      border = idxOf(borders, `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`);
    }
    let numFmtId = 0;
    if (s.numFmt && s.numFmt !== 'General') {
      if (BUILTIN_BY_CODE[s.numFmt] != null && s.numFmt !== 'dd.mm.yyyy') numFmtId = BUILTIN_BY_CODE[s.numFmt];
      else if (fmtIds[s.numFmt] != null) numFmtId = fmtIds[s.numFmt];
      else { numFmtId = nextFmt++; fmtIds[s.numFmt] = numFmtId; numFmts.push({ id: numFmtId, code: s.numFmt }); }
    }
    const al = (s.align || s.valign || s.wrap)
      ? `<alignment${s.align ? ` horizontal="${s.align}"` : ''}${s.valign ? ` vertical="${s.valign}"` : ''}${s.wrap ? ' wrapText="1"' : ''}/>` : '';
    const prot = s.unlocked || s.formulaHidden ? `<protection${s.unlocked ? ' locked="0"' : ''}${s.formulaHidden ? ' hidden="1"' : ''}/>` : '';
    const inner = al + prot;
    const xml = `<xf numFmtId="${numFmtId}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}${border ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1"' : ''}${prot ? ' applyProtection="1"' : ''}${inner ? `>${inner}</xf>` : '/>'}`;
    xfs.push(xml);
    index.set(k, xfs.length - 1);
    return xfs.length - 1;
  };
  // проходим все ячейки и правила заранее, чтобы таблица стилей была полной
  for (const sh of wb.sheets) for (const cell of sh.cells.values()) xfOf(cell.s);
  for (const sh of wb.sheets) for (const b of sh.cf || []) for (const r of b.rules) if (r.dxf && (!r.raw || r.edited)) dxfIdOf(r.dxf);
  xfOf.dxf = dxfIdOf;
  const added = (arr, n) => arr.slice(n);
  if (base) {
    return {
      xfOf,
      added: {
        numFmts: numFmts.map((f) => `<numFmt numFmtId="${f.id}" formatCode="${esc(f.code)}"/>`),
        fonts: added(fonts, base.counts.fonts), fills: added(fills, base.counts.fills),
        borders: added(borders, base.counts.borders), cellXfs: added(xfs, base.counts.xfs),
        dxfs: added(dxfs, base.counts.dxfs || 0),
      },
    };
  }
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS}">` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f) => `<numFmt numFmtId="${f.id}" formatCode="${esc(f.code)}"/>`).join('')}</numFmts>` : '') +
    `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills>` +
    `<borders count="${borders.length}">${borders.join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    (dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>` : '') + '</styleSheet>';
  return { xml, xfOf };
}

// Дописать элементы в конец списка <tag count="n">…</tag> исходного XML (count увеличивается)
function appendToList(xml, tag, items, insertAfterRoot = null) {
  if (!items.length) return xml;
  const m = new RegExp(`<${tag}\\b([^>]*?)(/?)>`).exec(xml);
  if (!m) {
    if (!insertAfterRoot) return xml;
    const root = new RegExp(`<${insertAfterRoot}\\b[^>]*>`).exec(xml);
    const at = root.index + root[0].length;
    return `${xml.slice(0, at)}<${tag} count="${items.length}">${items.join('')}</${tag}>${xml.slice(at)}`;
  }
  const n = +((/\bcount="(\d+)"/.exec(m[1]) || [])[1] || 0) + items.length;
  const attrs = /\bcount="/.test(m[1]) ? m[1].replace(/\bcount="\d+"/, `count="${n}"`) : `${m[1]} count="${n}"`;
  if (m[2]) return `${xml.slice(0, m.index)}<${tag}${attrs}>${items.join('')}</${tag}>${xml.slice(m.index + m[0].length)}`;
  const close = xml.indexOf(`</${tag}>`, m.index);
  return `${xml.slice(0, m.index)}<${tag}${attrs}>${xml.slice(m.index + m[0].length, close)}${items.join('')}${xml.slice(close)}`;
}

// ---------- Разметка листа ----------
// Порядок дочерних элементов <worksheet> по схеме — для вставки отсутствующих
const SHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr',
  'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks',
  'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF',
  'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
// Первый элемент <name> верхнего уровня: { start, end } (end — после закрывающего тега)
function findEl(xml, name, from = 0) {
  const re = new RegExp(`<${name}(?=[\\s/>])`, 'g');
  re.lastIndex = from;
  const m = re.exec(xml);
  if (!m) return null;
  const gt = xml.indexOf('>', m.index);
  if (xml[gt - 1] === '/') return { start: m.index, end: gt + 1 };
  const close = xml.indexOf(`</${name}>`, gt);
  return { start: m.index, end: close + name.length + 3 };
}
// Заменить элемент листа (content '' — удалить); отсутствующий — вставить на место по схеме
export function setSheetEl(xml, name, content) {
  const at = findEl(xml, name);
  if (at) return xml.slice(0, at.start) + content + xml.slice(at.end);
  if (!content) return xml;
  for (const next of SHEET_ORDER.slice(SHEET_ORDER.indexOf(name) + 1)) {
    const n = findEl(xml, next);
    if (n) return xml.slice(0, n.start) + content + xml.slice(n.start);
  }
  const end = xml.lastIndexOf('</worksheet>');
  return xml.slice(0, end) + content + xml.slice(end);
}

// XML ячейки
// fx — { ref, cm }: формула пишется как формула массива (t="array" ref=…; cm — признак динамического массива)
function cellXml(cell, ref, xfOf, str, fx = null) {
  const s = xfOf(cell.s);
  const sAttr = (s ? ` s="${s}"` : '') + (fx && fx.cm ? ` cm="${fx.cm}"` : '');
  const v = cell.v;
  if (cell.f != null) {
    const f = fx ? `<f t="array" ref="${fx.ref}">${esc(fileFormula(cell.f))}</f>` : `<f>${esc(fileFormula(cell.f))}</f>`;
    if (isErr(v)) return `<c r="${ref}"${sAttr} t="e">${f}<v>${esc(v.error)}</v></c>`;
    if (typeof v === 'string') return `<c r="${ref}"${sAttr} t="str">${f}<v>${esc(v)}</v></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
    return `<c r="${ref}"${sAttr}>${f}${typeof v === 'number' && isFinite(v) ? `<v>${v}</v>` : ''}</c>`;
  }
  if (v == null) return `<c r="${ref}"${sAttr}/>`;
  if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
  if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
  if (isErr(v)) return `<c r="${ref}"${sAttr} t="e"><v>${esc(v.error)}</v></c>`;
  return `<c r="${ref}"${sAttr} t="s"><v>${str(String(v))}</v></c>`;
}
const lineAttrs = (info, withXf, styleAttr) => (info
  ? `${withXf && info.xf ? (styleAttr === 's' ? ` s="${info.xf}" customFormat="1"` : ` style="${info.xf}"`) : ''}${info.hidden ? ' hidden="1"' : ''}${info.level ? ` outlineLevel="${info.level}"` : ''}${info.collapsed ? ' collapsed="1"' : ''}${info.bestFit ? ' bestFit="1"' : ''}`
  : '');
// Как писать формулу ячейки: динамический массив (якорь «разлива» или формула с поэлементными операциями) —
// t="array" + cm; формула массива Ctrl+Shift+Enter — t="array" без cm; иначе — обычная <f>
function formulaFx(sheet, k, cell, cm) {
  if (cell.f == null) return null;
  const [r, c] = unkey(k);
  if (cell.cse) return { ref: cell.cse };
  if (cell.legacy) return null;
  const rg = sheet.spillAt && sheet.spillAt.get(k);
  if (rg) return { ref: rangeName(rg), cm };
  if (cell.arr || cell.ax) return { ref: addr(r, c), cm };
  return null;
}
// <sheetData>: cellOf(r, c, cell) — XML ячейки; значения «разлива» дописываются обычными значениями (как пишет Excel)
function sheetDataXml(sheet, cellOf, withXf, xfOf, str) {
  const rows = new Map(); // r → [[c, xml]]
  const add = (r, c, xml) => { if (!rows.has(r)) rows.set(r, []); rows.get(r).push([c, xml]); };
  for (const [k, cell] of sheet.cells) {
    const [r, c] = unkey(k);
    const ghost = sheet.spill && cell.f == null && cell.v == null ? sheet.spill.get(k) : undefined;
    add(r, c, ghost !== undefined ? cellXml({ ...cell, v: ghost }, addr(r, c), xfOf, str) : cellOf(r, c, cell));
  }
  if (sheet.spill) {
    for (const [k, v] of sheet.spill) {
      if (sheet.cells.has(k)) continue;
      const [r, c] = unkey(k);
      add(r, c, cellXml({ v }, addr(r, c), xfOf, str));
    }
  }
  for (const r of Object.keys(sheet.rowH)) if (!rows.has(+r)) rows.set(+r, []);
  for (const r of Object.keys(sheet.rowInfo || {})) if (!rows.has(+r)) rows.set(+r, []);
  const data = [...rows.keys()].sort((a, b) => a - b).map((r) => {
    const cells = rows.get(r).sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('');
    const h = sheet.rowH[r];
    return `<row r="${r + 1}"${lineAttrs((sheet.rowInfo || {})[r], withXf, 's')}${h ? ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"` : ''}>${cells}</row>`;
  }).join('');
  return `<sheetData>${data}</sheetData>`;
}
// <cols>: соседние столбцы с одинаковыми свойствами — одним диапазоном
function colsXml(sheet, withXf) {
  const info = sheet.colInfo || {};
  const idx = [...new Set([...Object.keys(sheet.colW), ...Object.keys(info)].map(Number))].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < idx.length;) {
    const c = idx[i];
    const sig = (x) => `${sheet.colW[x] ?? ''}|${lineAttrs(info[x], withXf, 'style')}`;
    let j = i;
    while (j + 1 < idx.length && idx[j + 1] === idx[j] + 1 && sig(idx[j + 1]) === sig(c)) j++;
    const w = sheet.colW[c];
    // без своей ширины: скрытый — 0 (как пишет Excel), иначе ширина по умолчанию Пера
    const width = w != null ? pxToWidth(w) : info[c] && info[c].hidden ? 0 : pxToWidth(DEFAULT_COL_W);
    out.push(`<col min="${c + 1}" max="${idx[j] + 1}" width="${width}"${lineAttrs(info[c], withXf, 'style')} customWidth="1"/>`);
    i = j + 1;
  }
  return out.length ? `<cols>${out.join('')}</cols>` : '';
}
const mergesXml = (sheet) => (sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${addr(m.r1, m.c1)}:${addr(m.r2, m.c2)}"/>`).join('')}</mergeCells>` : '');
function dimensionXml(sheet) {
  let r1 = Infinity; let c1 = Infinity; let r2 = 0; let c2 = 0;
  for (const k of sheet.cells.keys()) { const [r, c] = unkey(k); r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c); }
  if (r1 === Infinity) return '<dimension ref="A1"/>';
  return `<dimension ref="${rangeName({ r1, c1, r2, c2 })}"/>`;
}

// Уровни структуры в <sheetFormatPr> (по ним Excel рисует поле с кнопками групп)
function formatPrXml(tag, sheet) {
  let t = tag.replace(/\s+outlineLevel(Row|Col)="[^"]*"/g, '');
  const lr = maxLevel(sheet, 'row'); const lc = maxLevel(sheet, 'col');
  const attrs = `${lr ? ` outlineLevelRow="${lr}"` : ''}${lc ? ` outlineLevelCol="${lc}"` : ''}`;
  return attrs ? t.replace(/\s*(\/?)>$/, `${attrs}$1>`) : t;
}
// Закрепление в первом <sheetView> исходного листа: прежние <pane>/<selection> заменяются
function setFreezeXml(xml, f) {
  const sv = findEl(xml, 'sheetView');
  if (!sv) return f ? setSheetEl(xml, 'sheetViews', `<sheetViews><sheetView workbookViewId="0">${paneXml(f)}</sheetView></sheetViews>`) : xml;
  let el = xml.slice(sv.start, sv.end);
  if (!el.includes('</sheetView>')) el = el.replace(/\s*\/>$/, '></sheetView>');
  el = el.replace(/<pane\b[^>]*\/>/g, '').replace(/<selection\b[^>]*\/>/g, '');
  el = el.replace(/^<sheetView\b[^>]*>/, (open) => open + paneXml(f));
  return xml.slice(0, sv.start) + el + xml.slice(sv.end);
}

// Новый лист целиком
function newSheetXml(sheet, active, xfOf, str, cm = '1') {
  const cols = colsXml(sheet, false);
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS}" xmlns:r="${R}">` +
    `${dimensionXml(sheet)}<sheetViews><sheetView workbookViewId="0"${active ? ' tabSelected="1"' : ''}${sheet.freeze ? `>${paneXml(sheet.freeze)}</sheetView>` : '/>'}</sheetViews>` +
    `${formatPrXml('<sheetFormatPr defaultRowHeight="15"/>', sheet)}${cols}${sheetDataXml(sheet, (r, c, cell) => cellXml(cell, addr(r, c), xfOf, str, formulaFx(sheet, key(r, c), cell, cm)), false, xfOf, str)}${protectionXml(sheet.protection)}${autoFilterXml(sheet.autoFilter)}${mergesXml(sheet)}${cfXml(sheet.cf, xfOf.dxf, rangeName)}${dvXml(sheet.dv, rangeName)}` +
    `${marginsXml((sheet.page && sheet.page.margins) || {})}${breaksXml(sheet.breaks, 'rowBreaks', 16383)}${breaksXml(sheet.breaks, 'colBreaks', 1048575)}</worksheet>`;
  return sheet.page ? setPageXml(xml, sheet.page) : xml;
}

// Сдвиг диапазонов sqref/ref после вставки/удаления строк и столбцов (удалённые целиком — пропадают)
function shiftSqref(sqref, shifts) {
  const out = [];
  for (const part of sqref.split(/\s+/).filter(Boolean)) {
    let rg = parseRange(part);
    if (!rg) { out.push(part); continue; }
    for (const s of shifts) { rg = shiftRange(rg, s); if (!rg) break; }
    if (rg) out.push(rangeName(rg));
  }
  return out.join(' ');
}
function shiftRange(rg, { axis, at, count }) {
  const [a, b] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
  const out = { ...rg };
  if (count > 0) {
    if (rg[a] >= at) out[a] += count;
    if (rg[b] >= at) out[b] += count;
    return out;
  }
  const n = -count; const end = at + n - 1;
  const map = (v) => (v > end ? v - n : v >= at ? at : v);
  if (rg[a] >= at && rg[b] <= end) return null;
  out[a] = map(rg[a]);
  out[b] = rg[b] > end ? rg[b] - n : rg[b] >= at ? at - 1 : rg[b];
  return out[b] < out[a] ? null : out;
}

// Выбранный лист — только активный (иначе Excel откроет несколько листов «группой»)
function selectTab(xml, active) {
  const m = /<sheetView\b[^>]*?(?=\/?>)/.exec(xml);
  if (!m) return xml;
  const sel = /\stabSelected="(1|true)"/.test(m[0]);
  if (sel === active) return xml;
  const t = m[0].replace(/\s+tabSelected="[^"]*"/, '');
  return xml.slice(0, m.index) + (active ? `${t} tabSelected="1"` : t) + xml.slice(m.index + m[0].length);
}

// Лист поверх исходного XML: заменяются данные, столбцы, объединения, размер, выбор листа; остальное — как в файле
function patchSheetXml(xml, sheet, active, root, book, xfOf, str, cm) {
  // исходный XML ячеек по адресу (неизменённые ячейки из файла пишутся им)
  const rawList = [];
  const sd = findEl(xml, 'sheetData');
  if (sd) {
    const re = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
    const part = xml.slice(sd.start, sd.end);
    let m;
    while ((m = re.exec(part))) rawList.push(m[0]);
  }
  const byRef = new Map();
  for (const raw of rawList) { const m = /\br="([A-Z]+\d+)"/.exec(raw.slice(0, raw.indexOf('>'))); if (m) byRef.set(m[1], raw); }
  const orig = sheet.src && sheet.src.cells && sheet.src.noRaw ? sheet.src.cells : null;
  const cellOf = (r, c, cell) => {
    const fx = formulaFx(sheet, key(r, c), cell, cm);
    // неизменённая ячейка из файла на своём месте — исходной разметкой (форматированный текст и т. п.)
    const k = key(r, c);
    if (!fx && orig && orig.get(k) === cell && !sheet.src.noRaw.has(k)) { const raw = byRef.get(addr(r, c)); if (raw) return raw; }
    return cellXml(cell, addr(r, c), xfOf, str, fx);
  };
  let out = xml;
  out = setSheetEl(out, 'sheetData', sheetDataXml(sheet, cellOf, true, xfOf, str));
  out = setSheetEl(out, 'cols', colsXml(sheet, true));
  out = setSheetEl(out, 'mergeCells', mergesXml(sheet));
  out = setSheetEl(out, 'dimension', dimensionXml(sheet));
  out = selectTab(out, active);
  if (JSON.stringify(readFreeze(root)) !== JSON.stringify(sheet.freeze || null)) out = setFreezeXml(out, sheet.freeze || null);
  if (JSON.stringify(readPage(root)) !== JSON.stringify(sheet.page || null)) out = setPageXml(out, sheet.page || {});
  if (JSON.stringify(readProtection(root)) !== JSON.stringify(sheet.protection || null)) out = setSheetEl(out, 'sheetProtection', protectionXml(sheet.protection));
  const afChanged = JSON.stringify(readAutoFilter(root)) !== JSON.stringify(sheet.autoFilter || null);
  const cfChanged = cfSignature(readSheetCf(root, book)) !== cfSignature(sheet.cf);
  const dvChanged = dvSignature(readSheetDv(root)) !== dvSignature(sheet.dv);
  const brChanged = JSON.stringify(readBreaks(root)) !== JSON.stringify(sheet.breaks || null);
  const fp = findEl(out, 'sheetFormatPr');
  if (fp) out = out.slice(0, fp.start) + formatPrXml(out.slice(fp.start, fp.end), sheet) + out.slice(fp.end);
  else if (maxLevel(sheet, 'row') || maxLevel(sheet, 'col')) out = setSheetEl(out, 'sheetFormatPr', formatPrXml('<sheetFormatPr defaultRowHeight="15"/>', sheet));
  // вставка/удаление строк и столбцов: диапазоны условного форматирования, проверки, ссылок, фильтра
  const shifts = (sheet.src && sheet.src.shifts) || [];
  if (shifts.length) {
    out = out.replace(/<(conditionalFormatting|dataValidation|hyperlink|autoFilter)\b[^>]*>/g, (tag) => tag
      .replace(/\b(sqref|ref)="([^"]*)"/, (_, a, v) => `${a}="${shiftSqref(v, shifts)}"`));
  }
  // автофильтр изменён в Пере (или сдвинут) — заново из модели (sortState Excel при этом не сохраняется)
  if (afChanged || shifts.length) out = setSheetEl(out, 'autoFilter', autoFilterXml(sheet.autoFilter));
  // условное форматирование изменено (или сдвинуто) — все блоки заново из модели; правила, которые Перо не меняло,
  // пишутся исходной разметкой
  if (cfChanged || shifts.length) {
    out = out.replace(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, '');
    out = setSheetEl(out, 'conditionalFormatting', cfXml(sheet.cf, xfOf.dxf, rangeName));
  }
  // проверка данных изменена (или сдвинута) — заново из модели; неизменённые правила — исходной разметкой
  if (dvChanged || shifts.length) out = setSheetEl(out, 'dataValidations', dvXml(sheet.dv, rangeName));
  if (brChanged || shifts.length) {
    out = setSheetEl(out, 'rowBreaks', breaksXml(sheet.breaks, 'rowBreaks', 16383));
    out = setSheetEl(out, 'colBreaks', breaksXml(sheet.breaks, 'colBreaks', 1048575));
  }
  return out;
}

// Номер строки/столбца после сдвигов (внутри удалённых — к началу удаления)
function shiftIndex(v, shifts, axis) {
  for (const s of shifts) {
    if (s.axis !== axis) continue;
    if (s.count > 0) { if (v >= s.at) v += s.count; } else if (v >= s.at - s.count) v += s.count; else if (v >= s.at) v = s.at;
  }
  return v;
}
// Части, привязанные к ячейкам листа (рисунки, примечания, «умные» таблицы), — после вставки/удаления строк и столбцов
async function shiftSheetParts(zip, sheetPart, shifts) {
  const relsPart = `${sheetPart.slice(0, sheetPart.lastIndexOf('/'))}/_rels/${sheetPart.slice(sheetPart.lastIndexOf('/') + 1)}.rels`;
  const rels = zip.file(relsPart) ? await zip.file(relsPart).async('string') : '';
  const dir = sheetPart.slice(0, sheetPart.lastIndexOf('/'));
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const type = (/\bType="([^"]*)"/.exec(m[0]) || [])[1] || '';
    const target = (/\bTarget="([^"]*)"/.exec(m[0]) || [])[1];
    if (!target || /TargetMode="External"/.test(m[0])) continue;
    const part = resolve(dir, target);
    const f = zip.file(part);
    if (!f) continue;
    let t = await f.async('string');
    const u0 = t;
    if (/\/drawing$/.test(type)) {
      t = t.replace(/<xdr:(from|to)>([\s\S]*?)<\/xdr:\1>/g, (all, tag, body) => `<xdr:${tag}>${body
        .replace(/<xdr:row>(\d+)<\/xdr:row>/, (_, v) => `<xdr:row>${shiftIndex(+v, shifts, 'row')}</xdr:row>`)
        .replace(/<xdr:col>(\d+)<\/xdr:col>/, (_, v) => `<xdr:col>${shiftIndex(+v, shifts, 'col')}</xdr:col>`)}</xdr:${tag}>`);
    } else if (/\/comments$/.test(type) || /\/table$/.test(type)) {
      t = t.replace(/<(comment|table|autoFilter)\b[^>]*>/g, (tag) => tag.replace(/\bref="([^"]*)"/, (_, v) => `ref="${shiftSqref(v, shifts) || v}"`));
    } else if (/\/vmlDrawing$/.test(type)) {
      t = t.replace(/<x:Row>(\d+)<\/x:Row>/g, (_, v) => `<x:Row>${shiftIndex(+v, shifts, 'row')}</x:Row>`)
        .replace(/<x:Column>(\d+)<\/x:Column>/g, (_, v) => `<x:Column>${shiftIndex(+v, shifts, 'col')}</x:Column>`)
        .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/g, (_, v) => {
          const a = v.split(',').map((x) => +x.trim());
          if (a.length !== 8 || a.some(isNaN)) return `<x:Anchor>${v}</x:Anchor>`;
          a[0] = shiftIndex(a[0], shifts, 'col'); a[4] = shiftIndex(a[4], shifts, 'col');
          a[2] = shiftIndex(a[2], shifts, 'row'); a[6] = shiftIndex(a[6], shifts, 'row');
          return `<x:Anchor>\n    ${a.join(', ')}</x:Anchor>`;
        });
    }
    if (t !== u0) zip.file(part, t);
  }
}
// Ссылки на лист в формулах файла (имена, ряды диаграмм) — после вставки/удаления строк и столбцов
function shiftRefs(text, name, shifts) {
  let x = text;
  try { for (const sh of shifts) x = shiftFormula(x, { sheet: name, current: null, ...sh }); } catch { return text; }
  return x;
}

// Часть metadata.xml, как пишет Excel: тип XLDAPR — «динамический массив»; ячейки ссылаются на него атрибутом cm="1"
const METADATA_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
const METADATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const METADATA_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const needsDynamic = (wb) => wb.sheets.some((sh) => [...sh.cells.values()].some((c) => c.f != null && !c.legacy && !c.cse && (c.arr || c.ax)));
// Номер cm динамического массива в существующем metadata.xml: блок cellMetadata, ссылающийся на тип XLDAPR
function dynamicCm(metaXml) {
  const types = [...metaXml.matchAll(/<metadataType\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1]);
  const t = types.indexOf('XLDAPR') + 1;
  if (!t) return null;
  const cell = /<cellMetadata\b[^>]*>([\s\S]*?)<\/cellMetadata>/.exec(metaXml);
  if (!cell) return null;
  const bks = [...cell[1].matchAll(/<bk>([\s\S]*?)<\/bk>/g)];
  const i = bks.findIndex((b) => new RegExp(`<rc\\b[^>]*\\bt="${t}"`).test(b[1]));
  return i >= 0 ? String(i + 1) : null;
}

// Переименование листа в формулах файла (именованные диапазоны, ряды диаграмм): 'Старое'!A1 → 'Новое'!A1
const quoteSheet = (name) => (/^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
function renameRefs(text, old, name) {
  const e = old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const quoted = `'${e.replace(/'/g, "''")}'`;
  return text.replace(new RegExp(`(^|[^\\wА-Яа-яЁё'.])(${quoted}|${e})!`, 'g'), (_, pre) => `${pre}${name == null ? '' : quoteSheet(name) + '!'}`);
}
const xmlEsc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const xmlUnesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// Книга поверх исходного пакета: переписываются листы, стили (дописываются), общие строки (дописываются), список листов;
// диаграммы, сводные, условное форматирование, проверка, примечания, таблицы, тема, настройки — байт в байт
async function exportWithSource(wb, source) {
  const zip = await JSZip.loadAsync(source);
  const book = await readBook(zip);
  const { text } = book;
  const CT = '[Content_Types].xml';
  let ct = await text(CT);
  let relsXml = await text(book.relsPath);
  let wbXml = await text(book.wbPath);
  const relTarget = (p) => (p.startsWith(`${book.dir}/`) ? p.slice(book.dir.length + 1) : `/${p}`);

  // стили
  const stylesRoot = book.stylesDoc && book.stylesDoc.documentElement;
  const count = (tag, item) => kids(kid(stylesRoot, tag), item).length;
  const fmts = {};
  for (const f of kids(kid(stylesRoot, 'numFmts'), 'numFmt')) fmts[f.getAttribute('formatCode')] = +f.getAttribute('numFmtId');
  const base = stylesRoot
    ? { styles: book.styles, dxfs: book.dxfs, fmts, counts: { fonts: count('fonts', 'font'), fills: count('fills', 'fill'), borders: count('borders', 'border'), xfs: count('cellXfs', 'xf'), dxfs: count('dxfs', 'dxf') } }
    : null;
  const st = buildStyles(wb, base);

  // общие строки: исходные остаются на местах, новые — в конец
  const sstIndex = new Map();
  book.sst.forEach((s, i) => { if (book.sstPlain[i] && !sstIndex.has(s)) sstIndex.set(s, i); });
  const newStrings = [];
  const str = (s) => {
    if (!sstIndex.has(s)) { sstIndex.set(s, book.sst.length + newStrings.length); newStrings.push(s); }
    return sstIndex.get(s);
  };

  // листы
  const removed = book.sheets.filter((s) => s.worksheet && !wb.sheets.some((x) => x.src && x.src.rid === s.rid));
  const sheetEls = {};
  for (const m of wbXml.matchAll(/<sheet\b[^>]*?\/>/g)) { const id = /\br:id="([^"]+)"/.exec(m[0]); if (id) sheetEls[id[1]] = m[0]; }
  let maxSheetId = Math.max(0, ...[...wbXml.matchAll(/<sheet\b[^>]*?\bsheetId="(\d+)"/g)].map((m) => +m[1]));
  const newRels = [];
  const newOverrides = [];
  const sheetTags = [];
  let peroN = 0;
  const renames = []; // [старое, новое|null]
  const shifted = []; // [имя листа, сдвиги]
  // признак динамического массива: из metadata.xml файла; нет части — добавим свою
  let cm = null;
  const metaRel = Object.values(book.rels).find((r) => r.type === METADATA_REL);
  if (metaRel) cm = dynamicCm((await text(metaRel.target)) || '');
  else if (needsDynamic(wb)) {
    cm = '1';
    zip.file(`${book.dir}/metadata.xml`, METADATA_XML);
    newRels.push(`<Relationship Id="rIdPeroMeta" Type="${METADATA_REL}" Target="metadata.xml"/>`);
    newOverrides.push(`<Override PartName="/${book.dir}/metadata.xml" ContentType="${METADATA_CT}"/>`);
  }
  for (let i = 0; i < wb.sheets.length; i++) {
    const sheet = wb.sheets[i];
    const s = sheet.src && book.sheets.find((x) => x.rid === sheet.src.rid && x.worksheet);
    if (s && sheetEls[s.rid]) {
      const xml = await text(s.target);
      const isSheet = xml != null && /^[\s\S]*?<worksheet\b/.test(xml);
      if (isSheet) {
        // неизменённый лист — без разбора: только отметка выбранного листа
        zip.file(s.target, pristine.has(sheet)
          ? selectTab(xml, i === wb.active)
          : patchSheetXml(xml, sheet, i === wb.active, parseSheet(xml, { data: false }).documentElement, book, st.xfOf, str, cm));
        sheetTags.push(sheetEls[s.rid].replace(/\bname="[^"]*"/, `name="${esc(sheet.name)}"`));
        const shifts = (!pristine.has(sheet) && sheet.src.shifts) || [];
        if (shifts.length) { shifted.push([sheet.name, shifts, s.name]); await shiftSheetParts(zip, s.target, shifts); }
        if (sheet.name !== s.name) renames.push([s.name, sheet.name]);
        continue;
      }
    }
    // новый лист (или лист в непривычной разметке — пишется заново)
    let part;
    do { peroN++; part = `${book.dir}/worksheets/sheet_pero${peroN}.xml`; } while (zip.file(part));
    const rid = `rIdPeroS${peroN}`;
    zip.file(part, newSheetXml(sheet, i === wb.active, st.xfOf, str, cm));
    newRels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${relTarget(part)}"/>`);
    newOverrides.push(`<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    sheetTags.push(`<sheet name="${esc(sheet.name)}" sheetId="${++maxSheetId}" r:id="${rid}"/>`);
  }
  // листы диаграмм (в модели их нет) остаются в книге после рабочих листов
  for (const s of book.sheets) if (!s.worksheet && sheetEls[s.rid]) sheetTags.push(sheetEls[s.rid]);
  for (const s of removed) {
    renames.push([s.name, null]);
    zip.remove(s.target);
    const relsPart = `${s.target.slice(0, s.target.lastIndexOf('/'))}/_rels/${s.target.slice(s.target.lastIndexOf('/') + 1)}.rels`;
    if (zip.file(relsPart)) zip.remove(relsPart);
    relsXml = relsXml.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${s.rid}"[^>]*/>`), '');
    ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${s.target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
  }
  wbXml = wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${sheetTags.join('')}</sheets>`);

  // именованные диапазоны: номера листов (localSheetId) после удаления, имена листов после переименования
  const oldIdx = book.sheets.map((s) => s.rid);
  const newIdx = (old) => {
    const rid = oldIdx[old];
    const i = wb.sheets.findIndex((x) => x.src && x.src.rid === rid);
    if (i >= 0) return i;
    const nonWs = book.sheets.filter((s) => !s.worksheet).findIndex((s) => s.rid === rid);
    return nonWs >= 0 ? wb.sheets.length + nonWs : -1;
  };
  const fixFormula = (t) => { let x = xmlUnesc(t); for (const [o, nn] of renames) x = nn == null ? x.replace(new RegExp(`(?:'${o.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/'/g, "''")}'|${o.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})![$A-Z0-9:]+`, 'g'), '#REF!') : renameRefs(x, o, nn); return xmlEsc(x); };
  wbXml = wbXml.replace(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g, (all, attrs, body) => {
    const m = /\blocalSheetId="(\d+)"/.exec(attrs);
    let a = attrs;
    if (m) {
      const i = newIdx(+m[1]);
      if (i < 0) return '';
      a = attrs.replace(/\blocalSheetId="\d+"/, `localSheetId="${i}"`);
    }
    let b = renames.length ? fixFormula(body) : body;
    for (const [nm, sh] of shifted) b = xmlEsc(shiftRefs(xmlUnesc(b), nm, sh));
    return `<definedName${a}>${b}</definedName>`;
  }).replace(/<definedNames>\s*<\/definedNames>/, '');
  // область печати и сквозные строки: у изменённых в Пере листов — заново
  const printNames = [];
  wb.sheets.forEach((sheet, i) => {
    const orig = sheet.src ? book.sheets.findIndex((x) => x.rid === sheet.src.rid) : -1;
    const was = orig >= 0 ? readPrintNames(book.wbDoc, orig) : { printArea: null, printTitles: null };
    const area = sheet.printArea || null;
    const titles = sheet.printTitles && (sheet.printTitles.rows || sheet.printTitles.cols) ? sheet.printTitles : null;
    if (JSON.stringify(was.printArea) !== JSON.stringify(area) || orig < 0) {
      wbXml = wbXml.replace(new RegExp(`<definedName name="_xlnm\\.Print_Area" localSheetId="${i}"[^>]*>[\\s\\S]*?</definedName>`), '');
      if (area) printNames.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${xmlEsc(printAreaText(sheet.name, area))}</definedName>`);
    }
    // служебное имя диапазона фильтра (Excel создаёт его сам; пишем, чтобы файл был как у Excel)
    {
      wbXml = wbXml.replace(new RegExp(`<definedName name="_xlnm\\._FilterDatabase" localSheetId="${i}"[^>]*>[\\s\\S]*?</definedName>`), '');
      if (sheet.autoFilter) printNames.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${xmlEsc(printAreaText(sheet.name, sheet.autoFilter.ref))}</definedName>`);
    }
    if (JSON.stringify(was.printTitles) !== JSON.stringify(titles) || orig < 0) {
      wbXml = wbXml.replace(new RegExp(`<definedName name="_xlnm\\.Print_Titles" localSheetId="${i}"[^>]*>[\\s\\S]*?</definedName>`), '');
      if (titles) printNames.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${xmlEsc(printTitlesText(sheet.name, titles))}</definedName>`);
    }
  });
  if (printNames.length) {
    if (/<definedNames>/.test(wbXml)) wbXml = wbXml.replace('<definedNames>', `<definedNames>${printNames.join('')}`);
    else {
      const after = /<\/externalReferences>/.test(wbXml) ? '</externalReferences>' : '</sheets>';
      wbXml = wbXml.replace(after, `${after}<definedNames>${printNames.join('')}</definedNames>`);
    }
  }
  if (wb.names && !pristine.has(wb.names)) {
    wbXml = wbXml.replace(/<definedName\b(?![^>]*\bname="_xlnm\.)[^>]*>[\s\S]*?<\/definedName>/g, '');
    const user = namesXml(wb);
    if (user) {
      if (/<definedNames>/.test(wbXml)) wbXml = wbXml.replace('</definedNames>', `${user}</definedNames>`);
      else {
        const after = /<\/externalReferences>/.test(wbXml) ? '</externalReferences>' : '</sheets>';
        wbXml = wbXml.replace(after, `${after}<definedNames>${user}</definedNames>`);
      }
    }
  }
  wbXml = wbXml.replace(/<definedNames>\s*<\/definedNames>/, '');
  // активный лист; пересчёт при открытии (кэш формул мог устареть)
  wbXml = wbXml.replace(/<workbookView\b[^>]*?(?=\/?>)/, (tag) => {
    const t = tag.replace(/\s+(activeTab|firstSheet)="[^"]*"/g, '');
    return wb.active ? `${t} activeTab="${wb.active}"` : t;
  });
  // режим вычислений: calcMode="manual" — как в книге Перо (по умолчанию «авто» — атрибута нет)
  const calcAttr = wb.calcMode === 'manual' ? ' calcMode="manual"' : '';
  if (/<calcPr\b/.test(wbXml)) wbXml = wbXml.replace(/<calcPr\b([^>]*?)(\/?)>/, (_, a, sc) => `<calcPr${a.replace(/\s+(fullCalcOnLoad|calcMode)="[^"]*"/g, '')}${calcAttr} fullCalcOnLoad="1"${sc}>`);
  else wbXml = wbXml.replace(/(<\/definedNames>|<\/sheets>)(?![\s\S]*(<\/definedNames>|<\/sheets>))/, `$1<calcPr${calcAttr} fullCalcOnLoad="1"/>`);
  // ряды диаграмм и источники сводных — по новым именам листов
  if (renames.some(([, nn]) => nn != null) || shifted.length) {
    for (const p of Object.keys(zip.files)) {
      if (/^xl\/charts\/chart[^/]*\.xml$/.test(p)) {
        const t = await zip.file(p).async('string');
        const u = t.replace(/<c:f>([\s\S]*?)<\/c:f>/g, (_, f) => {
          let x = xmlUnesc(f);
          for (const [o, nn] of renames) if (nn != null) x = renameRefs(x, o, nn);
          for (const [nm, sh] of shifted) x = shiftRefs(x, nm, sh);
          return `<c:f>${xmlEsc(x)}</c:f>`;
        });
        if (u !== t) zip.file(p, u);
      } else if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(p)) {
        const t = await zip.file(p).async('string');
        const u = t.replace(/<worksheetSource\b[^>]*>/g, (tag) => {
          const sheetName = xmlUnesc((/\bsheet="([^"]*)"/.exec(tag) || [])[1] || '');
          const ren = renames.find(([o, nn]) => o === sheetName && nn != null);
          let out = ren ? tag.replace(/\bsheet="[^"]*"/, `sheet="${esc(ren[1])}"`) : tag;
          const sh = shifted.find(([, , orig]) => orig === sheetName);
          if (sh) out = out.replace(/\bref="([^"]*)"/, (_, v) => `ref="${shiftSqref(v, sh[1]) || v}"`);
          return out;
        });
        if (u !== t) zip.file(p, u);
      }
    }
  }

  // цепочка вычислений устарела — Excel построит заново
  const calc = Object.entries(book.rels).find(([, r]) => /\/calcChain$/.test(r.type));
  if (calc) {
    zip.remove(calc[1].target);
    relsXml = relsXml.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${calc[0]}"[^>]*/>`), '');
    ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${calc[1].target}"[^>]*/>`), '');
  }

  // стили — дописать новые
  if (base) {
    let sx = await text(book.stylesPath);
    const a = st.added;
    sx = appendToList(sx, 'numFmts', a.numFmts, 'styleSheet');
    sx = appendToList(sx, 'fonts', a.fonts);
    sx = appendToList(sx, 'fills', a.fills);
    sx = appendToList(sx, 'borders', a.borders);
    sx = appendToList(sx, 'cellXfs', a.cellXfs);
    if (a.dxfs.length) {
      // dxfs стоят после cellStyles (перед tableStyles/colors/extLst); нет — вставить туда
      if (/<dxfs\b/.test(sx)) sx = appendToList(sx, 'dxfs', a.dxfs);
      else {
        const m = /<(tableStyles|colors|extLst)\b|<\/styleSheet>/.exec(sx);
        sx = `${sx.slice(0, m.index)}<dxfs count="${a.dxfs.length}">${a.dxfs.join('')}</dxfs>${sx.slice(m.index)}`;
      }
    }
    zip.file(book.stylesPath, sx);
  } else {
    // в книге без стилей номера xf считались по новой таблице — она и пишется
    zip.file(`${book.dir}/styles.xml`, st.xml);
    newRels.push(`<Relationship Id="rIdPeroSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
    newOverrides.push(`<Override PartName="/${book.dir}/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`);
  }
  // общие строки — дописать новые
  if (newStrings.length) {
    const items = newStrings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`);
    if (book.sstPath) {
      let sx = await text(book.sstPath);
      const total = book.sst.length + newStrings.length;
      sx = sx.replace(/<sst\b[^>]*>/, (tag) => tag.replace(/\s+(count|uniqueCount)="\d+"/g, '').replace(/<sst\b/, `<sst count="${total}" uniqueCount="${total}"`));
      sx = sx.includes('</sst>') ? sx.replace('</sst>', `${items.join('')}</sst>`) : sx.replace(/<sst\b([^>]*?)\/>/, `<sst$1>${items.join('')}</sst>`);
      zip.file(book.sstPath, sx);
    } else {
      zip.file(`${book.dir}/sharedStrings.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS}" count="${items.length}" uniqueCount="${items.length}">${items.join('')}</sst>`);
      newRels.push('<Relationship Id="rIdPeroSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>');
      newOverrides.push(`<Override PartName="/${book.dir}/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`);
    }
  }

  relsXml = relsXml.replace('</Relationships>', `${newRels.join('')}</Relationships>`);
  ct = ct.replace('</Types>', `${newOverrides.join('')}</Types>`);
  zip.file(book.relsPath, relsXml);
  zip.file(book.wbPath, wbXml);
  zip.file(CT, ct);
  // дата изменения
  const core = await text('docProps/core.xml');
  if (core) {
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', core.replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${now}$2`));
  }
  await precompress(JSZip, zip); // большие части — встроенным сжатием (zipBytes.js)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function freshPrintNames(wb) {
  const out = [];
  wb.sheets.forEach((s, i) => {
    if (s.printArea) out.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${xmlEsc(printAreaText(s.name, s.printArea))}</definedName>`);
    if (s.printTitles && (s.printTitles.rows || s.printTitles.cols)) out.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${xmlEsc(printTitlesText(s.name, s.printTitles))}</definedName>`);
  });
  const all = out.join('') + namesXml(wb);
  return all ? `<definedNames>${all}</definedNames>` : '';
}

/** source — байты исходного .xlsx: всё, что Перо не редактирует, переносится из него */
export async function exportXlsx(wb, { source = null } = {}) {
  if (source) return exportWithSource(wb, source);
  const { xml: stylesXml, xfOf } = buildStyles(wb);
  const sst = [];
  const sstIndex = new Map();
  const str = (s) => { if (!sstIndex.has(s)) { sstIndex.set(s, sst.length); sst.push(s); } return sstIndex.get(s); };
  const dynamic = needsDynamic(wb);
  const sheetXml = wb.sheets.map((sheet, si) => newSheetXml(sheet, si === wb.active, xfOf, str, '1'));

  const zip = new JSZip();
  const sheetsCt = wb.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheetsCt}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>${dynamic ? `<Override PartName="/xl/metadata.xml" ContentType="${METADATA_CT}"/>` : ''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Pero</Application></Properties>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="${R}"><bookViews><workbookView activeTab="${wb.active || 0}"/></bookViews><sheets>${wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${freshPrintNames(wb)}<calcPr calcId="191029"${wb.calcMode === 'manual' ? ' calcMode="manual"' : ''} fullCalcOnLoad="1"/></workbook>`);
  const n = wb.sheets.length;
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wb.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>${dynamic ? `<Relationship Id="rId${n + 3}" Type="${METADATA_REL}" Target="metadata.xml"/>` : ''}</Relationships>`);
  if (dynamic) zip.file('xl/metadata.xml', METADATA_XML);
  sheetXml.forEach((x, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, x));
  zip.file('xl/styles.xml', stylesXml);
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS}" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`);
  await precompress(JSZip, zip); // большие части — встроенным сжатием (zipBytes.js)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
