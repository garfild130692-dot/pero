// Импорт/экспорт .xlsx — собственный код (JSZip + xmldom), как и для .docx.
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { addr, colIndex, key, parseRange, unkey } from './address.js';
import { canonical, isErr, offsetFormula } from './formula.js';
import { BUILTIN_FMT } from './format.js';
import { newSheet, recalc, DEFAULT_COL_W, DEFAULT_ROW_H } from './model.js';

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
const widthToPx = (w) => Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
const pxToWidth = (px) => Math.trunc((px / 7) * 256) / 256;

// ---------- Цвета ----------
// Стандартная палитра индексных цветов Excel (0–63)
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
// Тон (tint) как в Excel: осветление/затемнение по светлоте HSL
function applyTint(hex, tint) {
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
    return Object.keys(s).length ? s : null;
  });
}

// ---------- Импорт ----------
/** @returns {Promise<{ wb, warnings }>} */
export async function importXlsx(data) {
  const zip = await JSZip.loadAsync(data);
  const read = async (p) => { const f = zip.file(p); return f ? parseXml(await f.async('string')) : null; };
  const rootRels = await read('_rels/.rels');
  let wbPath = 'xl/workbook.xml';
  for (const rel of kids(rootRels && rootRels.documentElement, 'Relationship')) {
    if (/officeDocument$/.test(rel.getAttribute('Type'))) wbPath = resolve('', rel.getAttribute('Target'));
  }
  const wbDoc = await read(wbPath);
  if (!wbDoc) throw new Error('Это не книга Excel: нет xl/workbook.xml');
  const dir = wbPath.slice(0, wbPath.lastIndexOf('/'));
  const relsDoc = await read(`${dir}/_rels/${wbPath.slice(dir.length + 1)}.rels`);
  const rels = {};
  for (const rel of kids(relsDoc && relsDoc.documentElement, 'Relationship')) rels[rel.getAttribute('Id')] = { type: rel.getAttribute('Type'), target: resolve(dir, rel.getAttribute('Target')) };
  const relOf = (type) => Object.values(rels).find((r) => r.type.endsWith(`/${type}`));

  const theme = parseTheme(relOf('theme') ? await read(relOf('theme').target) : null);
  const styles = parseStyles(relOf('styles') ? await read(relOf('styles').target) : null, theme);
  const sstDoc = relOf('sharedStrings') ? await read(relOf('sharedStrings').target) : null;
  const sst = sstDoc ? kids(sstDoc.documentElement, 'si').map((si) => {
    const t = kid(si, 't');
    if (t) return t.textContent;
    return kids(si, 'r').map((r) => (kid(r, 't') || {}).textContent || '').join('');
  }) : [];
  const date1904 = (() => { const p = kid(wbDoc.documentElement, 'workbookPr'); return !!p && ['1', 'true'].includes(p.getAttribute('date1904')); })();

  const warnings = {};
  const warn = (k) => { warnings[k] = (warnings[k] || 0) + 1; };
  const sheets = [];
  const sheetEls = kids(kid(wbDoc.documentElement, 'sheets'), 'sheet');
  for (const se of sheetEls) {
    const rid = se.getAttributeNS(R, 'id') || se.getAttribute('r:id');
    const rel = rels[rid];
    const sheet = newSheet(se.getAttribute('name'));
    if (!rel || !/\/worksheet$/.test(rel.type)) { warn('листы диаграмм'); continue; }
    const doc = await read(rel.target);
    if (!doc) continue;
    const root = doc.documentElement;
    for (const col of kids(kid(root, 'cols'), 'col')) {
      const w = +col.getAttribute('width');
      if (!w || !['1', 'true'].includes(col.getAttribute('customWidth') || '1')) continue;
      const min = +col.getAttribute('min') - 1;
      const max = Math.min(+col.getAttribute('max') - 1, min + 300);
      for (let c = min; c <= max; c++) if (widthToPx(w) !== DEFAULT_COL_W) sheet.colW[c] = widthToPx(w);
    }
    const shared = {}; // si → { f, r, c }
    for (const row of kids(kid(root, 'sheetData'), 'row')) {
      const ri = +row.getAttribute('r') - 1;
      const ht = row.getAttribute('ht');
      if (ht && ['1', 'true'].includes(row.getAttribute('customHeight'))) sheet.rowH[ri] = Math.round(+ht * 4 / 3);
      let cAuto = 0;
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r');
        let r = ri; let ci = cAuto;
        if (ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); ci = colIndex(m[1]); r = +m[2] - 1; }
        cAuto = ci + 1;
        const t = c.getAttribute('t') || 'n';
        const sIdx = +(c.getAttribute('s') || 0);
        const v = kid(c, 'v');
        const fEl = kid(c, 'f');
        const cell = {};
        const st = styles[sIdx];
        if (st) cell.s = st;
        // значение
        const vt = v ? v.textContent : null;
        if (t === 's') cell.v = sst[+vt] ?? '';
        else if (t === 'inlineStr') cell.v = kids(kid(c, 'is'), 't').map((x) => x.textContent).join('') || (kid(c, 'is') ? kid(c, 'is').textContent : '');
        else if (t === 'str') cell.v = vt ?? '';
        else if (t === 'b') cell.v = vt === '1';
        else if (t === 'e') cell.v = { error: vt || '#N/A' };
        else if (t === 'd') { const d = new Date(vt); cell.v = isNaN(d) ? vt : (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000; }
        else if (vt != null && vt !== '') {
          cell.v = +vt;
          if (date1904 && st && st.numFmt && /[dy]/i.test(st.numFmt)) cell.v += 1462;
        } else cell.v = null;
        // формула (общие формулы — развернуть со смещением от главной ячейки)
        if (fEl) {
          const ft = fEl.getAttribute('t');
          const text = fEl.textContent;
          if (ft === 'shared') {
            const si = fEl.getAttribute('si');
            if (text) { shared[si] = { f: text, r, c: ci }; cell.f = canonical(text); }
            else if (shared[si]) cell.f = offsetFormula(canonical(shared[si].f), r - shared[si].r, ci - shared[si].c);
          } else if (text) cell.f = canonical(text);
        }
        if (cell.v == null && cell.f == null && !cell.s) continue;
        if (cell.v === undefined) cell.v = null;
        sheet.cells.set(key(r, ci), cell);
      }
    }
    for (const m of kids(kid(root, 'mergeCells'), 'mergeCell')) {
      const rg = parseRange(m.getAttribute('ref'));
      if (rg) sheet.merges.push(rg);
    }
    if (kid(root, 'conditionalFormatting')) warn('условное форматирование');
    if (kid(root, 'dataValidations')) warn('проверка данных');
    if (kid(root, 'drawing')) warn('рисунки и диаграммы');
    if (kid(root, 'tableParts')) warn('«умные» таблицы');
    if (kid(root, 'legacyDrawing')) warn('примечания к ячейкам');
    if (kid(root, 'autoFilter')) warn('фильтры');
    sheets.push(sheet);
  }
  if (Object.values(rels).some((r) => /pivotCacheDefinition$/.test(r.type))) warn('сводные таблицы');
  // именованные диапазоны (служебные _xlnm.* — область печати, фильтр — не считаем)
  const names = kid(wbDoc.documentElement, 'definedNames');
  if (names && kids(names, 'definedName').some((n) => !/^_xlnm\./.test(n.getAttribute('name') || ''))) warn('именованные диапазоны');
  if (!sheets.length) sheets.push(newSheet('Лист1'));
  const view = kid(kid(wbDoc.documentElement, 'bookViews'), 'workbookView');
  const active = Math.min(+(view && view.getAttribute('activeTab')) || 0, sheets.length - 1);
  return { wb: recalc({ sheets, active }), warnings };
}

// ---------- Экспорт ----------
// Таблица стилей: уникальные шрифты/заливки/границы/форматы → cellXfs
function buildStyles(wb) {
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'];
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const numFmts = [];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const index = new Map(); // JSON стиля → номер xf
  const idxOf = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
  const BUILTIN_BY_CODE = Object.fromEntries(Object.entries(BUILTIN_FMT).map(([id, code]) => [code, +id]));
  const color = (c) => (c ? `<color rgb="FF${c.slice(1).toUpperCase()}"/>` : '');
  const xfOf = (s) => {
    if (!s) return 0;
    const k = JSON.stringify(s);
    if (index.has(k)) return index.get(k);
    const font = (s.bold || s.italic || s.underline || s.strike || s.color || s.fontSize || s.fontFamily)
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
      else {
        let i = numFmts.findIndex((x) => x === s.numFmt);
        if (i < 0) { numFmts.push(s.numFmt); i = numFmts.length - 1; }
        numFmtId = 164 + i;
      }
    }
    const al = (s.align || s.valign || s.wrap)
      ? `<alignment${s.align ? ` horizontal="${s.align}"` : ''}${s.valign ? ` vertical="${s.valign}"` : ''}${s.wrap ? ' wrapText="1"' : ''}/>` : '';
    const xml = `<xf numFmtId="${numFmtId}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}${border ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`;
    xfs.push(xml);
    index.set(k, xfs.length - 1);
    return xfs.length - 1;
  };
  // проходим все ячейки заранее, чтобы xml стилей был полным
  for (const sh of wb.sheets) for (const cell of sh.cells.values()) xfOf(cell.s);
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS}">` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(f)}"/>`).join('')}</numFmts>` : '') +
    `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills>` +
    `<borders count="${borders.length}">${borders.join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  return { xml, xfOf };
}

export async function exportXlsx(wb) {
  const { xml: stylesXml, xfOf } = buildStyles(wb);
  const sst = [];
  const sstIndex = new Map();
  const str = (s) => { if (!sstIndex.has(s)) { sstIndex.set(s, sst.length); sst.push(s); } return sstIndex.get(s); };

  const sheetXml = wb.sheets.map((sheet, si) => {
    const rows = new Map(); // r → [[c, xml]]
    for (const [k, cell] of sheet.cells) {
      const [r, c] = unkey(k);
      const ref = addr(r, c);
      const s = xfOf(cell.s);
      const sAttr = s ? ` s="${s}"` : '';
      const v = cell.v;
      let xml;
      if (cell.f != null) {
        const f = `<f>${esc(cell.f)}</f>`;
        if (isErr(v)) xml = `<c r="${ref}"${sAttr} t="e">${f}<v>${esc(v.error)}</v></c>`;
        else if (typeof v === 'string') xml = `<c r="${ref}"${sAttr} t="str">${f}<v>${esc(v)}</v></c>`;
        else if (typeof v === 'boolean') xml = `<c r="${ref}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
        else xml = `<c r="${ref}"${sAttr}>${f}${typeof v === 'number' && isFinite(v) ? `<v>${v}</v>` : ''}</c>`;
      } else if (v == null) xml = `<c r="${ref}"${sAttr}/>`;
      else if (typeof v === 'number') xml = `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
      else if (typeof v === 'boolean') xml = `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
      else if (isErr(v)) xml = `<c r="${ref}"${sAttr} t="e"><v>${esc(v.error)}</v></c>`;
      else xml = `<c r="${ref}"${sAttr} t="s"><v>${str(String(v))}</v></c>`;
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r).push([c, xml]);
    }
    for (const r of Object.keys(sheet.rowH)) if (!rows.has(+r)) rows.set(+r, []);
    const rowList = [...rows.keys()].sort((a, b) => a - b);
    let maxR = 0; let maxC = 0;
    for (const k of sheet.cells.keys()) { const [r, c] = unkey(k); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
    const cols = Object.keys(sheet.colW).map(Number).sort((a, b) => a - b)
      .map((c) => `<col min="${c + 1}" max="${c + 1}" width="${pxToWidth(sheet.colW[c])}" customWidth="1"/>`).join('');
    const data = rowList.map((r) => {
      const cells = rows.get(r).sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('');
      const h = sheet.rowH[r];
      return `<row r="${r + 1}"${h ? ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"` : ''}>${cells}</row>`;
    }).join('');
    const merges = sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${addr(m.r1, m.c1)}:${addr(m.r2, m.c2)}"/>`).join('')}</mergeCells>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS}" xmlns:r="${R}">` +
      `<dimension ref="A1${maxR || maxC ? `:${addr(maxR, maxC)}` : ''}"/>` +
      `<sheetViews><sheetView workbookViewId="0"${si === wb.active ? ' tabSelected="1"' : ''}/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${data}</sheetData>${merges}` +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  });

  const zip = new JSZip();
  const sheetsCt = wb.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheetsCt}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Pero</Application></Properties>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="${R}"><bookViews><workbookView activeTab="${wb.active || 0}"/></bookViews><sheets>${wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  const n = wb.sheets.length;
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wb.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
  sheetXml.forEach((x, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, x));
  zip.file('xl/styles.xml', stylesXml);
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS}" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
