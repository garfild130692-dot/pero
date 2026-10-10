// Параметры страницы листа (как «Разметка страницы» Excel): бумага, ориентация, поля, масштаб / «вписать»,
// сетка при печати, область печати и сквозные строки/столбцы. Общие для ленты, печати и .xlsx.
//
// sheet.page = { orientation: 'portrait'|'landscape', paper (код paperSize OOXML), scale (%), fitWidth, fitHeight
//               (число страниц, 0 — «авто»; заданы — режим «вписать»), margins { left, right, top, bottom, header, footer }
//               (дюймы, как в файле), gridLines, centerH }
// sheet.printArea = { r1, c1, r2, c2 } | null;  sheet.printTitles = { rows: [r1, r2] | null, cols: [c1, c2] | null } | null
import { colIndex, colName } from './address.js';

// Размеры бумаги (мм) по кодам paperSize; без атрибута — Letter (по спецификации)
export const PAPERS = {
  9: { name: 'A4', w: 210, h: 297 },
  8: { name: 'A3', w: 297, h: 420 },
  11: { name: 'A5', w: 148, h: 210 },
  1: { name: 'Letter', w: 215.9, h: 279.4 },
  5: { name: 'Legal', w: 215.9, h: 355.6 },
};
// Поля (дюймы) — наборы Excel
export const MARGIN_PRESETS = {
  normal: { name: 'Обычные', left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { name: 'Широкие', left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { name: 'Узкие', left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
};
// Новый лист Перо: A4 книжная, поля «Обычные»
export const DEFAULT_PAGE = { orientation: 'portrait', paper: 9, scale: 100, fitWidth: null, fitHeight: null, margins: MARGIN_PRESETS.normal, gridLines: false, centerH: false };
export const pageOf = (sheet) => ({ ...DEFAULT_PAGE, ...(sheet.page || {}), margins: { ...DEFAULT_PAGE.margins, ...((sheet.page && sheet.page.margins) || {}) } });

// Изменить параметры страницы листа (patch — частичный)
export function setPage(wb, si, patch) {
  const sheet = wb.sheets[si];
  const page = { ...(sheet.page || {}), ...patch };
  return { ...wb, sheets: wb.sheets.map((s, i) => (i === si ? { ...sheet, page } : s)) };
}
export function setPrintArea(wb, si, rg) {
  const sheet = { ...wb.sheets[si], printArea: rg || null };
  if (!rg) delete sheet.printArea;
  return { ...wb, sheets: wb.sheets.map((s, i) => (i === si ? sheet : s)) };
}
export function setPrintTitles(wb, si, titles) {
  const t = titles && (titles.rows || titles.cols) ? { rows: titles.rows || null, cols: titles.cols || null } : null;
  const sheet = { ...wb.sheets[si], printTitles: t };
  if (!t) delete sheet.printTitles;
  return { ...wb, sheets: wb.sheets.map((s, i) => (i === si ? sheet : s)) };
}

// ---------- Имена печати в книге (_xlnm.Print_Area, _xlnm.Print_Titles) ----------
// Текст определённого имени: Лист!$A$1:$D$14 (несколько областей через запятую — берётся первая)
const AREA_RE = /!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/;
export function parsePrintArea(text) {
  const m = AREA_RE.exec(text || '');
  if (!m) return null;
  const c1 = colIndex(m[1]); const r1 = +m[2] - 1;
  const c2 = m[3] ? colIndex(m[3]) : c1; const r2 = m[4] ? +m[4] - 1 : r1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}
// Лист!$1:$2 и/или Лист!$A:$B
export function parsePrintTitles(text) {
  const out = { rows: null, cols: null };
  for (const part of String(text || '').split(',')) {
    const ref = part.slice(part.lastIndexOf('!') + 1);
    let m = /^\$?(\d+):\$?(\d+)$/.exec(ref);
    if (m) { out.rows = [+m[1] - 1, +m[2] - 1]; continue; }
    m = /^\$?([A-Z]+):\$?([A-Z]+)$/.exec(ref);
    if (m) out.cols = [colIndex(m[1]), colIndex(m[2])];
  }
  return out.rows || out.cols ? out : null;
}
const sheetRef = (name) => (/^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
export const printAreaText = (name, rg) => `${sheetRef(name)}!$${colName(rg.c1)}$${rg.r1 + 1}:$${colName(rg.c2)}$${rg.r2 + 1}`;
export function printTitlesText(name, t) {
  const parts = [];
  if (t.cols) parts.push(`${sheetRef(name)}!$${colName(t.cols[0])}:$${colName(t.cols[1])}`);
  if (t.rows) parts.push(`${sheetRef(name)}!$${t.rows[0] + 1}:$${t.rows[1] + 1}`);
  return parts.join(',');
}

// ---------- Печать ----------
const MM_PER_IN = 25.4;
const PX_PER_MM = 96 / 25.4;
// Размер листа и полей для @page; область текста в px — для масштаба «вписать»
export function printGeometry(page) {
  const p = PAPERS[page.paper] || PAPERS[9];
  const land = page.orientation === 'landscape';
  const w = land ? p.h : p.w; const h = land ? p.w : p.h;
  const m = page.margins;
  return {
    css: `@page { size: ${w}mm ${h}mm; margin: ${m.top * MM_PER_IN}mm ${m.right * MM_PER_IN}mm ${m.bottom * MM_PER_IN}mm ${m.left * MM_PER_IN}mm; }`,
    contentW: (w - (m.left + m.right) * MM_PER_IN) * PX_PER_MM,
    contentH: (h - (m.top + m.bottom) * MM_PER_IN) * PX_PER_MM,
  };
}
// Масштаб печати: «вписать» по ширине/высоте (страниц) или заданный процент
export function printScale(page, tableW, tableH) {
  const g = printGeometry(page);
  if (page.fitWidth != null || page.fitHeight != null) {
    let k = 1;
    if (page.fitWidth) k = Math.min(k, (g.contentW * page.fitWidth) / tableW);
    if (page.fitHeight) k = Math.min(k, (g.contentH * page.fitHeight) / tableH);
    return Math.max(0.1, k);
  }
  return (page.scale || 100) / 100;
}
