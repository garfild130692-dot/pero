// Разбиение листа на печатные страницы (как Excel): область печати (иначе использованная), скрытые строки и столбцы
// не печатаются, сквозные строки/столбцы повторяются на каждой странице, ручные разрывы (sheet.breaks), масштаб или
// «вписать» (fitWidth × fitHeight страниц), порядок «вниз, затем поперёк». Сверено с разрывами самого Excel
// (tests/make_pages_xlsx_fixture.ps1 — HPageBreaks/VPageBreaks).
import { PAPERS, pageOf } from './pageSetup.js';
import { DEFAULT_COL_W, DEFAULT_ROW_H, isHidden, usedRange } from './model.js';

// Excel печатает шрифтом принтера: высота строки на бумаге ≈ 0,97 экранной, ширина столбца ≈ 1,076 экранной; каждая
// строка и столбец округляются до точек принтера (600 точек на дюйм). Подобрано по разрывам Excel с принтером по
// умолчанию; на других принтерах Excel может разбить лист на строку иначе.
export const PRINT_K = { row: 0.97, col: 1.076 };
const PT_PER_PX = 0.75;
const DPI = 600;
const dots = (pt) => Math.round((pt * DPI) / 72);

// Область текста страницы, pt
export function pageArea(page) {
  const p = PAPERS[page.paper] || PAPERS[9];
  const land = page.orientation === 'landscape';
  const w = ((land ? p.h : p.w) / 25.4) * 72; const h = ((land ? p.w : p.h) / 25.4) * 72;
  const m = page.margins;
  return { paperW: w, paperH: h, W: w - (m.left + m.right) * 72, H: h - (m.top + m.bottom) * 72 };
}
// Размеры на бумаге (pt, при 100 %)
export const rowPrintPt = (sheet, r) => (sheet.rowH[r] ?? DEFAULT_ROW_H) * PT_PER_PX * PRINT_K.row;
export const colPrintPt = (sheet, c) => (sheet.colW[c] ?? DEFAULT_COL_W) * PT_PER_PX * PRINT_K.col;

// Нарезка по одной оси: позиции (видимые индексы) → массивы страниц; avail(firstOnPage) — место на странице
function cut(list, size, avail, manual) {
  const pages = []; let cur = []; let used = 0;
  for (const i of list) {
    const s = size(i);
    const startNew = cur.length && (manual.has(i) || used + s > avail(cur[0]));
    if (startNew) { pages.push(cur); cur = []; used = 0; }
    cur.push(i); used += s;
  }
  if (cur.length) pages.push(cur);
  return pages;
}

function layout(sheet, page, scale) {
  const area = sheet.printArea || usedRange(sheet);
  if (!area) return null;
  const rows = []; const cols = [];
  for (let r = area.r1; r <= area.r2; r++) if (!isHidden(sheet, 'row', r)) rows.push(r);
  for (let c = area.c1; c <= area.c2; c++) if (!isHidden(sheet, 'col', c)) cols.push(c);
  const { W, H } = pageArea(page);
  const tr = sheet.printTitles && sheet.printTitles.rows; const tc = sheet.printTitles && sheet.printTitles.cols;
  const titleRows = []; const titleCols = [];
  if (tr) for (let r = tr[0]; r <= tr[1]; r++) if (!isHidden(sheet, 'row', r)) titleRows.push(r);
  if (tc) for (let c = tc[0]; c <= tc[1]; c++) if (!isHidden(sheet, 'col', c)) titleCols.push(c);
  // размеры — в точках принтера с учётом масштаба
  const rowD = (r) => dots(rowPrintPt(sheet, r) * scale);
  const colD = (c) => dots(colPrintPt(sheet, c) * scale);
  const titleH = titleRows.reduce((a, r) => a + rowD(r), 0);
  const titleW = titleCols.reduce((a, c) => a + colD(c), 0);
  const availH = Math.floor((H * DPI) / 72); const availW = Math.floor((W * DPI) / 72);
  const brk = sheet.breaks || {};
  // сквозные строки печатаются на страницах, которые начинаются ниже них
  const rowPages = cut(rows, rowD, (first) => availH - (tr && first > tr[1] ? titleH : 0), new Set(brk.rows || []));
  const colPages = cut(cols, colD, (first) => availW - (tc && first > tc[1] ? titleW : 0), new Set(brk.cols || []));
  return { area, rowPages, colPages, titleRows, titleCols, tr, tc };
}

// Страницы листа: { scale, pages: [{ rows, cols, titleRows, titleCols, n }], rowBreaks, colBreaks } —
// rowBreaks/colBreaks — индексы строк/столбцов, с которых начинаются страницы (как HPageBreaks/VPageBreaks)
export function paginate(sheet) {
  const page = pageOf(sheet);
  let scale = (page.scale || 100) / 100;
  let lay;
  if (page.fitWidth != null || page.fitHeight != null) {
    // «вписать»: наибольший целый процент (10–100), при котором страниц не больше заданного
    const fw = page.fitWidth || 0; const fh = page.fitHeight || 0;
    let s = 100;
    for (; s > 10; s--) {
      lay = layout(sheet, page, s / 100);
      if (!lay) break;
      if ((!fw || lay.colPages.length <= fw) && (!fh || lay.rowPages.length <= fh)) break;
    }
    scale = s / 100;
  }
  lay = layout(sheet, page, scale);
  if (!lay) return { scale, pages: [], rowBreaks: [], colBreaks: [] };
  const pages = [];
  const order = page.pageOrder === 'overThenDown'
    ? lay.rowPages.flatMap((rp) => lay.colPages.map((cp) => [rp, cp]))
    : lay.colPages.flatMap((cp) => lay.rowPages.map((rp) => [rp, cp]));
  order.forEach(([rp, cp], i) => {
    pages.push({
      n: i + 1, rows: rp, cols: cp,
      titleRows: lay.tr && rp[0] > lay.tr[1] ? lay.titleRows : [],
      titleCols: lay.tc && cp[0] > lay.tc[1] ? lay.titleCols : [],
    });
  });
  return { scale, pages, rowBreaks: lay.rowPages.slice(1).map((p) => p[0]), colBreaks: lay.colPages.slice(1).map((p) => p[0]), area: lay.area };
}
