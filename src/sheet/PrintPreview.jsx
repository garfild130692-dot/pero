// Печать листа готовыми страницами (printPages.js) и «Предварительный просмотр» (Ctrl+P, Файл → Печать) — как в
// Excel: слева параметры (ориентация, бумага, поля, масштабирование), справа страница, внизу листание.
import React, { useMemo, useState } from 'react';
import { key } from './address.js';
import { cellText } from './Grid.jsx';
import { formatCode } from './numfmt.js';
import { DEFAULT_COL_W, DEFAULT_ROW_H, mergeAt, valueAt } from './model.js';
import { PAPERS, MARGIN_PRESETS, pageOf } from './pageSetup.js';
import { PRINT_K, paginate, pageArea } from './printPages.js';

const BORDER_CSS = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted' };
const bcss = (b) => (b ? `${BORDER_CSS[b.style] || '1px solid'} ${b.color || '#000'}` : undefined);
const PX_PER_PT = 96 / 72;

// Одна страница: лист бумаги (px при 96 dpi), поля, таблица ячеек страницы в масштабе печати
export function SheetPage({ sheet, pg, scale, cf, className = '' }) {
  const page = pageOf(sheet);
  const a = pageArea(page);
  const m = page.margins;
  const rows = [...pg.titleRows, ...pg.rows];
  const cols = [...pg.titleCols, ...pg.cols];
  const colPx = (c) => (sheet.colW[c] ?? DEFAULT_COL_W) * PRINT_K.col;
  const rowPx = (r) => (sheet.rowH[r] ?? DEFAULT_ROW_H) * PRINT_K.row;
  const rowSet = new Set(rows); const colSet = new Set(cols);
  const tr = (r) => {
    const tds = [];
    for (const c of cols) {
      const mg = sheet.merges.length ? mergeAt(sheet, r, c) : null;
      // объединение: рисует первая видимая на этой странице ячейка, остальные пропускаются
      if (mg) {
        const firstR = rows.find((x) => x >= mg.r1 && x <= mg.r2); const firstC = cols.find((x) => x >= mg.c1 && x <= mg.c2);
        if (r !== firstR || c !== firstC) continue;
      }
      const own = sheet.cells.get(key(mg ? mg.r1 : r, mg ? mg.c1 : c));
      const vv = valueAt(sheet, mg ? mg.r1 : r, mg ? mg.c1 : c);
      const base = own && (own.v != null || own.f != null) ? own : { ...(own || {}), v: vv };
      const cfv = cf && cf.get(key(mg ? mg.r1 : r, mg ? mg.c1 : c));
      const cell = cfv ? { ...base, s: { ...(base.s || {}), ...cfv.style } } : base;
      const s = (cell && cell.s) || {};
      const b = s.border || {};
      const fmtColor = s.numFmt && typeof cell.v === 'number' ? formatCode(cell.v, s.numFmt).color : null;
      tds.push(
        <td key={c} rowSpan={mg ? rows.filter((x) => x >= mg.r1 && x <= mg.r2).length : undefined} colSpan={mg ? cols.filter((x) => x >= mg.c1 && x <= mg.c2).length : undefined}
          style={{ fontWeight: s.bold ? 'bold' : undefined, fontStyle: s.italic ? 'italic' : undefined, background: s.fill, color: fmtColor || s.color,
            textDecoration: [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
            fontSize: s.fontSize ? `${s.fontSize}pt` : undefined, fontFamily: s.fontFamily || undefined,
            textAlign: s.align || (cell && typeof cell.v === 'number' ? 'right' : typeof cell.v === 'boolean' ? 'center' : 'left'),
            verticalAlign: s.valign === 'top' ? 'top' : s.valign === 'center' ? 'middle' : 'bottom',
            whiteSpace: s.wrap ? 'pre-wrap' : 'pre',
            borderTop: bcss(b.top), borderBottom: bcss(b.bottom), borderLeft: bcss(b.left), borderRight: bcss(b.right) }}>
          {cfv && cfv.bar ? <span className="pp-bar" style={{ width: `${cfv.bar.pct * 100}%`, background: `linear-gradient(90deg, ${cfv.bar.color}, ${cfv.bar.color}33)` }} /> : null}
          {cellText(cell)}
        </td>,
      );
    }
    return <tr key={r} style={{ height: rowPx(r) }}>{tds}</tr>;
  };
  const tableW = cols.reduce((x, c) => x + colPx(c), 0);
  void rowSet; void colSet;
  return (
    <div className={`pp-paper ${className}`} style={{ width: a.paperW * PX_PER_PT, height: a.paperH * PX_PER_PT }}>
      <div className="pp-content" style={{ left: m.left * 96, top: m.top * 96, width: a.W * PX_PER_PT, height: a.H * PX_PER_PT }}>
        <table className={`pp-table${page.gridLines ? ' grid' : ''}`} style={{ width: tableW, zoom: scale, margin: page.centerH ? '0 auto' : undefined }}>
          <colgroup>{cols.map((c) => <col key={c} style={{ width: colPx(c) }} />)}</colgroup>
          <tbody>{rows.map(tr)}</tbody>
        </table>
      </div>
    </div>
  );
}

// Все страницы для печати (видны только в @media print): каждая — отдельный лист
export function PrintPages({ sheet, cf }) {
  const p = useMemo(() => paginate(sheet), [sheet]);
  const page = pageOf(sheet);
  const a = pageArea(page);
  if (!p.pages.length) return null;
  return (
    <div className="pp-print">
      <style>{`@media print { @page { size: ${a.paperW}pt ${a.paperH}pt; margin: 0; } }`}</style>
      {p.pages.map((pg) => <SheetPage key={pg.n} sheet={sheet} pg={pg} scale={p.scale} cf={cf} className="pp-sheet" />)}
    </div>
  );
}

const SCALINGS = [
  ['cur', 'Текущий', 'Листы печатаются в натуральную величину'],
  ['sheet', 'Вписать лист на одну страницу', 'Уменьшить так, чтобы всё поместилось на одной странице'],
  ['cols', 'Вписать все столбцы на одну страницу', 'Уменьшить так, чтобы ширина помещалась на одной странице'],
  ['rows', 'Вписать все строки на одну страницу', 'Уменьшить так, чтобы высота помещалась на одной странице'],
];
const scalingOf = (pg) => (pg.fitWidth === 1 && pg.fitHeight === 1 ? 'sheet' : pg.fitWidth === 1 && !pg.fitHeight ? 'cols' : !pg.fitWidth && pg.fitHeight === 1 ? 'rows' : 'cur');

export default function PrintPreview({ sheet, cf, onPage, onPrint, onPdf, onClose }) {
  const p = useMemo(() => paginate(sheet), [sheet]);
  const page = pageOf(sheet);
  const [n, setN] = useState(0);
  const cur = Math.min(n, Math.max(0, p.pages.length - 1));
  const pg = p.pages[cur];
  const marginsName = Object.entries(MARGIN_PRESETS).find(([, mp]) => ['left', 'right', 'top', 'bottom'].every((k) => Math.abs(mp[k] - page.margins[k]) < 1e-6));
  const setScaling = (k) => onPage(k === 'sheet' ? { fitWidth: 1, fitHeight: 1 } : k === 'cols' ? { fitWidth: 1, fitHeight: 0 } : k === 'rows' ? { fitWidth: 0, fitHeight: 1 } : { fitWidth: null, fitHeight: null, scale: 100 });
  return (
    <div className="pp-overlay" role="dialog" aria-label="Печать" onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'PageDown') setN(Math.min(cur + 1, p.pages.length - 1)); if (e.key === 'PageUp') setN(Math.max(cur - 1, 0)); }} tabIndex={-1}>
      <div className="pp-side">
        <button type="button" className="pp-back" onClick={onClose} title="Назад к книге (Esc)">← Назад</button>
        <h2>Печать</h2>
        <div className="pp-actions">
          <button type="button" className="btn primary" onClick={onPrint}>Печать</button>
          <button type="button" className="btn" onClick={onPdf}>Экспорт в PDF</button>
        </div>
        <div className="pp-sub">Параметры</div>
        <label className="pp-field"><span>Ориентация</span>
          <select value={page.orientation} onChange={(e) => onPage({ orientation: e.target.value })}><option value="portrait">Книжная</option><option value="landscape">Альбомная</option></select>
        </label>
        <label className="pp-field"><span>Размер бумаги</span>
          <select value={page.paper} onChange={(e) => onPage({ paper: +e.target.value })}>{Object.entries(PAPERS).map(([id, x]) => <option key={id} value={id}>{x.name}</option>)}</select>
        </label>
        <label className="pp-field"><span>Поля</span>
          <select value={marginsName ? marginsName[0] : ''} onChange={(e) => { const mp = MARGIN_PRESETS[e.target.value]; if (mp) onPage({ margins: { left: mp.left, right: mp.right, top: mp.top, bottom: mp.bottom, header: mp.header, footer: mp.footer } }); }}>
            {!marginsName && <option value="">Настраиваемые</option>}
            {Object.entries(MARGIN_PRESETS).map(([k, mp]) => <option key={k} value={k}>{mp.name}</option>)}
          </select>
        </label>
        <label className="pp-field"><span>Масштабирование</span>
          <select value={scalingOf(page)} onChange={(e) => setScaling(e.target.value)}>{SCALINGS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
        </label>
        {scalingOf(page) === 'cur' && (
          <label className="pp-field"><span>Масштаб, %</span><input type="number" className="pd-input" min="10" max="400" value={page.scale || 100} onChange={(e) => onPage({ scale: Math.max(10, Math.min(400, +e.target.value || 100)) })} /></label>
        )}
        <label className="fc-check"><input type="checkbox" checked={!!page.gridLines} onChange={(e) => onPage({ gridLines: e.target.checked })} /> Печатать сетку</label>
        <label className="fc-check"><input type="checkbox" checked={!!page.centerH} onChange={(e) => onPage({ centerH: e.target.checked })} /> Центрировать по горизонтали</label>
        <div className="pp-note">Масштаб печати: {Math.round(p.scale * 100)} %</div>
      </div>
      <div className="pp-main">
        <div className="pp-stage">
          {pg ? <SheetPage sheet={sheet} pg={pg} scale={p.scale} cf={cf} className="pp-shadow" /> : <div className="pp-empty">Нечего печатать: лист пуст</div>}
        </div>
        <div className="pp-nav">
          <button type="button" className="btn small" disabled={cur === 0} onClick={() => setN(cur - 1)} title="Предыдущая страница (PageUp)">◀</button>
          <span>{p.pages.length ? `${cur + 1} из ${p.pages.length}` : '0 из 0'}</span>
          <button type="button" className="btn small" disabled={cur >= p.pages.length - 1} onClick={() => setN(cur + 1)} title="Следующая страница (PageDown)">▶</button>
        </div>
      </div>
    </div>
  );
}
