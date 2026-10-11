// «Макет → Колонки»: меню заготовок (одна, две, три, слева, справа) и диалог «Колонки» (число, ширина и промежуток,
// разделитель, «Применить»). Выделен текст — колонки ему (отдельный раздел «на текущей странице»), иначе — разделу.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { columnsAt } from './editor/colCommands.js';
import { parseSectPr, TWIPS_PER_CM } from './editor/page.js';
import { fmtNum } from './editor/paragraph.js';

const cm = (tw) => fmtNum(tw / TWIPS_PER_CM, 2);
const num = (s) => { const v = parseFloat(String(s).replace(',', '.')); return isFinite(v) ? v : null; };
const DEFAULT_SPACE = 708; // 1,25 см — промежуток Word по умолчанию

// Ширина текста раздела (твипы)
const textWidth = (xml) => { const g = parseSectPr(xml); return g.w - g.left - g.right; };

// Заготовки: колонки разной ширины — 1/3 и 2/3 (как «Слева»/«Справа» Word)
export function presetCols(kind, W, space = DEFAULT_SPACE) {
  if (kind === 'one') return null;
  if (kind === 'two') return { num: 2, space, sep: false, equal: true, cols: [] };
  if (kind === 'three') return { num: 3, space, sep: false, equal: true, cols: [] };
  const narrow = Math.round((W - space) / 3); const wide = W - space - narrow;
  const cols = kind === 'left' ? [{ w: narrow, space }, { w: wide, space: 0 }] : [{ w: wide, space }, { w: narrow, space: 0 }];
  return { num: 2, space, sep: false, equal: false, cols };
}
const presetOf = (c) => {
  if (!c) return 'one';
  if (c.equal) return c.num === 2 ? 'two' : c.num === 3 ? 'three' : null;
  if (c.num === 2 && c.cols.length === 2) return c.cols[0].w < c.cols[1].w ? 'left' : 'right';
  return null;
};

const Icon = ({ kind }) => {
  const bars = { one: [[3, 18]], two: [[3, 7.5], [13.5, 7.5]], three: [[3, 4.4], [9.8, 4.4], [16.6, 4.4]], left: [[3, 4.5], [10.5, 10.5]], right: [[3, 10.5], [16.5, 4.5]] }[kind];
  return (
    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.4" aria-hidden="true">
      {bars.map(([x, w]) => [5, 9, 13, 17].map((y) => <path key={`${x}-${y}`} d={`M${x} ${y}h${w}`} />))}
    </svg>
  );
};
const PRESETS = [['one', 'Одна'], ['two', 'Две'], ['three', 'Три'], ['left', 'Слева'], ['right', 'Справа']];

export function ColumnsButton({ editor }) {
  const pop = usePopover();
  const [dlg, setDlg] = useState(false);
  const { cols, xml } = columnsAt(editor.state);
  const cur = presetOf(cols);
  const scope = editor.state.selection.empty ? 'section' : 'selection';
  const apply = (kind) => { pop.setOpen(false); editor.chain().focus().setColumns(presetCols(kind, textWidth(xml)), scope).run(); };
  return (
    <>
      <div className="rb-popover-anchor" ref={pop.ref}>
        <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title="Колонки текста"
          onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true">
            <rect x="3" y="3" width="18" height="18" rx="1" /><path d="M6 7h5M6 10h5M6 13h5M6 16h5M13 7h5M13 10h5M13 13h5M13 16h3" strokeWidth="1.2" />
          </svg>
          <span className="rb-label">Колонки ▾</span>
        </button>
        {pop.open && (
          <div className="menu-pop" style={pop.style} role="menu">
            {PRESETS.map(([k, label]) => (
              <button key={k} type="button" className={`menu-item cols-item${cur === k ? ' on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => apply(k)}>
                <Icon kind={k} />{label}
              </button>
            ))}
            <div className="menu-sep" />
            <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); setDlg(true); }}>
              <span className="check" />Другие колонки…
            </button>
          </div>
        )}
      </div>
      {dlg && <ColumnsDialog editor={editor} onClose={() => setDlg(false)} />}
    </>
  );
}

export function ColumnsDialog({ editor, onClose }) {
  const { cols, xml } = columnsAt(editor.state);
  const W = textWidth(xml);
  const hasSel = !editor.state.selection.empty;
  // строки ширины/промежутка (твипы)
  const rowsOf = (c) => {
    const n = c ? c.num : 1;
    if (c && !c.equal && c.cols.length === n) return c.cols.map((x) => ({ w: x.w, space: x.space }));
    const space = c ? c.space : DEFAULT_SPACE;
    const w = Math.round((W - space * (n - 1)) / n);
    return Array.from({ length: n }, (_, i) => ({ w, space: i < n - 1 ? space : 0 }));
  };
  const [n, setN] = useState(cols ? cols.num : 1);
  const [equal, setEqual] = useState(!cols || cols.equal);
  const [sep, setSep] = useState(!!(cols && cols.sep));
  const [rows, setRows] = useState(() => rowsOf(cols).map((r) => ({ w: cm(r.w), space: cm(r.space) })));
  const [scope, setScope] = useState(hasSel ? 'selection' : 'section');
  const [err, setErr] = useState('');

  // ширины поровну при числе колонок count (промежуток — как у первой колонки)
  const recalc = (count) => {
    const sp = num(rows[0] && rows[0].space);
    const space = sp != null && sp > 0 ? Math.round(sp * TWIPS_PER_CM) : DEFAULT_SPACE;
    const r = rowsOf(count > 1 ? { num: count, space, equal: true, cols: [] } : null);
    setRows(r.map((x) => ({ w: cm(x.w), space: cm(x.space) })));
  };
  const pick = (kind) => {
    const c = presetCols(kind, W);
    setN(c ? c.num : 1); setEqual(!c || c.equal); setErr('');
    setRows(rowsOf(c).map((x) => ({ w: cm(x.w), space: cm(x.space) })));
  };
  const setCount = (v) => {
    const c = Math.max(1, Math.min(12, parseInt(v, 10) || 1));
    setN(c); setErr(''); recalc(c);
  };
  // правка ширины/промежутка: при «одинаковой ширине» — ширины всех колонок из промежутка
  const edit = (i, k, v) => {
    setErr('');
    const next = rows.map((r, j) => (j === i ? { ...r, [k]: v } : r));
    if (equal && k === 'space' && num(v) != null) {
      const sp = Math.round(num(v) * TWIPS_PER_CM);
      const w = Math.round((W - sp * (n - 1)) / n);
      setRows(next.map((r, j) => ({ w: cm(w), space: j < n - 1 ? cm(sp) : cm(0) })));
      return;
    }
    if (equal && k === 'w' && num(v) != null) {
      const w = Math.round(num(v) * TWIPS_PER_CM);
      const sp = n > 1 ? Math.round((W - w * n) / (n - 1)) : 0;
      setRows(next.map((r, j) => ({ w: j === i ? v : cm(w), space: j < n - 1 ? cm(Math.max(0, sp)) : cm(0) })));
      return;
    }
    setRows(next);
  };

  const apply = () => {
    let out = null;
    if (n > 1) {
      const tw = rows.slice(0, n).map((r) => ({ w: num(r.w) == null ? null : Math.round(num(r.w) * TWIPS_PER_CM), space: num(r.space) == null ? null : Math.round(num(r.space) * TWIPS_PER_CM) }));
      if (tw.some((r) => r.w == null || r.w < 284 || r.space == null || r.space < 0)) { setErr('Ширина колонки — не меньше 0,5 см, промежуток — не меньше 0'); return; }
      const total = tw.reduce((s, r, i) => s + r.w + (i < n - 1 ? r.space : 0), 0);
      if (total > W + 30) { setErr(`Колонки с промежутками шире текста (${cm(W)} см)`); return; }
      out = equal ? { num: n, space: tw[0].space, sep, equal: true, cols: [] } : { num: n, space: tw[0].space, sep, equal: false, cols: tw.map((r, i) => ({ w: r.w, space: i < n - 1 ? r.space : 0 })) };
    }
    editor.chain().focus().setColumns(out, scope).run();
    onClose();
  };

  // образец: лист с полосами колонок
  const preview = () => {
    const tw = rows.slice(0, n).map((r) => Math.max(0, num(r.w) || 0));
    const sp = rows.slice(0, n).map((r) => Math.max(0, num(r.space) || 0));
    const total = tw.reduce((s, w, i) => s + w + (i < n - 1 ? sp[i] : 0), 0) || 1;
    const k = 76 / total;
    let x = 12;
    const out = [];
    for (let i = 0; i < n; i++) {
      for (let y = 14; y < 100; y += 6) out.push(<path key={`${i}-${y}`} d={`M${x.toFixed(1)} ${y}h${(tw[i] * k).toFixed(1)}`} />);
      if (sep && i < n - 1) out.push(<path key={`s${i}`} d={`M${(x + tw[i] * k + sp[i] * k / 2).toFixed(1)} 12v90`} stroke="#000" />);
      x += (tw[i] + (sp[i] || 0)) * k;
    }
    return <svg className="cols-preview" width="100" height="112" viewBox="0 0 100 112" fill="none" stroke="#888" strokeWidth="1.5"><rect x="2" y="2" width="96" height="108" fill="#fff" stroke="#999" />{out}</svg>;
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd cols-dlg" role="dialog" aria-label="Колонки" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Колонки</div>
        <div className="modal-body">
          <div className="cols-presets">
            {PRESETS.map(([k, label]) => (
              <button key={k} type="button" className={`cols-preset${presetOf(n > 1 ? { num: n, equal, cols: rows.map((r) => ({ w: num(r.w) || 0 })) } : null) === k ? ' on' : ''}`} onClick={() => pick(k)}>
                <Icon kind={k} /><span>{label}</span>
              </button>
            ))}
          </div>
          <div className="cols-top">
            <label className="pd-field"><span>Число колонок</span>
              <span className="pd-input small"><input type="number" min="1" max="12" value={n} onChange={(e) => setCount(e.target.value)} /></span>
            </label>
            <label className="pd-check"><input type="checkbox" checked={sep} disabled={n < 2} onChange={(e) => setSep(e.target.checked)} /> Разделитель</label>
          </div>
          <fieldset><legend>Ширина и промежуток</legend>
            <table className="cols-table">
              <thead><tr><th>Колонка</th><th>Ширина, см</th><th>Промежуток, см</th></tr></thead>
              <tbody>
                {rows.slice(0, n).map((r, i) => (
                  <tr key={i}>
                    <td>{i + 1}:</td>
                    <td><input disabled={equal && i > 0} value={r.w} onChange={(e) => edit(i, 'w', e.target.value)} /></td>
                    <td><input disabled={i === n - 1 || (equal && i > 0)} value={i === n - 1 ? '' : r.space} onChange={(e) => edit(i, 'space', e.target.value)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <label className="pd-check"><input type="checkbox" checked={equal} disabled={n < 2} onChange={(e) => { setEqual(e.target.checked); if (e.target.checked) recalc(n); }} /> колонки одинаковой ширины</label>
            {!equal && n > 1 && <div className="hf-hint">Колонки разной ширины сохраняются в файле как заданы; на экране и при печати Перо показывает их одинаковыми.</div>}
          </fieldset>
          <div className="ps-bottom">
            <fieldset className="ps-sample"><legend>Образец</legend>{preview()}</fieldset>
            <label className="pd-field"><span>Применить</span>
              <select value={scope} onChange={(e) => setScope(e.target.value)}>
                <option value="section">к этому разделу</option>
                {hasSel && <option value="selection">к выделенному тексту</option>}
                <option value="forward">до конца документа</option>
                <option value="all">ко всему документу</option>
              </select>
            </label>
          </div>
          {err && <div className="pd-error" role="alert">{err}</div>}
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
