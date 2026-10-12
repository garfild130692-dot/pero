// «Упорядочение» (вкладки «Рисунок» и «Макет»): обтекание текстом, положение на странице (9 мест, как в Word) и диалог
// «Макет» — положение по горизонтали/вертикали (выравнивание или смещение от колонки/поля/листа/абзаца), обтекание,
// с какой стороны текст и отступы от текста. Модель — anchor.js, показ — floatLayout.js.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { DEFAULT_DIST, WRAP_KINDS, makeAnchor, wrapKind } from './editor/anchor.js';
import { fmtNum } from './editor/paragraph.js';

const PX_PER_CM = 96 / 2.54;
const cm = (px) => fmtNum((px || 0) / PX_PER_CM, 2);
const num = (s) => { const v = parseFloat(String(s).replace(',', '.')); return isFinite(v) ? v : null; };

// значки обтекания: прямоугольник рисунка и строки текста вокруг
function WrapIcon({ kind }) {
  const lines = { inline: [[3, 21]], square: [[3, 8], [16, 21]], tight: [[3, 8], [16, 21]], through: [[3, 8], [16, 21]], topBottom: [], behind: [[3, 21]], front: [[3, 21]] }[kind] || [];
  return (
    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.3" aria-hidden="true">
      {kind === 'topBottom' ? <><path d="M3 4h18M3 20h18" /><rect x="8" y="7" width="8" height="10" fill="#cfe3dc" /></>
        : <>
          {[5, 9, 13, 17].map((y) => lines.map(([a, b]) => <path key={`${y}-${a}`} d={`M${a} ${y}h${b - a}`} opacity={kind === 'behind' ? 0.9 : 1} />))}
          <rect x={kind === 'inline' ? 3 : 9} y={kind === 'inline' ? 15 : 7} width={kind === 'inline' ? 5 : 6} height={kind === 'inline' ? 6 : 9}
            fill={kind === 'behind' ? '#cfe3dc80' : '#cfe3dc'} stroke={kind === 'front' ? '#1f6f5c' : 'currentColor'} />
        </>}
    </svg>
  );
}

// 9 положений Word («Положение»): поле страницы, обтекание вокруг рамки
const POSITIONS = [
  ['top', 'left', 'Сверху слева'], ['top', 'center', 'Сверху по центру'], ['top', 'right', 'Сверху справа'],
  ['center', 'left', 'Посередине слева'], ['center', 'center', 'В центре'], ['center', 'right', 'Посередине справа'],
  ['bottom', 'left', 'Снизу слева'], ['bottom', 'center', 'Снизу по центру'], ['bottom', 'right', 'Снизу справа'],
];

function Menu({ icon, label, title, disabled, children }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title={title} disabled={disabled}
        onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
        {icon}<span className="rb-label">{label} ▾</span>
      </button>
      {pop.open && <div className="menu-pop" style={pop.style} role="menu">{children(() => pop.setOpen(false))}</div>}
    </div>
  );
}
const Item = ({ onClick, on, icon, children }) => (
  <button type="button" className={`menu-item arrange-item${on ? ' on' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
    {icon || <span className="check">{on ? '✓' : ''}</span>}{children}
  </button>
);

// Рисунок из строки текста получает обтекание там, где стоял (как Word): смещение от колонки и от абзаца по экрану
function inlinePlace(editor) {
  const { view, state } = editor;
  const pos = state.selection.from;
  const dom = view.nodeDOM(pos);
  const img = dom && dom.querySelector && dom.querySelector('img');
  const para = dom && dom.closest && dom.closest('p, h1, h2, h3, h4, h5, h6');
  if (!img || !para) return null;
  const root = view.dom;
  const rr = root.getBoundingClientRect();
  const scale = root.offsetWidth ? rr.width / root.offsetWidth : 1;
  const colL = rr.left + parseFloat(getComputedStyle(root).paddingLeft) * scale;
  const r = img.getBoundingClientRect();
  const round = (v) => Math.round(v * 100) / 100;
  return { h: { rel: 'column', align: null, off: round((r.left - colL) / scale) }, v: { rel: 'paragraph', align: null, off: round((r.top - para.getBoundingClientRect().top) / scale) } };
}

export function ArrangeGroup({ editor, image, Group }) {
  const [dlg, setDlg] = useState(false);
  const a = image ? image.anchor : null;
  const kind = image ? wrapKind(a) : null;
  const set = (anchor) => editor.chain().focus().setImageAnchor(anchor).run();
  const posOn = (v, h) => a && a.wrap !== 'none' && a.v.rel === 'margin' && a.v.align === v && a.h.rel === 'margin' && a.h.align === h;
  return (
    <Group title="Упорядочение">
      <Menu icon={<WrapIcon kind="square" />} label="Обтекание текстом" title="Как текст обтекает рисунок" disabled={!image}>
        {(close) => (
          <>
            {WRAP_KINDS.map(([k, label]) => (
              <Item key={k} on={kind === k} icon={<WrapIcon kind={k} />} onClick={() => {
                close();
                const next = makeAnchor(k, a);
                set(next && !a ? { ...next, ...(inlinePlace(editor) || {}) } : next);
              }}>{label}</Item>
            ))}
            <div className="menu-sep" />
            <Item onClick={() => { close(); setDlg(true); }}>Другие параметры разметки…</Item>
          </>
        )}
      </Menu>
      <Menu icon={<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.4" aria-hidden="true"><rect x="4" y="2" width="16" height="20" /><rect x="12" y="5" width="5" height="5" fill="#cfe3dc" /><path d="M7 13h10M7 16h10M7 19h7M7 6h3M7 9h3" strokeWidth="1.1" /></svg>}
        label="Положение" title="Положение рисунка на странице" disabled={!image}>
        {(close) => (
          <>
            <Item on={kind === 'inline'} onClick={() => { close(); set(null); }}>В тексте</Item>
            <div className="menu-sep" />
            <div className="pos-grid">
              {POSITIONS.map(([v, h, label]) => (
                <button key={label} type="button" title={`${label}, обтекание вокруг рамки`} className={`pos-cell${posOn(v, h) ? ' on' : ''}`}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => { close(); set({ ...makeAnchor('square', a), h: { rel: 'margin', align: h, off: 0 }, v: { rel: 'margin', align: v, off: 0 } }); }}>
                  <span style={{ justifyContent: { left: 'flex-start', center: 'center', right: 'flex-end' }[h], alignItems: { top: 'flex-start', center: 'center', bottom: 'flex-end' }[v] }}><i /></span>
                </button>
              ))}
            </div>
            <div className="menu-sep" />
            <Item onClick={() => { close(); setDlg(true); }}>Другие параметры разметки…</Item>
          </>
        )}
      </Menu>
      {dlg && image && <LayoutDialog editor={editor} image={image} onClose={() => setDlg(false)} />}
    </Group>
  );
}

const H_REL = [['column', 'колонки'], ['margin', 'поля'], ['page', 'страницы'], ['character', 'символа'], ['leftMargin', 'левого поля'], ['rightMargin', 'правого поля']];
const V_REL = [['paragraph', 'абзаца'], ['margin', 'поля'], ['page', 'страницы'], ['line', 'строки'], ['topMargin', 'верхнего поля'], ['bottomMargin', 'нижнего поля']];
const H_ALIGN = [['left', 'по левому краю'], ['center', 'по центру'], ['right', 'по правому краю']];
const V_ALIGN = [['top', 'сверху'], ['center', 'по центру'], ['bottom', 'снизу']];
const SIDES = [['bothSides', 'с обеих сторон'], ['left', 'только слева'], ['right', 'только справа'], ['largest', 'по большей стороне']];

// Диалог «Макет»: положение и обтекание
export function LayoutDialog({ editor, image, onClose }) {
  const a0 = image.anchor || makeAnchor('square', null);
  const [tab, setTab] = useState(image.anchor ? 'position' : 'wrap');
  const [kind, setKind] = useState(wrapKind(image.anchor));
  const [hMode, setHMode] = useState(a0.h.align ? 'align' : 'off');
  const [h, setH] = useState({ rel: a0.h.rel, align: a0.h.align || 'left', off: cm(a0.h.off) });
  const [vMode, setVMode] = useState(a0.v.align ? 'align' : 'off');
  const [v, setV] = useState({ rel: a0.v.rel, align: a0.v.align || 'top', off: cm(a0.v.off) });
  const [side, setSide] = useState(a0.side || 'bothSides');
  const d0 = a0.dist || DEFAULT_DIST;
  const [dist, setDist] = useState({ t: cm(d0.t), b: cm(d0.b), l: cm(d0.l), r: cm(d0.r) });
  const [err, setErr] = useState('');
  const floating = kind !== 'inline';
  const apply = () => {
    if (!floating) { editor.chain().focus().setImageAnchor(null).run(); onClose(); return; }
    const vals = [h.off, v.off, dist.t, dist.b, dist.l, dist.r].map(num);
    if (vals.some((x) => x == null) || [dist.t, dist.b, dist.l, dist.r].some((x) => num(x) < 0)) { setErr('Значения — числа в сантиметрах, отступы — не меньше 0'); return; }
    const base = makeAnchor(kind, a0);
    const next = {
      ...base,
      side,
      h: { rel: h.rel, align: hMode === 'align' ? h.align : null, off: hMode === 'align' ? 0 : num(h.off) * PX_PER_CM },
      v: { rel: v.rel, align: vMode === 'align' ? v.align : null, off: vMode === 'align' ? 0 : num(v.off) * PX_PER_CM },
      dist: { t: num(dist.t) * PX_PER_CM, b: num(dist.b) * PX_PER_CM, l: num(dist.l) * PX_PER_CM, r: num(dist.r) * PX_PER_CM },
    };
    editor.chain().focus().setImageAnchor(next).run();
    onClose();
  };
  const sel = (value, set, opts) => <select value={value} disabled={!floating} onChange={(e) => set(e.target.value)}>{opts.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>;
  const field = (value, set) => <span className="pd-input small"><input disabled={!floating} value={value} onChange={(e) => { setErr(''); set(e.target.value); }} /><em>см</em></span>;
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd layout-dlg" role="dialog" aria-label="Макет" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Макет</div>
        <div className="modal-body">
          <div className="pd-tabs" role="tablist">
            {[['position', 'Положение'], ['wrap', 'Обтекание текстом']].map(([k, l]) => <button key={k} type="button" role="tab" className={tab === k ? 'active' : ''} onClick={() => setTab(k)}>{l}</button>)}
          </div>
          {tab === 'position' && (
            <>
              {!floating && <div className="hf-hint">Рисунок в тексте: положение задаётся, когда выбрано обтекание.</div>}
              <fieldset><legend>По горизонтали</legend>
                <div className="lay-row"><label className="pd-check"><input type="radio" disabled={!floating} checked={hMode === 'align'} onChange={() => setHMode('align')} /> выравнивание</label>
                  {sel(h.align, (x) => setH({ ...h, align: x }), H_ALIGN)}<span>относительно</span>{sel(h.rel, (x) => setH({ ...h, rel: x }), H_REL)}</div>
                <div className="lay-row"><label className="pd-check"><input type="radio" disabled={!floating} checked={hMode === 'off'} onChange={() => setHMode('off')} /> положение</label>
                  {field(h.off, (x) => setH({ ...h, off: x }))}<span>правее</span>{sel(h.rel, (x) => setH({ ...h, rel: x }), H_REL)}</div>
              </fieldset>
              <fieldset><legend>По вертикали</legend>
                <div className="lay-row"><label className="pd-check"><input type="radio" disabled={!floating} checked={vMode === 'align'} onChange={() => setVMode('align')} /> выравнивание</label>
                  {sel(v.align, (x) => setV({ ...v, align: x }), V_ALIGN)}<span>относительно</span>{sel(v.rel, (x) => setV({ ...v, rel: x }), V_REL)}</div>
                <div className="lay-row"><label className="pd-check"><input type="radio" disabled={!floating} checked={vMode === 'off'} onChange={() => setVMode('off')} /> положение</label>
                  {field(v.off, (x) => setV({ ...v, off: x }))}<span>ниже</span>{sel(v.rel, (x) => setV({ ...v, rel: x }), V_REL)}</div>
              </fieldset>
            </>
          )}
          {tab === 'wrap' && (
            <>
              <fieldset><legend>Обтекание</legend>
                <div className="wrap-kinds">
                  {WRAP_KINDS.map(([k, l]) => (
                    <button key={k} type="button" className={`cols-preset${kind === k ? ' on' : ''}`} onClick={() => setKind(k)}><WrapIcon kind={k} /><span>{l}</span></button>
                  ))}
                </div>
              </fieldset>
              <fieldset><legend>Текст</legend>
                <div className="pd-checks ps-row">
                  {SIDES.map(([k, l]) => <label key={k} className="pd-check"><input type="radio" disabled={!['square', 'tight', 'through'].includes(kind)} checked={side === k} onChange={() => setSide(k)} /> {l}</label>)}
                </div>
              </fieldset>
              <fieldset><legend>Расстояние от текста</legend>
                <div className="pd-grid">
                  <label className="pd-field"><span>сверху</span>{field(dist.t, (x) => setDist({ ...dist, t: x }))}</label>
                  <label className="pd-field"><span>слева</span>{field(dist.l, (x) => setDist({ ...dist, l: x }))}</label>
                  <label className="pd-field"><span>снизу</span>{field(dist.b, (x) => setDist({ ...dist, b: x }))}</label>
                  <label className="pd-field"><span>справа</span>{field(dist.r, (x) => setDist({ ...dist, r: x }))}</label>
                </div>
              </fieldset>
              {kind === 'tight' || kind === 'through' ? <div className="hf-hint">«По контуру» и «Сквозное» Перо показывает как «Вокруг рамки»; в файле сохраняются как заданы.</div> : null}
            </>
          )}
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
