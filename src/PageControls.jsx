// Страница: вкладка «Макет» (поля, ориентация, размер), колонтитулы и номера страниц, CSS для печати/PDF.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { EMPTY_HF, MARGIN_PRESETS, PAPER, TWIPS_PER_CM, hfText, pagePx, parseSectPr, patchSectPr } from './editor/page.js';
import { fmtNum } from './editor/paragraph.js';

// Изменить атрибут документа одной транзакцией (с отменой по Ctrl+Z)
const setDocAttr = (editor, name, value) => editor.chain().focus().command(({ tr }) => { tr.setDocAttribute(name, value); return true; }).run();
const setPage = (editor, patch) => setDocAttr(editor, 'sectPr', patchSectPr(editor.state.doc.attrs.sectPr, patch));
const cm = (tw) => fmtNum(tw / TWIPS_PER_CM, 2);
const tw = (s) => {
  const v = parseFloat(String(s).replace(',', '.'));
  return isFinite(v) && v >= 0 && v <= 55 ? Math.round(v * TWIPS_PER_CM) : null;
};

function MenuButton({ icon, label, title, children }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title={title || label}
        onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
        {icon}<span className="rb-label">{label} ▾</span>
      </button>
      {pop.open && <div className="menu-pop" style={pop.style} role="menu">{children(() => pop.setOpen(false))}</div>}
    </div>
  );
}
const Item = ({ onClick, checked, children }) => (
  <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
    <span className="check">{checked ? '✓' : ''}</span>{children}
  </button>
);

const S = ({ children }) => (
  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" aria-hidden="true">{children}</svg>
);
const MarginsIcon = () => <S><rect x="5" y="3" width="14" height="18" /><rect x="8" y="6" width="8" height="12" strokeDasharray="2 1.5" /></S>;
const OrientIcon = () => <S><rect x="3" y="7" width="13" height="10" /><rect x="11" y="3" width="9" height="13" opacity=".5" /></S>;
const SizeIcon = () => <S><rect x="5" y="3" width="14" height="18" /><path d="M8 8h8M8 12h8M8 16h5" /></S>;
const HfIcon = () => <S><rect x="5" y="3" width="14" height="18" /><path d="M5 7h14M5 17h14" /><path d="M8 5h5M8 19h3" strokeWidth="1.2" /></S>;
const NumIcon = () => <S><rect x="5" y="3" width="14" height="18" /><text x="12" y="19" fontSize="6" textAnchor="middle" fill="currentColor" stroke="none">1</text></S>;

// ---------- Вкладка «Макет» ----------
export function LayoutPanel({ editor, Group, onMargins, onHf }) {
  const g = parseSectPr(editor.state.doc.attrs.sectPr);
  const same = (p) => Math.abs(p.top - g.top) < 15 && Math.abs(p.right - g.right) < 15 && Math.abs(p.bottom - g.bottom) < 15 && Math.abs(p.left - g.left) < 15;
  const paper = PAPER.find(([, w, h]) => (Math.abs(w - g.w) < 30 && Math.abs(h - g.h) < 30) || (Math.abs(h - g.w) < 30 && Math.abs(w - g.h) < 30));
  const hf = editor.state.doc.attrs.hf || EMPTY_HF;
  const pageNum = (where) => {
    const para = [{ text: '{PAGE}', align: where.endsWith('right') ? 'right' : 'center' }];
    const next = { ...EMPTY_HF, ...hf, own: { ...EMPTY_HF.own, ...(hf.own || {}) } };
    if (where.startsWith('top')) { next.header = para; next.own = { ...next.own, header: true }; }
    else { next.footer = para; next.own = { ...next.own, footer: true }; }
    setDocAttr(editor, 'hf', next);
  };
  return (
    <>
      <Group title="Параметры страницы">
        <MenuButton icon={<MarginsIcon />} label="Поля" title="Поля страницы">
          {(close) => (
            <>
              {MARGIN_PRESETS.map(([key, label, m]) => (
                <Item key={key} checked={same(m)} onClick={() => { close(); setPage(editor, m); }}>
                  {label}: {cm(m.top)} / {cm(m.right)} / {cm(m.bottom)} / {cm(m.left)} см
                </Item>
              ))}
              <div className="menu-sep" />
              <Item onClick={() => { close(); onMargins(); }}>Настраиваемые поля…</Item>
            </>
          )}
        </MenuButton>
        <MenuButton icon={<OrientIcon />} label="Ориентация" title="Ориентация страницы">
          {(close) => (
            <>
              <Item checked={g.orient !== 'landscape'} onClick={() => { close(); if (g.orient === 'landscape') setPage(editor, { w: g.h, h: g.w, orient: 'portrait' }); }}>Книжная</Item>
              <Item checked={g.orient === 'landscape'} onClick={() => { close(); if (g.orient !== 'landscape') setPage(editor, { w: g.h, h: g.w, orient: 'landscape' }); }}>Альбомная</Item>
            </>
          )}
        </MenuButton>
        <MenuButton icon={<SizeIcon />} label="Размер" title="Размер бумаги">
          {(close) => PAPER.map(([name, w, h]) => (
            <Item key={name} checked={paper && paper[0] === name} onClick={() => {
              close();
              const land = g.orient === 'landscape';
              setPage(editor, { w: land ? h : w, h: land ? w : h });
            }}>{name}: {cm(w)} × {cm(h)} см</Item>
          ))}
        </MenuButton>
      </Group>
      <Group title="Колонтитулы">
        <button type="button" className="rb-btn big" title="Верхний и нижний колонтитулы" onMouseDown={(e) => e.preventDefault()} onClick={onHf}>
          <HfIcon /><span className="rb-label">Колонтитулы</span>
        </button>
        <MenuButton icon={<NumIcon />} label="Номер страницы" title="Номера страниц">
          {(close) => (
            <>
              <Item onClick={() => { close(); pageNum('bottom-center'); }}>Внизу по центру</Item>
              <Item onClick={() => { close(); pageNum('bottom-right'); }}>Внизу справа</Item>
              <Item onClick={() => { close(); pageNum('top-right'); }}>Вверху справа</Item>
            </>
          )}
        </MenuButton>
      </Group>
    </>
  );
}

// ---------- Диалог «Поля» ----------
export function MarginsDialog({ editor, onClose }) {
  const g = parseSectPr(editor.state.doc.attrs.sectPr);
  const [f, setF] = useState({ top: cm(g.top), bottom: cm(g.bottom), left: cm(g.left), right: cm(g.right), header: cm(g.header), footer: cm(g.footer) });
  const [err, setErr] = useState('');
  const apply = () => {
    const out = {};
    for (const k of Object.keys(f)) {
      const v = tw(f[k]);
      if (v === null) { setErr('Поля — числа от 0 до 55 см'); return; }
      out[k] = v;
    }
    if (out.left + out.right > g.w - 1000 || out.top + out.bottom > g.h - 1000) { setErr('Поля слишком большие для этого размера бумаги'); return; }
    setPage(editor, out);
    onClose();
  };
  const field = (k, label) => (
    <label className="pd-field"><span>{label}</span>
      <span className="pd-input"><input value={f[k]} onChange={(e) => { setErr(''); setF({ ...f, [k]: e.target.value }); }} /><em>см</em></span>
    </label>
  );
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Поля" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Параметры страницы — поля</div>
        <div className="modal-body">
          <fieldset><legend>Поля</legend>
            <div className="pd-grid">{field('top', 'Верхнее')}{field('bottom', 'Нижнее')}{field('left', 'Левое')}{field('right', 'Правое')}</div>
          </fieldset>
          <fieldset><legend>Колонтитулы от края</legend>
            <div className="pd-grid">{field('header', 'Верхнего')}{field('footer', 'Нижнего')}</div>
          </fieldset>
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

// ---------- Диалог «Колонтитулы» ----------
const toText = (paras) => (paras || []).map((p) => p.text).join('\n');
const toParas = (text, align) => {
  const lines = String(text).split('\n');
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  return lines.length ? lines.map((t) => ({ text: t, align })) : null;
};

export function HeaderFooterDialog({ editor, onClose }) {
  const hf = editor.state.doc.attrs.hf || EMPTY_HF;
  const own = hf.own || {};
  const [head, setHead] = useState(toText(hf.header));
  const [foot, setFoot] = useState(toText(hf.footer));
  const [headAlign, setHeadAlign] = useState((hf.header && hf.header[0] && hf.header[0].align) || 'center');
  const [footAlign, setFootAlign] = useState((hf.footer && hf.footer[0] && hf.footer[0].align) || 'center');
  const fromFile = (!own.header && hf.header) || (!own.footer && hf.footer);
  const apply = () => {
    const next = { header: hf.header, footer: hf.footer, own: { header: !!own.header, footer: !!own.footer } };
    if (head !== toText(hf.header) || (hf.header && headAlign !== hf.header[0].align)) { next.header = toParas(head, headAlign); next.own.header = true; }
    if (foot !== toText(hf.footer) || (hf.footer && footAlign !== hf.footer[0].align)) { next.footer = toParas(foot, footAlign); next.own.footer = true; }
    setDocAttr(editor, 'hf', next);
    onClose();
  };
  const block = (label, value, set, align, setAlign, id) => (
    <fieldset>
      <legend>{label}</legend>
      <textarea id={id} className="hf-text" rows={2} value={value} onChange={(e) => set(e.target.value)} placeholder="Текст колонтитула (каждая строка — абзац)" />
      <div className="hf-tools">
        <select value={align} onChange={(e) => setAlign(e.target.value)} aria-label={`${label}: выравнивание`}>
          <option value="left">По левому краю</option><option value="center">По центру</option><option value="right">По правому краю</option>
        </select>
        <button type="button" className="btn small" onClick={() => set(`${value}{PAGE}`)}>+ номер страницы</button>
        <button type="button" className="btn small" onClick={() => set(`${value}{NUMPAGES}`)}>+ всего страниц</button>
      </div>
    </fieldset>
  );
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Колонтитулы" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Колонтитулы</div>
        <div className="modal-body">
          {block('Верхний колонтитул', head, setHead, headAlign, setHeadAlign, 'hf-head')}
          {block('Нижний колонтитул', foot, setFoot, footAlign, setFootAlign, 'hf-foot')}
          <div className="hf-hint">{'{PAGE}'} — номер страницы, {'{NUMPAGES}'} — всего страниц.</div>
          {fromFile && <div className="hf-hint warn">Колонтитул из файла при изменении заменится простым текстом: рисунки и оформление в нём не сохранятся. Неизменённый колонтитул сохраняется как был.</div>}
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Колонтитулы первой и последней страницы (остальные рисуют распорки разбиения) ----------
export function PageChrome({ geo, hf, total }) {
  const line = (p, i, n) => <div key={i} className={`hf-line hf-${p.align || 'left'}`}>{hfText(p.text, n, total) || ' '}</div>;
  return (
    <>
      <div className="pero-hf pero-header page-chrome-top" style={{ height: geo.top, paddingTop: geo.header, paddingLeft: geo.left, paddingRight: geo.right }}>
        {(hf.header || []).map((p, i) => line(p, i, 1))}
      </div>
      <div className="pero-hf pero-footer page-chrome-bottom" style={{ height: geo.bottom, paddingBottom: geo.footer, paddingLeft: geo.left, paddingRight: geo.right }}>
        {(hf.footer || []).map((p, i) => line(p, i, total))}
      </div>
    </>
  );
}

// ---------- CSS печати/PDF: размер листа, поля и колонтитулы в полях страницы (@page) ----------
const cssString = (s) => `"${String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
function marginBox(paras, page = 'counter(page)', pages = 'counter(pages)') {
  return paras.map((p) => String(p.text).split(/(\{PAGE\}|\{NUMPAGES\})/).filter(Boolean)
    .map((x) => (x === '{PAGE}' ? page : x === '{NUMPAGES}' ? pages : cssString(x))).join(' ')).join(' "\\A" ');
}
export function printCss(sectPr, hf) {
  const g = parseSectPr(sectPr);
  const px = pagePx(g);
  const boxes = [];
  for (const [kind, where] of [['header', 'top'], ['footer', 'bottom']]) {
    const paras = (hf && hf[kind]) || [];
    for (const align of ['left', 'center', 'right']) {
      const ps = paras.filter((p) => (p.align || 'left') === align && p.text);
      if (!ps.length) continue;
      const pad = kind === 'header' ? `padding-top: ${px.header}px; vertical-align: top;` : `padding-bottom: ${px.footer}px; vertical-align: bottom;`;
      boxes.push(`@${where}-${align} { content: ${marginBox(ps)}; white-space: pre; font: 10pt Calibri, Carlito, sans-serif; ${pad} }`);
    }
  }
  return `@media print { @page { size: ${px.width}px ${px.height}px; margin: ${px.top}px ${px.right}px ${px.bottom}px ${px.left}px; ${boxes.join(' ')} } }`;
}
