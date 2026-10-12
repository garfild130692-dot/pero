// Вкладка «Конструктор»: шрифт и интервал документа, подложка, цвет страницы, границы страниц; показ подложки и рамки
// на листах (экран — по листам разбиения, печать — position: fixed повторяется на каждой странице).
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { ColorPalette, usePopover } from './FontControls.jsx';
import { cssFontStack } from './editor/font.js';
import {
  BORDER_CSS, DOC_FONTS, PAGE_BORDER_STYLES, SPACING_PRESETS, WATERMARK_DEFAULT, WATERMARK_PRESETS, pageBorderRect, parsePgBorders, watermarkBox,
} from './editor/design.js';
import { parseSectPr, sectionFrame, sectionList } from './editor/page.js';
import { pageInfo, pageSects } from './editor/pagination.js';
import { pageTops } from './VRuler.jsx';

const PX = 4 / 3; // px в pt

function Menu({ Btn, icon, label, title, children }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big icon={icon} label={label} title={title} active={pop.open} onClick={() => pop.setOpen(!pop.open)} />
      {pop.open && <div className="menu-pop design-pop" style={pop.style} role="menu">{children(() => pop.setOpen(false))}</div>}
    </div>
  );
}
const Item = ({ close, onClick, checked, children, title, style }) => (
  <button type="button" className="menu-item" title={title} style={style} onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); onClick(); }}>
    <span className="check">{checked ? '✓' : ''}</span>{children}
  </button>
);

const samePara = (a, b) => ['spaceAfter', 'lineSpacing'].every((k) => Math.abs((a[k] ?? 0) - (b[k] ?? 0)) < 0.01) && Math.abs((a.indentFirst || 0) - (b.indentFirst || 0)) < 0.5;

export function DesignPanel({ Group, Btn, editor, onDialog }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      font: e.state.doc.attrs.fontFamily, para: e.state.doc.attrs.para || {},
      wm: e.state.doc.attrs.watermark, color: e.state.doc.attrs.pageColor,
    }),
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  const c = () => editor.chain().focus();
  return (
    <>
      <Group title="Форматирование документа">
        <Menu Btn={Btn} icon="fonts" label="Шрифт документа" title="Шрифт основного текста документа (стиль «Обычный»)">
          {(close) => DOC_FONTS.map((f) => (
            <Item key={f} close={close} checked={s.font === f} style={{ fontFamily: cssFontStack(f) }} onClick={() => c().setDocFont(f).run()}>{f}</Item>
          ))}
        </Menu>
        <Menu Btn={Btn} icon="paraSpacing" label="Интервал между абзацами" title="Интервалы и межстрочный интервал документа по умолчанию">
          {(close) => SPACING_PRESETS.map(([k, label, p]) => (
            <Item key={k} close={close} checked={samePara(s.para, p)} title={`После абзаца ${p.spaceAfter} пт, множитель ${String(p.lineSpacing).replace('.', ',')}`}
              onClick={() => c().setParaSpacing(k).run()}>{label}</Item>
          ))}
        </Menu>
      </Group>
      <Group title="Фон страницы">
        <Menu Btn={Btn} icon="watermark" label="Подложка" title="Надпись за текстом на каждой странице">
          {(close) => (
            <>
              <div className="pop-title">Подложка по диагонали</div>
              {WATERMARK_PRESETS.map((t) => (
                <Item key={t} close={close} checked={!!s.wm && s.wm.text === t} onClick={() => c().setWatermark({ ...WATERMARK_DEFAULT, ...(s.wm || {}), text: t }).run()}>{t}</Item>
              ))}
              <div className="menu-sep" />
              <Item close={close} onClick={() => onDialog('watermark')}>Настраиваемая подложка…</Item>
              <Item close={close} onClick={() => c().setWatermark(null).run()}>Удалить подложку</Item>
            </>
          )}
        </Menu>
        <Menu Btn={Btn} icon="pageColor" label="Цвет страницы" title="Цвет фона страницы (на экране; при печати не выводится, как в Word)">
          {(close) => (
            <ColorPalette autoLabel="Нет цвета" autoClass="none" current={s.color || '#ffffff'}
              onAuto={() => { close(); c().setPageColor(null).run(); }} onPick={(x) => { close(); c().setPageColor(x).run(); }} />
          )}
        </Menu>
        <Btn big icon="pageBorders" label="Границы страниц" title="Рамка вокруг каждой страницы…" onClick={() => onDialog('pageBorders')} />
      </Group>
    </>
  );
}

// «Подложка»: нет / текст (готовый или свой), шрифт, размер, цвет, полупрозрачная, по диагонали / горизонтально
export function WatermarkDialog({ editor, onClose }) {
  const cur = editor.state.doc.attrs.watermark;
  const [on, setOn] = useState(cur ? 'text' : 'none');
  const [wm, setWm] = useState({ ...WATERMARK_DEFAULT, ...(cur || {}) });
  const set = (k) => (v) => setWm((w) => ({ ...w, [k]: v }));
  const apply = (close = true) => {
    editor.chain().focus().setWatermark(on === 'none' || !wm.text.trim() ? null : { ...wm, text: wm.text.trim() }).run();
    if (close) onClose();
  };
  const dis = on === 'none';
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd wm-dlg" role="dialog" aria-label="Подложка" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Подложка</div>
        <div className="modal-body">
          <label className="pd-check"><input type="radio" checked={on === 'none'} onChange={() => setOn('none')} /> Нет подложки</label>
          <label className="pd-check"><input type="radio" checked={on === 'text'} onChange={() => setOn('text')} /> Текст</label>
          <div className="wm-grid">
            <label className="fd-col"><span>Текст</span>
              <input list="wm-texts" value={wm.text} disabled={dis} onChange={(e) => set('text')(e.target.value)} />
              <datalist id="wm-texts">{WATERMARK_PRESETS.map((t) => <option key={t} value={t} />)}</datalist>
            </label>
            <label className="fd-col"><span>Шрифт</span>
              <select value={wm.font} disabled={dis} onChange={(e) => set('font')(e.target.value)}>
                {[...new Set([wm.font, ...DOC_FONTS])].map((f) => <option key={f} value={f}>{f}</option>)}
              </select>
            </label>
            <label className="fd-col"><span>Размер</span>
              <select value={wm.size || ''} disabled={dis} onChange={(e) => set('size')(e.target.value ? +e.target.value : null)}>
                <option value="">Авто</option>
                {[36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144].map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
            </label>
            <label className="fd-col"><span>Цвет</span>
              <span className="fd-color">
                <input type="color" value={wm.color} disabled={dis} onChange={(e) => set('color')(e.target.value)} />
                <label className="pd-check"><input type="checkbox" checked={wm.semi} disabled={dis} onChange={(e) => set('semi')(e.target.checked)} /> Полупрозрачный</label>
              </span>
            </label>
          </div>
          <div className="wm-layout">
            <span>Расположение:</span>
            <label className="pd-check"><input type="radio" disabled={dis} checked={wm.layout === 'diagonal'} onChange={() => set('layout')('diagonal')} /> По диагонали</label>
            <label className="pd-check"><input type="radio" disabled={dis} checked={wm.layout === 'horizontal'} onChange={() => set('layout')('horizontal')} /> Горизонтально</label>
          </div>
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn" onClick={() => apply(false)}>Применить</button>
          <button className="btn primary" onClick={() => apply()}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// «Границы страниц»: нет / рамка; тип линии, ширина, цвет; отступ от края страницы или от текста
export function PageBordersDialog({ editor, onClose }) {
  const cur = parsePgBorders(editor.state.doc.attrs.sectPr);
  const [on, setOn] = useState(!!cur);
  const [b, setB] = useState(cur || { offsetFrom: 'page', style: 'single', width: 0.5, color: null, space: 24 });
  const set = (k) => (v) => setB((x) => ({ ...x, [k]: v }));
  const apply = () => { editor.chain().focus().setPageBorders(on ? b : null).run(); onClose(); };
  const fromPage = (v) => setB((x) => ({ ...x, offsetFrom: v, space: v === 'page' ? 24 : 4 }));
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd pb-dlg" role="dialog" aria-label="Границы страниц" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Границы страниц</div>
        <div className="modal-body pb-body">
          <div className="pb-types">
            <label className="pd-check"><input type="radio" checked={!on} onChange={() => setOn(false)} /> Нет</label>
            <label className="pd-check"><input type="radio" checked={on} onChange={() => setOn(true)} /> Рамка</label>
          </div>
          <div className="pb-opts">
            <label className="fd-col"><span>Тип</span>
              <select value={b.style} disabled={!on} onChange={(e) => set('style')(e.target.value)}>
                {PAGE_BORDER_STYLES.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
              </select>
            </label>
            <label className="fd-col"><span>Ширина</span>
              <select value={b.width} disabled={!on} onChange={(e) => set('width')(+e.target.value)}>
                {[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => <option key={w} value={w}>{String(w).replace('.', ',')} пт</option>)}
              </select>
            </label>
            <label className="fd-col"><span>Цвет</span>
              <span className="fd-color">
                <input type="color" value={b.color || '#000000'} disabled={!on} onChange={(e) => set('color')(e.target.value)} />
                <button type="button" className={`btn small${b.color ? '' : ' on'}`} disabled={!on} onClick={() => set('color')(null)}>Авто</button>
              </span>
            </label>
            <label className="fd-col"><span>Отсчитывать от</span>
              <select value={b.offsetFrom} disabled={!on} onChange={(e) => fromPage(e.target.value)}>
                <option value="page">края страницы</option>
                <option value="text">текста</option>
              </select>
            </label>
            <label className="fd-col"><span>Отступ, пт</span>
              <input type="number" min="0" max="31" value={b.space} disabled={!on} onChange={(e) => set('space')(Math.max(0, Math.min(31, +e.target.value || 0)))} />
            </label>
          </div>
          <div className="pb-preview">
            <div className="pb-sheet" style={on ? { outline: `${Math.max(1, b.width)}px ${BORDER_CSS[b.style] || 'solid'} ${b.color || '#000'}`, outlineOffset: b.offsetFrom === 'page' ? '-8px' : '-22px' } : undefined}>
              {Array.from({ length: 7 }, (_, i) => <div key={i} className="pb-line" />)}
            </div>
          </div>
          <div className="pd-note">Применяется ко всему документу (всем разделам).</div>
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// Надпись подложки (SVG), растянутая в рамку w×h px, как textpath Word с fitshape
function WatermarkSvg({ wm, w, h }) {
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} aria-hidden="true" style={{ overflow: 'visible' }}>
      <text x={w / 2} y={h * 0.8} textAnchor="middle" textLength={w} lengthAdjust="spacingAndGlyphs"
        style={{ fontFamily: cssFontStack(wm.font), fontSize: h * 0.95, fill: wm.color, fillOpacity: wm.semi ? 0.5 : 1 }}>{wm.text}</text>
    </svg>
  );
}

// Подложка и рамка на листах на экране: подложка — под текстом, рамка — поверх полос между листами
export function PageDecor({ editor }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      total: pageInfo(e.state).total, sects: pageSects(e.state), list: sectionList(e.state.doc),
      wm: e.state.doc.attrs.watermark,
    }),
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  const ref = useRef(null);
  const [tops, setTops] = useState([0]);
  // листы меряются, только если есть что рисовать (подложка или рамка) — на 200+ листах это заметно
  const need = !!s.wm || s.list.some((x) => parsePgBorders(x));
  const place = useCallback(() => {
    const page = need && ref.current && ref.current.closest('.page');
    if (page) setTops(pageTops(page).slice(0, s.total));
  }, [s.total, need]);
  useLayoutEffect(place, [place, s.list]);
  useEffect(() => {
    document.addEventListener('pero-paginated', place);
    return () => document.removeEventListener('pero-paginated', place);
  }, [place]);
  const frame = sectionFrame(s.list);
  const borders = s.list.map((x) => parsePgBorders(x));
  if (!s.wm && !borders.some(Boolean)) return <span ref={ref} hidden />;
  const pages = tops.map((top, i) => {
    const si = Math.min(s.sects[i] ?? s.sects[s.sects.length - 1] ?? 0, s.list.length - 1);
    const g = parseSectPr(s.list[si]);
    const off = (frame.sections[si] || {}).offset || 0;
    return { top, g, off, b: borders[si] };
  });
  return (
    <>
      <div className="page-decor under" ref={ref} aria-hidden="true">
        {s.wm && pages.map(({ top, g, off }, i) => {
          const tw = (g.w - g.left - g.right) / 20;
          const box = watermarkBox(s.wm, tw);
          const cx = off + (g.left / 20 + tw / 2) * PX;
          const cy = top + (g.top / 20 + (g.h - g.top - g.bottom) / 40) * PX;
          return (
            <div key={i} className="wm" style={{ left: cx - (box.w * PX) / 2, top: cy - (box.h * PX) / 2, transform: s.wm.layout === 'diagonal' ? 'rotate(-45deg)' : undefined }}>
              <WatermarkSvg wm={s.wm} w={box.w * PX} h={box.h * PX} />
            </div>
          );
        })}
      </div>
      <div className="page-decor over" aria-hidden="true">
        {pages.map(({ top, g, off, b }, i) => {
          if (!b) return null;
          const r = pageBorderRect(b, g);
          return (
            <div key={i} className="pgb" style={{
              left: off + r.left * PX, top: top + r.top * PX, width: (r.right - r.left) * PX, height: (r.bottom - r.top) * PX,
              border: `${b.style === 'double' ? Math.max(3, b.width * PX) : Math.max(1, b.width * PX)}px ${BORDER_CSS[b.style] || 'solid'} ${b.color || '#000'}`,
            }} />
          );
        })}
      </div>
    </>
  );
}

// Печать/PDF: подложка и рамка последнего раздела — fixed-блоки (Chromium повторяет их на каждой странице)
export function PrintDecor({ editor }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({ wm: e.state.doc.attrs.watermark, sectPr: e.state.doc.attrs.sectPr }),
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  const g = parseSectPr(s.sectPr);
  const b = parsePgBorders(s.sectPr);
  if (!s.wm && !b) return null;
  const tw = (g.w - g.left - g.right) / 20;
  const th = (g.h - g.top - g.bottom) / 20;
  const box = s.wm ? watermarkBox(s.wm, tw) : null;
  const r = b ? pageBorderRect(b, g) : null;
  // координаты fixed при печати — от области текста страницы (внутри полей @page)
  return (
    <div className="print-decor">
      {s.wm && (
        <div className="wm" style={{ left: `${tw / 2 - box.w / 2}pt`, top: `${th / 2 - box.h / 2}pt`, transform: s.wm.layout === 'diagonal' ? 'rotate(-45deg)' : undefined }}>
          <WatermarkSvg wm={s.wm} w={box.w * PX} h={box.h * PX} />
        </div>
      )}
      {r && (
        <div className="pgb" style={{
          left: `${r.left - g.left / 20}pt`, top: `${r.top - g.top / 20}pt`, width: `${r.right - r.left}pt`, height: `${r.bottom - r.top}pt`,
          border: `${Math.max(0.5, b.width)}pt ${BORDER_CSS[b.style] || 'solid'} ${b.color || '#000'}`,
        }} />
      )}
    </div>
  );
}
