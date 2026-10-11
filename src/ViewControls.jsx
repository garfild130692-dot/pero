// Вкладка «Вид» документа: режимы (разметка страницы, черновик, чтение), показ (линейка, область навигации,
// непечатаемые знаки), масштаб (100 %, по ширине, страница целиком, диалог), окна (упорядочить, перейти) —
// и сами область навигации и диалог «Масштаб».
import React, { useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { TextSelection } from '@tiptap/pm/state';
import { usePopover } from './FontControls.jsx';
import { levelLabel } from './editor/numbering.js';
import * as fileApi from './fileApi.js';

export const ZOOM_MIN = 10;
export const ZOOM_MAX = 500;
export const clampZoom = (z) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z)));

// Масштаб «по ширине страницы» / «страница целиком» для области просмотра avail { w, h } и листа page { width, height }
export function fitZoom(kind, avail, page) {
  const byW = ((avail.w - 48) / page.width) * 100;
  if (kind === 'width') return clampZoom(byW);
  const byH = ((avail.h - 32) / page.height) * 100;
  return clampZoom(Math.min(byW, byH));
}

// Настройки вида — у окна, запоминаются (линейка, область навигации)
export function loadViewPrefs() {
  try { return { ruler: true, nav: false, ...JSON.parse(localStorage.getItem('pero.view') || '{}') }; } catch { return { ruler: true, nav: false }; }
}
export function saveViewPrefs(p) { try { localStorage.setItem('pero.view', JSON.stringify(p)); } catch { /* не запоминаем */ } }

function WindowsButton({ Btn }) {
  const pop = usePopover();
  const [list, setList] = useState([]);
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big icon="windows" label="Перейти в окно" title="Перейти в другое окно Перо" active={pop.open}
        onClick={async () => { if (!pop.open) setList(await fileApi.winList()); pop.setOpen(!pop.open); }} />
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {list.map((w) => (
            <button key={w.id} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()}
              onClick={() => { pop.setOpen(false); fileApi.winFocus(w.id); }}>
              <span className="check">{w.current ? '✓' : ''}</span>{w.title}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

export function ViewPanel({ Group, Btn, editor, view, showMarks }) {
  const v = view;
  return (
    <>
      <Group title="Режимы">
        <Btn big icon="readMode" label="Режим чтения" title="Режим чтения: только текст на весь экран (Esc — выход)" active={v.mode === 'read'} onClick={() => v.setMode('read')} />
        <Btn big icon="printLayout" label="Разметка страницы" title="Разметка страницы: листы, поля, колонтитулы (Ctrl+Alt+P)" active={v.mode === 'print'} onClick={() => v.setMode('print')} />
        <Btn big icon="draft" label="Черновик" title="Черновик: сплошной текст без листов и колонтитулов (Ctrl+Alt+N)" active={v.mode === 'draft'} onClick={() => v.setMode('draft')} />
      </Group>
      <Group title="Показать">
        <div className="rb-rows">
          <label className="rb-check"><input type="checkbox" checked={v.ruler} onChange={v.toggleRuler} /> Линейка</label>
          <label className="rb-check"><input type="checkbox" checked={v.nav} onChange={v.toggleNav} /> Область навигации</label>
          <label className="rb-check"><input type="checkbox" checked={showMarks} onChange={() => editor.chain().focus().toggleFormattingMarks().run()} /> Непечатаемые знаки</label>
        </div>
      </Group>
      <Group title="Масштаб">
        <Btn big icon="zoom" label="Масштаб" title="Масштаб документа…" onClick={v.openZoom} />
        <Btn big icon="zoom100" label="100%" title="Масштаб 100 %" onClick={() => v.setZoom(100)} />
        <div className="rb-rows">
          <Btn title="Страница целиком в окне" onClick={() => v.fit('page')}><span className="rb-inline">Одна страница</span></Btn>
          <Btn title="Страница по ширине окна" onClick={() => v.fit('width')}><span className="rb-inline">По ширине страницы</span></Btn>
        </div>
      </Group>
      <Group title="Окно">
        <Btn big icon="newWindowDoc" label="Новое окно" title="Новый документ в отдельном окне (Ctrl+N)" onClick={() => fileApi.newWindow('doc')} />
        <Btn big icon="arrange" label="Упорядочить все" title="Расположить все окна Перо рядом" onClick={() => fileApi.winArrange()} />
        <WindowsButton Btn={Btn} />
      </Group>
    </>
  );
}

// «Масштаб»: готовые значения, по ширине, страница целиком, произвольный (10–500 %)
export function ZoomDialog({ zoom, onFit, onApply, onClose }) {
  const [val, setVal] = useState(String(zoom));
  const PRESETS = [200, 100, 75];
  const n = parseInt(val, 10);
  const ok = n >= ZOOM_MIN && n <= ZOOM_MAX;
  const apply = () => { if (ok) { onApply(n); onClose(); } };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd zoom-dlg" role="dialog" aria-label="Масштаб" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Масштаб</div>
        <div className="modal-body">
          <div className="zoom-grid">
            {PRESETS.map((p) => (
              <label key={p} className="pd-check"><input type="radio" checked={n === p} onChange={() => setVal(String(p))} /> {p}%</label>
            ))}
            <label className="pd-check"><input type="radio" checked={false} onChange={() => { onFit('width'); onClose(); }} /> По ширине страницы</label>
            <label className="pd-check"><input type="radio" checked={false} onChange={() => { onFit('page'); onClose(); }} /> Страница целиком</label>
          </div>
          <label className="pd-field"><span>Произвольный</span>
            <span className="pd-input small"><input autoFocus type="number" min={ZOOM_MIN} max={ZOOM_MAX} value={val} onChange={(e) => setVal(e.target.value)} /><em>%</em></span>
          </label>
          {!ok && <div className="pd-error">Масштаб — от {ZOOM_MIN} до {ZOOM_MAX} %</div>}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" disabled={!ok} onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// Заголовки документа с номерами (нумерация заголовков) и текущий — для области навигации
export function headingsOutline(state) {
  const out = [];
  const hn = state.doc.attrs.headingNumbering;
  const numbered = hn && !hn.off && hn.levels;
  const nums = Array(9).fill(0);
  const at = state.selection.from;
  let current = -1;
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'heading') return node.isBlock && node.type.name !== 'paragraph';
    const lvl = node.attrs.level;
    let label = '';
    if (numbered && !node.attrs.noNumber && !node.attrs.fileStyle) {
      for (let i = 0; i < 9; i++) {
        if (i === lvl - 1) nums[i] = nums[i] ? nums[i] + 1 : (hn.levels[i].start || 1);
        else if (i >= lvl) nums[i] = 0;
      }
      const filled = nums.map((x, i) => x || hn.levels[i].start || 1);
      if (hn.levels[lvl - 1] && hn.levels[lvl - 1].on) label = levelLabel(hn.levels, lvl - 1, filled);
    }
    const text = node.textBetween(0, node.content.size, ' ', '').replace(/\s+/g, ' ').trim();
    if (pos <= at) current = out.length;
    out.push({ pos, level: lvl, text: text || '(пустой заголовок)', label });
    return false;
  });
  return { list: out, current };
}

// Область навигации: заголовки документа (как в Word), фильтр, щелчок — переход
export function NavPane({ editor, onClose }) {
  const [q, setQ] = useState('');
  const { list, current } = useEditorState({
    editor,
    selector: ({ editor: e }) => headingsOutline(e.state),
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  const go = (h) => {
    const { state, view } = editor;
    view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, h.pos + 1)).scrollIntoView());
    view.focus();
    // заголовок — вверху области просмотра, как в Word
    const dom = view.nodeDOM(h.pos);
    if (dom && dom.scrollIntoView) dom.scrollIntoView({ block: 'start' });
  };
  const ql = q.trim().toLowerCase();
  const shown = list.map((h, i) => ({ ...h, i })).filter((h) => !ql || h.text.toLowerCase().includes(ql));
  return (
    <aside className="nav-pane" aria-label="Навигация">
      <div className="cp-head">
        <span>Навигация</span>
        <button type="button" className="cc-del" title="Скрыть область навигации" onClick={onClose}>✕</button>
      </div>
      <input className="start-search nav-search" type="search" placeholder="Поиск в заголовках" value={q} onChange={(e) => setQ(e.target.value)} />
      <div className="nav-list" role="tree">
        {!list.length && <div className="cp-empty">В документе нет заголовков. Примените стиль «Заголовок 1–6» (группа «Стили»), чтобы видеть структуру здесь.</div>}
        {shown.map((h) => (
          <button key={h.pos} type="button" role="treeitem" className={`nav-item${h.i === current ? ' current' : ''}`}
            style={{ paddingLeft: 10 + (h.level - 1) * 14 }} title={h.text} onClick={() => go(h)}>
            {h.label && <span className="nav-num">{h.label}</span>}{h.text}
          </button>
        ))}
      </div>
    </aside>
  );
}
