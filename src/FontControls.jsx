// Группа «Шрифт» на ленте: гарнитура, размер, цвет текста, цвет выделения.
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { COMMON_FONTS, FONT_SIZES, HIGHLIGHTS, cssFontStack, parseSize } from './editor/font.js';
import { effectiveFont } from './editor/extensions.js';

// ---------- Текущее состояние шрифта для ленты ----------
// Пустая строка — в выделении разные значения (Word в этом случае показывает пустое поле).
export function fontState(editor) {
  const { state } = editor;
  const { selection, doc } = state;
  const ts = state.schema.marks.textStyle;
  const hl = state.schema.marks.highlight;
  const fam = new Set();
  const size = new Set();
  const color = new Set();
  const take = (marks, parent) => {
    const t = ts.isInSet(marks);
    const eff = effectiveFont(t && t.attrs, parent, doc.attrs);
    fam.add(eff.fontFamily);
    size.add(eff.fontSize);
    color.add((t && t.attrs.color) || '');
  };
  if (!selection.empty) {
    let n = 0;
    doc.nodesBetween(selection.from, selection.to, (node, pos, parent) => {
      if (node.isText && n++ < 2000) take(node.marks, parent);
    });
  }
  if (!fam.size) take(state.storedMarks || selection.$from.marks(), selection.$from.parent);
  const one = (s) => (s.size === 1 ? [...s][0] : '');
  const h = hl.isInSet(state.storedMarks || selection.$from.marks());
  return {
    fontFamily: one(fam),
    fontSize: one(size),
    color: one(color),
    highlight: h ? h.attrs.color : '',
    superscript: editor.isActive('superscript'),
    subscript: editor.isActive('subscript'),
  };
}

// ---------- Всплывающая панель, закрывается кликом мимо, Esc и изменением размера окна ----------
// Панель позиционируется fixed от якоря: лента прокручивается по горизонтали (overflow), и absolute-панель
// обрезалась бы её границами. style — готовые координаты для панели.
export function usePopover() {
  const [open, setOpenState] = useState(false);
  const [style, setStyle] = useState(null);
  const ref = useRef(null);
  const setOpen = (v) => {
    if (v && ref.current) {
      const r = ref.current.getBoundingClientRect();
      setStyle({ position: 'fixed', top: r.bottom + 2, left: Math.max(4, Math.min(r.left, window.innerWidth - 260)) });
    }
    setOpenState(v);
  };
  useEffect(() => {
    if (!open) return undefined;
    const onDown = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpenState(false); };
    const onKey = (e) => { if (e.key === 'Escape') setOpenState(false); };
    const onResize = () => setOpenState(false);
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    window.addEventListener('resize', onResize);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
      window.removeEventListener('resize', onResize);
    };
  }, [open]);
  return { open, setOpen, ref, style };
}

// ---------- Список шрифтов: установленные в системе (Local Font Access), иначе — распространённые ----------
let systemFonts = null;
async function loadSystemFonts() {
  if (systemFonts) return systemFonts;
  try {
    if (typeof window.queryLocalFonts === 'function') {
      const list = await window.queryLocalFonts();
      const names = [...new Set(list.map((f) => f.family))].filter(Boolean);
      if (names.length) systemFonts = names.sort((a, b) => a.localeCompare(b, 'ru'));
    }
  } catch { /* нет доступа — останется запасной список */ }
  if (!systemFonts) systemFonts = [...COMMON_FONTS];
  return systemFonts;
}

// Шрифты, которые уже встречаются в документе, — показываем всегда, даже если их нет в системе
function documentFonts(editor) {
  const out = new Set([editor.state.doc.attrs.fontFamily]);
  editor.state.doc.descendants((node) => {
    for (const m of node.marks || []) if (m.type.name === 'textStyle' && m.attrs.fontFamily) out.add(m.attrs.fontFamily);
  });
  return [...out].filter(Boolean);
}

// Поле со списком: можно выбрать из списка или ввести своё значение и нажать Enter
function ComboBox({ value, title, width, items, renderItem, onOpen, onCommit }) {
  const { open, setOpen, ref, style } = usePopover();
  const [text, setText] = useState(value);
  const [filter, setFilter] = useState('');
  const listRef = useRef(null);
  useEffect(() => { setText(value); }, [value]);

  const shown = filter ? items.filter((i) => String(i).toLowerCase().includes(filter.toLowerCase())) : items;
  useEffect(() => {
    // при открытии прокручиваем к текущему значению
    if (!open || !listRef.current) return;
    const el = listRef.current.querySelector('.current');
    if (el) el.scrollIntoView({ block: 'center' });
  }, [open]);

  // после применения поле показывает действующее значение (неверный ввод просто откатывается)
  const commit = (v) => { setOpen(false); setFilter(''); setText(value); onCommit(v); };
  const toggle = () => { if (!open) onOpen && onOpen(); setFilter(''); setOpen(!open); };
  return (
    <div className="combo" ref={ref} style={{ width }} title={title}>
      <input
        value={text}
        onChange={(e) => { setText(e.target.value); setFilter(e.target.value); if (!open) { onOpen && onOpen(); setOpen(true); } }}
        onFocus={(e) => e.target.select()}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); commit(text); }
          else if (e.key === 'Escape') { setText(value); setOpen(false); }
        }}
        onBlur={() => { if (!open) setText(value); }}
        aria-label={title}
      />
      <button type="button" className="combo-arrow" tabIndex={-1} onMouseDown={(e) => e.preventDefault()} onClick={toggle} aria-label={title + ': список'}>▾</button>
      {open && (
        <div className="combo-list" ref={listRef} role="listbox" style={{ ...style, minWidth: width }}>
          {shown.map((item) => (
            <div
              key={item}
              role="option"
              aria-selected={String(item) === String(value)}
              className={`combo-item${String(item) === String(value) ? ' current' : ''}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => commit(String(item))}
            >
              {renderItem ? renderItem(item) : item}
            </div>
          ))}
          {!shown.length && <div className="combo-empty">Нажмите Enter, чтобы применить «{filter}»</div>}
        </div>
      )}
    </div>
  );
}

// Название шрифта, нарисованное им самим. Гарнитуру применяем, только когда строка видна в списке:
// иначе при первом открытии браузер грузит файлы сразу всех (~250) шрифтов — это секунды задержки.
function FontPreview({ name }) {
  const ref = useRef(null);
  const [seen, setSeen] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el || typeof IntersectionObserver === 'undefined') { setSeen(true); return undefined; }
    const io = new IntersectionObserver(([e]) => { if (e.isIntersecting) { setSeen(true); io.disconnect(); } },
      { root: el.closest('.combo-list'), rootMargin: '120px 0px' });
    io.observe(el);
    return () => io.disconnect();
  }, []);
  return <span ref={ref} style={seen ? { fontFamily: cssFontStack(name) } : undefined}>{name}</span>;
}

export function FontFamilyBox({ editor, value }) {
  const [fonts, setFonts] = useState(systemFonts || COMMON_FONTS);
  const items = useMemo(() => [...new Set([...documentFonts(editor), ...fonts])].sort((a, b) => a.localeCompare(b, 'ru')),
    // список шрифтов документа пересчитываем при открытии списка — см. onOpen
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [fonts, value]);
  return (
    <ComboBox
      value={value}
      title="Шрифт"
      width={150}
      items={items}
      onOpen={() => loadSystemFonts().then(setFonts)}
      renderItem={(f) => <FontPreview name={f} />}
      onCommit={(v) => {
        const name = v.trim();
        if (name) editor.chain().setFontFamily(name).run();
        editor.view.focus(); // синхронно: focus() в TipTap откладывает до следующего кадра
      }}
    />
  );
}

export function FontSizeBox({ editor, value }) {
  const shown = value === '' ? '' : String(value).replace('.', ',');
  return (
    <ComboBox
      value={shown}
      title="Размер шрифта"
      width={56}
      items={FONT_SIZES.map((s) => String(s).replace('.', ','))}
      onCommit={(v) => {
        const n = parseSize(v);
        if (n) editor.chain().setFontSize(n).run(); // неверное значение — просто возвращаемся к тексту
        editor.view.focus();
      }}
    />
  );
}

// ---------- Цвета ----------
const mix = (hex, to, k) => '#' + [1, 3, 5].map((i) => {
  const a = parseInt(hex.substr(i, 2), 16);
  const b = parseInt(to.substr(i, 2), 16);
  return Math.round(a + (b - a) * k).toString(16).padStart(2, '0');
}).join('');
// Свой набор основных цветов: 10 оттенков × (основной, 3 светлее, 2 темнее)
const BASE = ['#000000', '#ffffff', '#44546a', '#1f6f5c', '#2e75b6', '#c55a11', '#7f7f7f', '#bf9000', '#5b9bd5', '#70ad47'];
const PALETTE = [
  BASE,
  BASE.map((c, i) => (i === 0 ? '#7f7f7f' : i === 1 ? '#f2f2f2' : mix(c, '#ffffff', 0.8))),
  BASE.map((c, i) => (i === 0 ? '#595959' : i === 1 ? '#d9d9d9' : mix(c, '#ffffff', 0.6))),
  BASE.map((c, i) => (i === 0 ? '#3f3f3f' : i === 1 ? '#bfbfbf' : mix(c, '#ffffff', 0.4))),
  BASE.map((c, i) => (i === 0 ? '#262626' : i === 1 ? '#a6a6a6' : mix(c, '#000000', 0.25))),
  BASE.map((c, i) => (i === 0 ? '#0d0d0d' : i === 1 ? '#808080' : mix(c, '#000000', 0.5))),
];
const STANDARD = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
// Порядок цветов выделения — как привычно: сначала яркие, затем тёмные
const HIGHLIGHT_ORDER = ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen',
  'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black'];
const HIGHLIGHT_NAMES = {
  yellow: 'Жёлтый', green: 'Ярко-зелёный', cyan: 'Бирюзовый', magenta: 'Розовый', blue: 'Синий', red: 'Красный',
  darkBlue: 'Тёмно-синий', darkCyan: 'Сине-зелёный', darkGreen: 'Зелёный', darkMagenta: 'Фиолетовый',
  darkRed: 'Тёмно-красный', darkYellow: 'Коричнево-жёлтый', darkGray: 'Серый 50%', lightGray: 'Серый 25%', black: 'Чёрный',
};

const Swatch = ({ color, title, onPick }) => (
  <button type="button" className="swatch" title={title || color} style={{ background: color }}
    onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(color)} />
);

// Кнопка с разделённым действием: основная часть применяет последний цвет, стрелка открывает палитру
function SplitColor({ title, icon, color, children }) {
  const { open, setOpen, ref, style } = usePopover();
  return (
    <div className="split" ref={ref}>
      <button type="button" className="rb-btn split-main" title={title} onMouseDown={(e) => e.preventDefault()} onClick={() => children.apply()}>
        <span className="color-glyph">{icon}<span className="color-bar" style={{ background: color || 'transparent' }} /></span>
      </button>
      <button type="button" className="rb-btn split-arrow" title={title + ': выбрать цвет'} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>▾</button>
      {open && <div className="color-pop" style={style}>{children.palette(() => setOpen(false))}</div>}
    </div>
  );
}

export function TextColorButton({ editor }) {
  const [last, setLast] = useState('#c00000');
  const customRef = useRef(null);
  const apply = (c) => { setLast(c); editor.chain().focus().setColor(c).run(); };
  return (
    <SplitColor title="Цвет текста" color={last} icon={<span className="glyph-a">A</span>}>
      {{
        apply: () => apply(last),
        palette: (close) => (
          <>
            <button type="button" className="pop-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); editor.chain().focus().unsetColor().run(); }}>
              <span className="swatch auto" /> Авто
            </button>
            <div className="pop-title">Основные цвета</div>
            <div className="swatches">{PALETTE.flat().map((c, i) => <Swatch key={i} color={c} onPick={(x) => { close(); apply(x); }} />)}</div>
            <div className="pop-title">Стандартные цвета</div>
            <div className="swatches">{STANDARD.map((c) => <Swatch key={c} color={c} onPick={(x) => { close(); apply(x); }} />)}</div>
            <button type="button" className="pop-item" onMouseDown={(e) => e.preventDefault()} onClick={() => customRef.current && customRef.current.click()}>
              Другие цвета…
            </button>
            <input ref={customRef} type="color" className="hidden-color" defaultValue={last} onChange={(e) => { close(); apply(e.target.value); }} />
          </>
        ),
      }}
    </SplitColor>
  );
}

const HighlighterIcon = () => (
  <svg width="16" height="14" viewBox="0 0 24 20" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" aria-hidden="true">
    <path d="m14 2 6 6-9 9H5v-6z" /><path d="m11 5 6 6" />
  </svg>
);

export function HighlightButton({ editor }) {
  const [last, setLast] = useState(HIGHLIGHTS.yellow);
  const apply = (c) => { setLast(c); editor.chain().focus().setHighlight({ color: c }).run(); };
  return (
    <SplitColor title="Цвет выделения текста" color={last} icon={<HighlighterIcon />}>
      {{
        apply: () => (editor.isActive('highlight', { color: last }) ? editor.chain().focus().unsetHighlight().run() : apply(last)),
        palette: (close) => (
          <>
            <div className="swatches hl">
              {HIGHLIGHT_ORDER.map((k) => <Swatch key={k} color={HIGHLIGHTS[k]} title={HIGHLIGHT_NAMES[k]} onPick={(x) => { close(); apply(x); }} />)}
            </div>
            <button type="button" className="pop-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { close(); editor.chain().focus().unsetHighlight().run(); }}>
              <span className="swatch none" /> Нет цвета
            </button>
          </>
        ),
      }}
    </SplitColor>
  );
}
