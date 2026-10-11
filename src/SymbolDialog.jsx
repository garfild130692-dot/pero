// «Вставка → Символ»: меню недавних символов и диалог «Символ» (шрифт, набор знаков Unicode, код знака) с вкладкой
// «Специальные знаки» (тире, неразрывные пробел и дефис, кавычки, ©, ® …) — как в Word.
import React, { useEffect, useMemo, useState } from 'react';
import { COMMON_FONTS, cssFontStack } from './editor/font.js';
import { loadSystemFonts, usePopover } from './FontControls.jsx';

export const SUBSETS = [
  ['Основная латиница', 0x20, 0x7e], ['Дополнительная латиница-1', 0xa0, 0xff], ['Расширенная латиница-A', 0x100, 0x17f],
  ['Греческий', 0x370, 0x3ff], ['Кириллица', 0x400, 0x4ff], ['Знаки пунктуации', 0x2010, 0x205e],
  ['Символы валют', 0x20a0, 0x20c0], ['Буквоподобные символы', 0x2100, 0x214f], ['Числовые формы', 0x2150, 0x218b],
  ['Стрелки', 0x2190, 0x21ff], ['Математические операторы', 0x2200, 0x22ff], ['Технические символы', 0x2300, 0x23ff],
  ['Псевдографика', 0x2500, 0x257f], ['Блоки', 0x2580, 0x259f], ['Геометрические фигуры', 0x25a0, 0x25ff],
  ['Разные символы', 0x2600, 0x26ff], ['Дингбаты', 0x2700, 0x27bf],
];

export const SPECIAL = [
  ['—', 'Длинное тире', ''], ['–', 'Короткое тире', ''], ['‑', 'Неразрывный дефис', 'Ctrl+Shift+-'],
  ['­', 'Мягкий перенос', 'Ctrl+-'], [' ', 'Неразрывный пробел', 'Ctrl+Shift+Пробел'], [' ', 'Короткий пробел', ''],
  [' ', 'Длинный пробел', ''], [' ', '1/4 пробела', ''], ['​', 'Нулевая ширина: разрешить перенос', ''],
  ['©', 'Авторское право', ''], ['®', 'Охраняемый знак', ''], ['™', 'Товарный знак', ''], ['§', 'Параграф', ''], ['¶', 'Абзац', ''],
  ['…', 'Многоточие', ''], ['«', 'Открывающая кавычка-ёлочка', ''], ['»', 'Закрывающая кавычка-ёлочка', ''],
  ['„', 'Открывающая кавычка-лапка', ''], ['“', 'Закрывающая кавычка-лапка', ''], ['‘', 'Одинарная открывающая кавычка', ''],
  ['’', 'Одинарная закрывающая кавычка, апостроф', ''], ['№', 'Номер', ''], ['°', 'Градус', ''], ['±', 'Плюс-минус', ''],
  ['×', 'Знак умножения', ''], ['÷', 'Знак деления', ''], ['‰', 'Промилле', ''], ['€', 'Евро', ''], ['₽', 'Рубль', ''],
];

const DEFAULT_RECENT = ['€', '₽', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β', 'π', 'Ω', '∑'];
const RECENT_KEY = 'pero.recentSymbols';
export function recentSymbols() {
  try {
    const v = JSON.parse(localStorage.getItem(RECENT_KEY) || 'null');
    if (Array.isArray(v) && v.length) return [...new Set([...v, ...DEFAULT_RECENT])].slice(0, 20);
  } catch { /* нет хранилища — набор по умолчанию */ }
  return DEFAULT_RECENT;
}
function rememberSymbol(sym) {
  try { localStorage.setItem(RECENT_KEY, JSON.stringify([sym.ch, ...recentSymbols().filter((x) => x !== sym.ch)].slice(0, 20))); } catch { /* не сохраняем */ }
}

// Вставить символ; font — шрифт символа (null — шрифт текста у курсора)
export function insertSymbol(editor, ch, font = null) {
  const { from } = editor.state.selection;
  const chain = editor.chain().focus().insertContent(ch);
  if (font) chain.setTextSelection({ from, to: from + ch.length }).setFontFamily(font).setTextSelection(from + ch.length);
  chain.run();
  rememberSymbol({ ch });
}

const hex = (cp) => cp.toString(16).toUpperCase().padStart(4, '0');
const visible = (ch) => ({ ' ': '°', '­': '¬', '‑': '‑', ' ': '␣', ' ': '␣', ' ': '␣', '​': '|' }[ch] || ch);

// Кнопка ленты «Символ»: недавние символы и «Другие символы…»
export function SymbolButton({ editor, Btn, onMore }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big label="Символ" title="Вставить символ" active={pop.open} onClick={() => pop.setOpen(!pop.open)}><span className="sym-big">Ω</span></Btn>
      {pop.open && (
        <div className="menu-pop sym-pop" style={pop.style} role="menu">
          <div className="sym-quick">
            {recentSymbols().map((ch) => (
              <button key={ch} type="button" className="sym-cell" title={`U+${hex(ch.codePointAt(0))}`} onMouseDown={(e) => e.preventDefault()}
                onClick={() => { pop.setOpen(false); insertSymbol(editor, ch); }}>{ch}</button>
            ))}
          </div>
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onMore(); }}>
            <span className="check">Ω</span>Другие символы…
          </button>
        </div>
      )}
    </div>
  );
}

export default function SymbolDialog({ editor, onClose }) {
  const [tab, setTab] = useState('symbols');
  const [font, setFont] = useState(''); // '' — обычный текст (шрифт у курсора)
  const [fonts, setFonts] = useState(COMMON_FONTS);
  const [subset, setSubset] = useState(1);
  const [sel, setSel] = useState(0xa9);
  const [code, setCode] = useState('00A9');
  const [special, setSpecial] = useState(0);
  useEffect(() => { loadSystemFonts().then(setFonts); }, []);
  const chars = useMemo(() => {
    const out = [];
    for (const [, a, b] of SUBSETS) for (let cp = a; cp <= b; cp++) if (cp !== 0xad) out.push(cp);
    return out;
  }, []);
  const gridRef = React.useRef(null);
  // при открытии — к выбранному знаку
  useEffect(() => {
    const el = gridRef.current && gridRef.current.querySelector('.sym-cell.sel');
    if (el) gridRef.current.scrollTop = el.offsetTop - gridRef.current.offsetTop - 40;
  }, []);
  // выбор набора — прокрутка к первому знаку набора
  const goSubset = (i) => {
    setSubset(i);
    const cp = SUBSETS[i][1];
    setSel(cp); setCode(hex(cp));
    const el = gridRef.current && gridRef.current.querySelector(`[data-cp="${cp}"]`);
    if (el) gridRef.current.scrollTop = el.offsetTop - gridRef.current.offsetTop;
  };
  const pick = (cp) => {
    setSel(cp); setCode(hex(cp));
    const i = SUBSETS.findIndex(([, a, b]) => cp >= a && cp <= b);
    if (i >= 0) setSubset(i);
  };
  const doInsert = () => {
    if (tab === 'special') insertSymbol(editor, SPECIAL[special][0]);
    else insertSymbol(editor, String.fromCodePoint(sel), font || null);
  };
  const onCode = (v) => {
    setCode(v);
    const cp = parseInt(v, 16);
    if (/^[0-9a-f]{4,5}$/i.test(v) && cp >= 0x20 && cp <= 0x10ffff) {
      setSel(cp);
      const el = gridRef.current && gridRef.current.querySelector(`[data-cp="${cp}"]`);
      if (el) el.scrollIntoView({ block: 'nearest' });
    }
  };
  const fontList = useMemo(() => [...new Set(fonts)].sort((a, b) => a.localeCompare(b, 'ru')), [fonts]);
  const recent = recentSymbols();

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd sym-dlg" role="dialog" aria-label="Символ" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); doInsert(); } }}>
        <div className="modal-title">Символ</div>
        <div className="pd-tabs" role="tablist">
          <button type="button" role="tab" className={tab === 'symbols' ? 'active' : ''} onClick={() => setTab('symbols')}>Символы</button>
          <button type="button" role="tab" className={tab === 'special' ? 'active' : ''} onClick={() => setTab('special')}>Специальные знаки</button>
        </div>
        <div className="modal-body">
          {tab === 'symbols' ? (<>
            <div className="sym-top">
              <label className="fd-col"><span>Шрифт</span>
                <select value={font} onChange={(e) => setFont(e.target.value)}>
                  <option value="">(обычный текст)</option>
                  {fontList.map((f) => <option key={f} value={f}>{f}</option>)}
                </select>
              </label>
              <label className="fd-col"><span>Набор</span>
                <select value={subset} onChange={(e) => goSubset(+e.target.value)}>
                  {SUBSETS.map(([n], i) => <option key={n} value={i}>{n}</option>)}
                </select>
              </label>
            </div>
            <div className="sym-grid" ref={gridRef} style={{ fontFamily: font ? cssFontStack(font) : undefined }}>
              {chars.map((cp) => (
                <button key={cp} type="button" data-cp={cp} className={`sym-cell${cp === sel ? ' sel' : ''}`} title={`U+${hex(cp)}`}
                  onClick={() => pick(cp)} onDoubleClick={() => { pick(cp); insertSymbol(editor, String.fromCodePoint(cp), font || null); }}>
                  {String.fromCodePoint(cp)}
                </button>
              ))}
            </div>
            <div className="pop-title">Ранее использовавшиеся символы</div>
            <div className="sym-quick">
              {recent.map((ch) => <button key={ch} type="button" className="sym-cell" title={`U+${hex(ch.codePointAt(0))}`} onClick={() => pick(ch.codePointAt(0))}>{ch}</button>)}
            </div>
            <div className="sym-code">
              <span className="sym-sample" style={{ fontFamily: font ? cssFontStack(font) : undefined }}>{String.fromCodePoint(sel)}</span>
              <label className="pd-field"><span>Код знака (Юникод, шестн.)</span>
                <span className="pd-input small"><input value={code} onChange={(e) => onCode(e.target.value.trim())} /></span>
              </label>
            </div>
          </>) : (
            <div className="sym-special" role="listbox">
              {SPECIAL.map(([ch, name, keys], i) => (
                <div key={name} role="option" aria-selected={i === special} className={`sym-row${i === special ? ' sel' : ''}`}
                  onClick={() => setSpecial(i)} onDoubleClick={() => { setSpecial(i); insertSymbol(editor, ch); }}>
                  <span className="sym-ch">{visible(ch)}</span><span>{name}</span><em>{keys}</em>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={doInsert}>Вставить</button>
          <button className="btn" onClick={onClose}>Закрыть</button>
        </div>
      </div>
    </div>
  );
}
