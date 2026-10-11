// Панель «Найти и заменить» (Ctrl+F / Ctrl+H): поле поиска, счётчик, переход, параметры, замена; «Больше» — подстановочные
// знаки, специальные знаки (^p, ^t…), формат для поиска и для замены (как «Больше >>» в Word).
import React, { useEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { FORMAT_FLAGS, searchInfo } from './editor/search.js';
import { usePopover } from './FontControls.jsx';

const SPECIALS = [
  ['^p', 'Знак абзаца'], ['^t', 'Знак табуляции'], ['^l', 'Разрыв строки'], ['^m', 'Разрыв страницы'], ['^s', 'Неразрывный пробел'],
  ['^~', 'Неразрывный дефис'], ['^-', 'Мягкий перенос'], ['^?', 'Любой знак'], ['^#', 'Любая цифра'], ['^$', 'Любая буква'],
  ['^w', 'Пробелы (любые)'], ['^^', 'Знак «^»'], ['^&', 'Найденный текст (в замене)'],
];

// Меню вставки кода в поле
function SpecialMenu({ onPick }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className="btn small" onClick={() => pop.setOpen(!pop.open)} title="Вставить специальный знак в поле">Специальный ▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {SPECIALS.map(([code, label]) => (
            <button key={code} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onPick(code); }}>
              <span className="check">{code}</span>{label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// Формат: по каждому признаку — «не важно» / «да» / «нет» (щелчок по кругу); стиль абзаца
function FormatPicker({ title, value, onChange, styles }) {
  const cycle = (k) => {
    const v = value[k];
    const next = v == null ? true : v === true ? false : undefined;
    const o = { ...value };
    if (next === undefined) delete o[k]; else o[k] = next;
    onChange(o);
  };
  return (
    <div className="fb-format">
      <span className="fb-format-title">{title}</span>
      {FORMAT_FLAGS.map(([k, label]) => (
        <button key={k} type="button" className={`fb-chip${value[k] === true ? ' on' : value[k] === false ? ' off' : ''}`}
          title={`${label}: ${value[k] === true ? 'да' : value[k] === false ? 'нет' : 'не важно'} (щелчок — сменить)`} onClick={() => cycle(k)}>
          {value[k] === false ? 'не ' : ''}{label.toLowerCase()}
        </button>
      ))}
      {styles && (
        <select value={value.styleId === undefined ? '' : value.styleId ?? '__normal'} title="Стиль абзаца"
          onChange={(e) => { const o = { ...value }; const v = e.target.value; if (!v) delete o.styleId; else o.styleId = v === '__normal' ? null : v; onChange(o); }}>
          <option value="">стиль — любой</option>
          <option value="__normal">Обычный</option>
          {[1, 2, 3, 4, 5, 6].map((l) => <option key={l} value={`h${l}`}>Заголовок {l}</option>)}
          {styles.map((st) => <option key={st.id} value={st.id}>{st.name}</option>)}
        </select>
      )}
      {Object.keys(value).length > 0 && <button type="button" className="fb-chip clear" onClick={() => onChange({})}>без формата</button>}
    </div>
  );
}

export default function FindBar({ editor, mode, setMode, initialFormat = null, initialMore = false }) {
  const [query, setQuery] = useState(() => {
    if (initialFormat) return '';
    const { from, to } = editor.state.selection;
    const sel = editor.state.doc.textBetween(from, to, ' ');
    return sel && sel.length < 100 && !sel.includes('\n') ? sel : '';
  });
  const [repl, setRepl] = useState('');
  const [caseSensitive, setCase] = useState(false);
  const [wholeWord, setWhole] = useState(false);
  const [more, setMore] = useState(!!initialFormat || initialMore);
  const [wildcards, setWildcards] = useState(false);
  const [format, setFormat] = useState(initialFormat || {});
  const [replaceFormat, setReplaceFormat] = useState({});
  const [note, setNote] = useState('');
  const [lastField, setLastField] = useState('find');
  const findRef = useRef(null);
  const replRef = useRef(null);
  const info = useEditorState({ editor, selector: ({ editor: e }) => searchInfo(e.state) });
  const styles = useEditorState({ editor, selector: ({ editor: e }) => e.state.doc.attrs.styles || [] });

  useEffect(() => { findRef.current && findRef.current.select(); }, [mode]);
  useEffect(() => {
    editor.commands.setSearch(query, { caseSensitive, wholeWord, wildcards, format, replaceFormat });
    setNote('');
  }, [editor, query, caseSensitive, wholeWord, wildcards, format, replaceFormat]);
  useEffect(() => () => editor.commands.clearSearch(), [editor]);

  const close = () => { setMode(null); editor.commands.focus(); };
  const onKey = (e) => {
    if (e.key === 'Escape') { e.preventDefault(); close(); }
    if (e.key === 'Enter' && e.target === findRef.current) { e.preventDefault(); e.shiftKey ? editor.commands.findPrev() : editor.commands.findNext(); }
  };
  const searching = query || Object.keys(format).length;
  const counter = !searching ? '' : info.count ? `${info.index + 1} из ${info.count}` : 'Нет совпадений';
  const insertCode = (code) => {
    const ref = lastField === 'repl' && mode === 'replace' ? replRef : findRef;
    const set = ref === replRef ? setRepl : setQuery;
    const el = ref.current;
    const v = el ? el.value : '';
    const a = el ? el.selectionStart : v.length; const b = el ? el.selectionEnd : v.length;
    set(v.slice(0, a) + code + v.slice(b));
    setTimeout(() => { if (el) { el.focus(); el.setSelectionRange(a + code.length, a + code.length); } }, 0);
  };

  return (
    <div className="findbar" role="search" onKeyDown={onKey}>
      <div className="fb-row">
        <input ref={findRef} className="fb-input" value={query} onChange={(e) => setQuery(e.target.value)} onFocus={() => setLastField('find')}
          placeholder={wildcards ? 'Найти (подстановочные знаки: ? * [а-я] @ < >)' : 'Найти в документе'} aria-label="Найти" />
        <span className={`fb-count${searching && !info.count ? ' none' : ''}`}>{counter}</span>
        <button type="button" className="fb-btn" title="Предыдущее (Shift+Enter)" disabled={!info.count} onClick={() => editor.commands.findPrev()}>▲</button>
        <button type="button" className="fb-btn" title="Следующее (Enter)" disabled={!info.count} onClick={() => editor.commands.findNext()}>▼</button>
        <button type="button" className={`fb-btn opt${caseSensitive ? ' on' : ''}`} title="С учётом регистра" disabled={wildcards} onClick={() => setCase(!caseSensitive)}>Аа</button>
        <button type="button" className={`fb-btn opt${wholeWord ? ' on' : ''}`} title="Только слово целиком" disabled={wildcards} onClick={() => setWhole(!wholeWord)}>«ab»</button>
        <button type="button" className={`fb-btn opt${more ? ' on' : ''}`} title="Больше: подстановочные знаки, специальные знаки, формат" onClick={() => setMore(!more)}>⋯</button>
        <button type="button" className="fb-btn" title={mode === 'replace' ? 'Скрыть замену' : 'Заменить (Ctrl+H)'} onClick={() => setMode(mode === 'replace' ? 'find' : 'replace')}>⇄</button>
        <button type="button" className="fb-btn" title="Закрыть (Esc)" onClick={close}>✕</button>
      </div>
      {mode === 'replace' && (
        <div className="fb-row">
          <input ref={replRef} className="fb-input" value={repl} onChange={(e) => setRepl(e.target.value)} onFocus={() => setLastField('repl')}
            placeholder={wildcards ? 'Заменить на (\\1 — группа, ^& — найденное)' : 'Заменить на'} aria-label="Заменить на"
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.commands.replaceCurrent(repl); } }} />
          <button type="button" className="btn small" disabled={!info.count} onClick={() => editor.commands.replaceCurrent(repl)}>Заменить</button>
          <button type="button" className="btn small" disabled={!info.count} onClick={() => {
            const n = info.count;
            editor.commands.replaceAll(repl);
            setNote(`Заменено: ${n}`);
          }}>Заменить все</button>
          {note && <span className="fb-note">{note}</span>}
        </div>
      )}
      {more && (
        <div className="fb-more">
          <div className="fb-row">
            <label className="fb-check" title="? — любой знак, * — любые знаки, [абв] [а-я] [!а] — из набора, {2,4} — повторы, @ — один и больше, < > — начало/конец слова, ( ) и \1 в замене">
              <input type="checkbox" checked={wildcards} onChange={(e) => setWildcards(e.target.checked)} /> Подстановочные знаки
            </label>
            <SpecialMenu onPick={insertCode} />
          </div>
          <FormatPicker title="Формат поиска:" value={format} onChange={setFormat} styles={styles} />
          {mode === 'replace' && <FormatPicker title="Формат замены:" value={replaceFormat} onChange={setReplaceFormat} />}
        </div>
      )}
    </div>
  );
}
