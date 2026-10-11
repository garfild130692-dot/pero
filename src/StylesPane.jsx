// Панель «Стили» (кнопка в углу группы «Стили», Ctrl+Alt+Shift+S): все стили — абзацные (¶) и символьные (a),
// стиль под курсором выделен, щелчок — применить; создать, изменить, очистить формат — как в Word.
import React from 'react';
import { useEditorState } from '@tiptap/react';
import { BUILTIN_CHAR, BUILTIN_PARA, isBuiltinName, styleLabel } from './editor/builtinStyles.js';

// Все стили для панели и галереи: [{ key, label, kind: 'para' | 'char', apply(editor), active(state) }]
export function styleEntries(s) {
  const out = [];
  const byName = (list, name) => (list || []).find((x) => String(x.name).toLowerCase() === name.toLowerCase());
  out.push({ key: 'normal', label: 'Обычный', kind: 'para', cls: 'st-normal', active: !s.heading && !s.styleId, apply: (ed) => ed.chain().focus().setParagraphStyle(null).run() });
  for (const def of BUILTIN_PARA) {
    const have = byName(s.styles, def.name);
    out.push({ key: def.id, label: def.label, kind: 'para', def: have || def, active: !!have && s.styleId === have.id, apply: (ed) => ed.chain().focus().applyBuiltinParaStyle(def).run() });
    if (def.id === 'NoSpacing') {
      for (const lvl of [1, 2, 3, 4, 5, 6]) {
        out.push({ key: `h${lvl}`, label: `Заголовок ${lvl}`, kind: 'para', cls: `st-h${Math.min(lvl, 3)}`, heading: lvl, active: s.heading === lvl, apply: (ed) => ed.chain().focus().setHeading({ level: lvl }).run() });
      }
    }
  }
  for (const st of s.styles || []) {
    if (isBuiltinName(st.name)) continue; // уже в списке как встроенный
    out.push({ key: st.id, label: st.name, kind: 'para', def: st, custom: true, active: s.styleId === st.id, apply: (ed) => ed.chain().focus().setParagraphStyle(st.id).run() });
  }
  for (const def of BUILTIN_CHAR) {
    const have = byName(s.charStyles, def.name);
    out.push({ key: `c-${def.id}`, label: def.label, kind: 'char', def: have || def, active: !!have && s.charStyle === have.id, apply: (ed) => ed.chain().focus().applyCharStyle(def).run() });
  }
  for (const st of s.charStyles || []) {
    if (isBuiltinName(st.name)) continue;
    out.push({ key: `c-${st.id}`, label: st.name, kind: 'char', def: st, active: s.charStyle === st.id, apply: (ed) => ed.chain().focus().applyCharStyle(st).run() });
  }
  return out;
}

// Состояние для галереи и панели
export function styleState(e) {
  const { state } = e;
  const cs = (state.storedMarks || state.selection.$from.marks()).find((m) => m.type.name === 'charStyle');
  return {
    heading: [1, 2, 3, 4, 5, 6].find((l) => e.isActive('heading', { level: l })) || 0,
    styleId: state.selection.$from.parent.attrs.styleId || null,
    styles: state.doc.attrs.styles || [],
    charStyles: state.doc.attrs.charStyles || [],
    charStyle: cs ? cs.attrs.id : null,
  };
}

export const sampleOf = (def) => {
  const r = (def && def.run) || {};
  return {
    fontFamily: r.fontFamily || undefined, color: r.color || undefined,
    fontWeight: r.bold ? 'bold' : undefined, fontStyle: r.italic ? 'italic' : undefined, textDecoration: r.underline ? 'underline' : undefined,
  };
};

export default function StylesPane({ editor, onStyle, onClose }) {
  const s = useEditorState({ editor, selector: ({ editor: e }) => styleState(e), equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b) });
  const list = styleEntries(s);
  const cur = list.find((x) => x.active && x.kind === 'para');
  return (
    <aside className="comments-panel styles-pane" aria-label="Стили">
      <div className="cp-head">
        <span>Стили</span>
        <button type="button" className="cc-del" title="Скрыть панель" onClick={onClose}>✕</button>
      </div>
      <button type="button" className="sp-item sp-clear" onMouseDown={(e) => e.preventDefault()} onClick={() => editor.chain().focus().clearFormatting().setParagraphStyle(null).run()}>
        Очистить всё
      </button>
      <div className="sp-list" role="listbox">
        {list.map((x) => (
          <button key={x.key} type="button" role="option" aria-selected={x.active} className={`sp-item${x.active ? ' active' : ''}`}
            title={x.kind === 'char' ? 'Символьный стиль: применяется к выделенному тексту (без выделения — к слову)' : 'Стиль абзаца'}
            onMouseDown={(e) => e.preventDefault()} onClick={() => x.apply(editor)}>
            <span className={`sp-name ${x.cls || ''}`} style={x.def ? sampleOf(x.def) : undefined}>{x.label}</span>
            <span className="sp-kind">{x.kind === 'char' ? 'a' : '¶'}</span>
          </button>
        ))}
      </div>
      <div className="sp-actions">
        <button type="button" className="btn small" onClick={() => onStyle('new')} title="Создать стиль из текущего абзаца">Создать стиль</button>
        <button type="button" className="btn small" disabled={!cur || (!cur.custom && cur.key !== 'normal')} onClick={() => onStyle(cur.key === 'normal' ? 'normal' : cur.key)}
          title="Изменить стиль под курсором (свой или «Обычный»)">Изменить…</button>
      </div>
      <div className="sp-note">Текущий: {cur ? cur.label : '—'}{s.charStyle ? ` + ${styleLabel((s.charStyles.find((c) => c.id === s.charStyle) || {}).name || s.charStyle)}` : ''}</div>
    </aside>
  );
}
