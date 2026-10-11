// Параметры вставки (src/editor/paste.js): объединение форматирования и «только текст»
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { _test } from '../src/editor/paste.js';

const schema = getSchema(extensions);

describe('параметры вставки', () => {
  const src = schema.nodeFromJSON({
    type: 'doc',
    content: [
      { type: 'heading', attrs: { level: 2, textAlign: 'center', spaceAfter: 30 }, content: [{ type: 'text', text: 'Заголовок', marks: [{ type: 'textStyle', attrs: { fontFamily: 'Comic Sans MS', fontSize: 20, color: '#ff0000' } }] }] },
      { type: 'paragraph', attrs: { indentLeft: 50, spaceAfter: 24, textAlign: 'right' }, content: [
        { type: 'text', text: 'жирный ', marks: [{ type: 'bold' }, { type: 'highlight', attrs: { color: '#ffff00' } }] },
        { type: 'text', text: 'ссылка', marks: [{ type: 'link', attrs: { href: 'https://example.org' } }, { type: 'textStyle', attrs: { fontSize: 18 } }] },
      ] },
    ],
  });

  it('объединение: смысловое оформление и структура остаются, шрифты, цвета, интервалы и отступы — документа', () => {
    const frag = _test.mergeFragment(src.content, schema);
    const json = frag.toJSON();
    expect(json[0].type).toBe('heading');
    expect(json[0].attrs.level).toBe(2);
    expect(json[0].attrs.textAlign).toBe(null);
    expect(json[0].attrs.spaceAfter).toBe(null);
    expect(json[0].content[0].marks).toBeUndefined();
    expect(json[1].attrs.indentLeft).toBe(null);
    expect(json[1].content[0].marks).toEqual([{ type: 'bold' }]);
    expect(json[1].content[1].marks.map((m) => m.type)).toEqual(['link']);
  });

  it('только текст: строки — абзацы, оформление — как у места вставки', () => {
    const doc = schema.nodeFromJSON({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'курсив', marks: [{ type: 'italic' }] }] }] });
    let s = EditorState.create({ schema, doc });
    s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, 4)));
    const slice = _test.textSlice('первая\r\nвторая', s);
    expect(slice.content.childCount).toBe(2);
    expect(slice.content.child(0).textContent).toBe('первая');
    expect(slice.content.child(1).firstChild.marks.map((m) => m.type.name)).toEqual(['italic']);
    const r = s.apply(s.tr.replaceSelection(slice));
    expect(r.doc.childCount).toBe(2);
    expect(r.doc.child(0).textContent).toBe('курпервая');
    expect(r.doc.child(1).textContent).toBe('втораясив');
  });
});
