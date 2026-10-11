// Вкладка «Вид» (src/ViewControls.jsx): масштаб «по ширине»/«страница целиком», структура заголовков для области навигации
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { clampZoom, fitZoom, headingsOutline } from '../src/ViewControls.jsx';
import { makeHeadingNumbering } from '../src/editor/numbering.js';

const schema = getSchema(extensions);
const h = (level, text, attrs = {}) => ({ type: 'heading', attrs: { level, ...attrs }, content: [{ type: 'text', text }] });
const p = (text) => ({ type: 'paragraph', content: [{ type: 'text', text }] });

describe('вид', () => {
  it('масштаб: пределы 10–500 %, по ширине и страница целиком', () => {
    expect(clampZoom(5)).toBe(10);
    expect(clampZoom(900)).toBe(500);
    const a4 = { width: 794, height: 1123 };
    expect(fitZoom('width', { w: 1000, h: 700 }, a4)).toBe(120);
    expect(fitZoom('page', { w: 1000, h: 700 }, a4)).toBe(59);
  });

  it('область навигации: заголовки с номерами, текущий — по курсору', () => {
    const doc = schema.nodeFromJSON({
      type: 'doc',
      attrs: { headingNumbering: makeHeadingNumbering('dots') },
      content: [h(1, 'Общие'), p('текст'), h(2, 'Термины'), h(2, 'Без номера', { noNumber: true }), h(2, 'Права'), h(1, 'Итоги'), p('конец')],
    });
    let s = EditorState.create({ schema, doc });
    const pos = (t) => { let r = 0; s.doc.descendants((n, q) => { if (n.isText && n.text === t) r = q; }); return r; };
    s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, pos('Права') + 1)));
    const { list, current } = headingsOutline(s);
    expect(list.map((x) => `${x.label}|${x.text}`)).toEqual(['1.|Общие', '1.1.|Термины', '|Без номера', '1.2.|Права', '2.|Итоги']);
    expect(current).toBe(3);
  });
});
