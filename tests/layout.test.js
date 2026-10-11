// Вкладка «Макет»: счётчики абзаца (src/ParaSpinners.jsx)
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { paraValue } from '../src/ParaSpinners.jsx';

const schema = getSchema(extensions);

describe('«Макет → Абзац»', () => {
  it('действующее значение абзацев выделения, разные — null', () => {
    const doc = schema.nodeFromJSON({
      type: 'doc', attrs: { para: { spaceAfter: 8 } },
      content: [
        { type: 'paragraph', attrs: { indentLeft: 28.35 }, content: [{ type: 'text', text: 'раз' }] },
        { type: 'paragraph', content: [{ type: 'text', text: 'два' }] },
      ],
    });
    const at = (from, to) => EditorState.create({ schema, doc, selection: TextSelection.create(doc, from, to) });
    expect(paraValue(at(2, 2), 'indentLeft')).toBe(28.35);
    expect(paraValue(at(2, 2), 'spaceAfter')).toBe(8);
    expect(paraValue(at(2, 7), 'indentLeft')).toBe(null);
    expect(paraValue(at(2, 7), 'spaceAfter')).toBe(8);
  });
});
