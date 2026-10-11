// Регистр и формат по образцу (src/editor/runFormat.js) — на состоянии ProseMirror без DOM
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { applyFormatTr, captureFormat, changeCaseTr, nextCase, transformCase } from '../src/editor/runFormat.js';

const schema = getSchema(extensions);
const doc = (paras) => schema.nodeFromJSON({ type: 'doc', content: paras });
const p = (content, attrs = {}) => ({ type: 'paragraph', attrs, content });
const t = (text, marks) => (marks ? { type: 'text', text, marks } : { type: 'text', text });
const stateOf = (d, from, to = from) => {
  const s = EditorState.create({ schema, doc: d });
  return s.apply(s.tr.setSelection(TextSelection.create(s.doc, from, to)));
};

describe('регистр', () => {
  it('преобразования как в Word', () => {
    expect(transformCase('привет мир. как дела? хорошо', 'sentence')).toBe('Привет мир. Как дела? Хорошо');
    expect(transformCase('ПРИВЕТ МИР', 'lower')).toBe('привет мир');
    expect(transformCase('привет мир', 'upper')).toBe('ПРИВЕТ МИР');
    expect(transformCase('привет мир-друг', 'title')).toBe('Привет Мир-Друг');
    expect(transformCase('Привет Мир', 'toggle')).toBe('пРИВЕТ мИР');
    expect(transformCase('straße', 'upper')).toBe('STRASSE');
  });
  it('Shift+F3 по кругу: строчные → ПРОПИСНЫЕ → С Прописных → строчные', () => {
    expect(nextCase('привет мир')).toBe('upper');
    expect(nextCase('ПРИВЕТ МИР')).toBe('title');
    expect(nextCase('Привет Мир')).toBe('lower');
    expect(nextCase('Я')).toBe('lower');
  });
  it('меняет текст, сохраняя оформление; без выделения — слово под курсором', () => {
    const d = doc([p([t('один '), t('два', [{ type: 'bold' }]), t(' три')])]);
    const s1 = stateOf(d, 1, 13);
    const r1 = s1.apply(changeCaseTr(s1, 'upper'));
    expect(r1.doc.textContent).toBe('ОДИН ДВА ТРИ');
    expect(r1.doc.firstChild.child(1).marks.map((m) => m.type.name)).toEqual(['bold']);
    expect([r1.selection.from, r1.selection.to]).toEqual([1, 13]);
    const s2 = stateOf(d, 8); // внутри «два»
    expect(s2.apply(changeCaseTr(s2, 'cycle')).doc.textContent).toBe('один ДВА три');
    // «Как в предложениях» на середине абзаца: начало предложения — по тексту перед выделением
    const d3 = doc([p([t('Конец. начало фразы')])]);
    const s3 = stateOf(d3, 8, 20);
    expect(s3.apply(changeCaseTr(s3, 'sentence')).doc.textContent).toBe('Конец. Начало фразы');
  });
});

describe('формат по образцу', () => {
  it('переносит оформление текста и абзаца', () => {
    const d = doc([
      p([t('образец', [{ type: 'bold' }, { type: 'textStyle', attrs: { fontFamily: 'Arial', fontSize: 14, color: '#c00000' } }])], { textAlign: 'center', spaceAfter: 20 }),
      p([t('цель ', [{ type: 'italic' }]), t('текст')]),
    ]);
    const src = stateOf(d, 3); // курсор в образце — формат текста и абзаца
    const fmt = captureFormat(src);
    expect(fmt.run.map((m) => m.type).sort()).toEqual(['bold', 'textStyle']);
    expect(fmt.para.attrs.textAlign).toBe('center');
    // цель — весь второй абзац
    const tgt = stateOf(d, 10, 20);
    const r = tgt.apply(applyFormatTr(tgt, fmt));
    const para = r.doc.child(1);
    expect(para.attrs.textAlign).toBe('center');
    expect(para.attrs.spaceAfter).toBe(20);
    para.forEach((n) => {
      expect(n.marks.map((m) => m.type.name).sort()).toEqual(['bold', 'textStyle']);
      expect(n.marks.find((m) => m.type.name === 'textStyle').attrs.fontFamily).toBe('Arial');
    });
  });
  it('часть абзаца в образце — только оформление текста; пустая цель — слово под курсором', () => {
    const d = doc([p([t('аа '), t('бб', [{ type: 'underline' }])]), p([t('вв гг')], { textAlign: 'right' })]);
    const fmt = captureFormat(stateOf(d, 4, 6));
    expect(fmt.para).toBe(null);
    const tgt = stateOf(d, 12); // в слове «гг»
    const r = tgt.apply(applyFormatTr(tgt, fmt));
    const para = r.doc.child(1);
    expect(para.attrs.textAlign).toBe('right');
    expect(para.child(0).text).toBe('вв ');
    expect(para.child(1).marks.map((m) => m.type.name)).toEqual(['underline']);
  });
});
