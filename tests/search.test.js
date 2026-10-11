// Расширенный поиск (src/editor/search.js): специальные знаки ^p ^t, подстановочные знаки Word, поиск и замена формата
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { SearchReplace, findAll, replaceAllTr, replacementText, searchKey, wildcardPattern } from '../src/editor/search.js';

const p = (...content) => ({ type: 'paragraph', content });
const t = (text, marks) => (marks ? { type: 'text', text, marks: marks.map((m) => ({ type: m })) } : { type: 'text', text });
// Состояние с плагином поиска (полный набор плагинов редактора требует окна браузера)
const plugins = { schema: getSchema(extensions), pl: SearchReplace.config.addProseMirrorPlugins.call({}) };
function make(content) {
  const st = { state: EditorState.create({ schema: plugins.schema, doc: plugins.schema.nodeFromJSON({ type: 'doc', content }), plugins: plugins.pl }) };
  return {
    get state() { return st.state; },
    commands: {
      setSearch: (query, opts) => { st.state = st.state.apply(st.state.tr.setMeta(searchKey, { query, opts })); },
      replaceAll: (text) => { const tr = st.state.tr; const n = replaceAllTr(st.state, tr, text); st.state = st.state.apply(tr); return n > 0; },
    },
    destroy() {},
  };
}
const texts = (ed) => { const out = []; ed.state.doc.forEach((n) => out.push(n.textContent)); return out; };

describe('расширенный поиск', () => {
  it('подстановочные знаки Word → регулярное выражение', () => {
    expect(wildcardPattern('<(Ива)нов>')).toBe('(?<![\\p{L}\\p{N}_])(Ива)нов(?![\\p{L}\\p{N}_])');
    expect(wildcardPattern('[0-9]{2,4}')).toBe('[0-9]{2,4}');
    expect(wildcardPattern('[!а]?')).toBe('[^а][^\n]');
    expect(replacementText('\\2 \\1^p^t', ['x', 'a', 'b'], true)).toBe('b a\n\t');
    expect(replacementText('[^&]', ['найдено'], false)).toBe('[найдено]');
  });

  it('^p — граница абзацев: найти «пустой абзац» и склеить абзацы заменой', () => {
    const ed = make([p(t('Строка один')), p(t('Строка два')), p(), p(t('Строка три'))]);
    expect(findAll(ed.state.doc, '^p^p').length).toBe(1);
    ed.commands.setSearch('^p^p', {});
    ed.commands.replaceAll('^p');
    expect(texts(ed)).toEqual(['Строка один', 'Строка два', 'Строка три']);
    ed.commands.setSearch('два^pСтрока', {});
    ed.commands.replaceAll('два; строка');
    expect(texts(ed)).toEqual(['Строка один', 'Строка два; строка три']);
    ed.destroy();
  });

  it('подстановочные знаки с группами: «Иванов И.И.» → «И.И. Иванов»; ^t — табуляция', () => {
    const ed = make([p(t('Подписал Иванов И.И. и Петров П.П.'))]);
    ed.commands.setSearch('(<[А-Я][а-я]@>) ([А-Я].[А-Я].)', { wildcards: true });
    expect(ed.commands.replaceAll('\\2 \\1')).toBe(true);
    expect(texts(ed)).toEqual(['Подписал И.И. Иванов и П.П. Петров']);
    ed.commands.setSearch(' и ', {});
    ed.commands.replaceAll('^t');
    expect(texts(ed)).toEqual(['Подписал И.И. Иванов	П.П. Петров']); // «	» в узел табуляции превращает плагин редактора
    ed.destroy();
  });

  it('формат: найти полужирный (с текстом и без), заменить на курсив', () => {
    const ed = make([p(t('обычный '), t('жирный', ['bold']), t(' и снова '), t('жирный', ['bold']))]);
    expect(findAll(ed.state.doc, 'жирный', {}).length).toBe(2);
    expect(findAll(ed.state.doc, 'жирный', { format: { bold: false } }).length).toBe(0);
    expect(findAll(ed.state.doc, '', { format: { bold: true } }).map((r) => ed.state.doc.textBetween(r.from, r.to))).toEqual(['жирный', 'жирный']);
    ed.commands.setSearch('', { format: { bold: true }, replaceFormat: { bold: false, italic: true } });
    ed.commands.replaceAll('');
    const marks = [];
    ed.state.doc.descendants((n) => { if (n.isText && n.text === 'жирный') marks.push(n.marks.map((m) => m.type.name).join('+')); });
    expect(marks).toEqual(['italic', 'italic']);
    ed.destroy();
  });
});
