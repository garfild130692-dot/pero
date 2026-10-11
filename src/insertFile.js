// «Вставка → Экспресс-блоки → Текст из файла»: содержимое другого .docx — в место курсора.
// Переносятся текст, оформление, таблицы, списки, рисунки (как новые), стили абзацев и символьные стили, сноски
// (с новыми номерами). Что ссылается на части чужого пакета (фигуры, объекты, примечания) — не переносится.
import { Slice, Fragment } from '@tiptap/pm/model';
import { importDocx } from './docx/import.js';

const REL_RE = /\br:(?:id|embed|link|pict)="/;

// Подготовить JSON вставляемого документа: { content, styles, charStyles, footnotes, dropped }
export function prepareInsert(json, target) {
  const a = json.attrs || {};
  let maxNote = 0;
  for (const n of target.footnotes || []) maxNote = Math.max(maxNote, +n.id || 0);
  const noteMap = new Map();
  const footnotes = [];
  let dropped = 0;
  const walk = (node) => {
    if (node.type === 'opaque') {
      const { kind, xml } = node.attrs || {};
      if (kind === 'footnote') {
        const id = (/w:id="(-?\d+)"/.exec(xml || '') || [])[1];
        const note = (a.footnotes || []).find((f) => String(f.id) === String(id));
        if (!noteMap.has(id)) {
          const nid = String(++maxNote);
          noteMap.set(id, nid);
          footnotes.push({ id: nid, text: note ? note.text : '', changed: true });
        }
        return { ...node, attrs: { ...node.attrs, xml: xml.replace(/w:id="-?\d+"/, `w:id="${noteMap.get(id)}"`) } };
      }
      if (['comment', 'commentRange'].includes(kind) || REL_RE.test(xml || '')) { dropped++; return null; }
      return node;
    }
    if (node.type === 'image') return { ...node, attrs: { ...node.attrs, xml: null, target: null } }; // в этом пакете — новая картинка
    const out = { ...node };
    if (node.attrs && node.attrs.sectPr) out.attrs = { ...node.attrs, sectPr: null }; // разделы вставляемого файла — нет
    if (node.content) out.content = node.content.map(walk).filter(Boolean);
    return out;
  };
  const content = (json.content || []).map(walk).filter(Boolean);
  const have = new Set((target.styles || []).map((s) => s.id));
  const haveC = new Set((target.charStyles || []).map((s) => s.id));
  return {
    content,
    styles: (a.styles || []).filter((s) => !have.has(s.id)).map((s) => ({ ...s, origin: 'pero', modified: true })),
    charStyles: (a.charStyles || []).filter((s) => !haveC.has(s.id)).map((s) => ({ ...s, origin: 'pero' })),
    footnotes,
    dropped,
  };
}

// Вставить файл (данные .docx) в редактор; возвращает число непереносимых элементов
export async function insertDocxFile(editor, data) {
  const { doc } = await importDocx(data);
  const da = editor.state.doc.attrs;
  const p = prepareInsert(doc, da);
  const { schema } = editor.state;
  const nodes = p.content.map((n) => schema.nodeFromJSON(n));
  editor.chain().focus().command(({ tr }) => {
    tr.replaceSelection(new Slice(Fragment.from(nodes), 0, 0));
    if (p.styles.length) tr.setDocAttribute('styles', [...(da.styles || []), ...p.styles]);
    if (p.charStyles.length) tr.setDocAttribute('charStyles', [...(da.charStyles || []), ...p.charStyles]);
    if (p.footnotes.length) tr.setDocAttribute('footnotes', [...(da.footnotes || []), ...p.footnotes]).setDocAttribute('footnotesEdited', true);
    return true;
  }).run();
  return p.dropped;
}
