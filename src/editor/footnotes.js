// Сноски (обычные, внизу страницы) — как в Word.
//
// Тексты — doc.attrs.footnotes: [{ id, text }] (text — абзацы через \n), footnotesEdited — менялись в Пере.
// В тексте сноска — значок opaque kind 'footnote' (<w:footnoteReference w:id>), номер на экране — по порядку (CSS-счётчик).
// Разбиение на страницы (pagination.js) оставляет внизу листа место под сноски его ссылок и рисует их там.
// При сохранении: не менялись — footnotes.xml как был; менялись — неизменённые сноски пишутся исходной разметкой,
// изменённые и новые — заново (простой текст), разделители — как были (нет части — создаются).
import { Extension } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const FOOTNOTES_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';

const idOf = (xml) => { const m = /w:id="(-?\d+)"/.exec(xml || ''); return m ? m[1] : null; };
export const isFootnoteRef = (n) => n && n.type.name === 'opaque' && n.attrs.kind === 'footnote';

// Ссылки на сноски по порядку: [{ pos, id, n }] (n — номер, как покажет Word)
export function footnoteRefs(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (isFootnoteRef(node)) { const id = idOf(node.attrs.xml); if (id != null) out.push({ pos, id, n: out.length + 1 }); return false; }
    return true;
  });
  return out;
}

// Разбор footnotes.xml (xmldom) → обычные сноски (разделители пропускаются)
export function readFootnotes(xdoc) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const out = [];
  for (const f of Array.from(xdoc.getElementsByTagNameNS(W, 'footnote'))) {
    const get = (n) => f.getAttributeNS(W, n) || f.getAttribute('w:' + n) || '';
    if (get('type') && get('type') !== 'normal') continue;
    const paras = Array.from(f.getElementsByTagNameNS(W, 'p')).map((p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join(''));
    out.push({ id: get('id'), text: paras.join('\n').replace(/^\s+/, '') });
  }
  return out;
}

// Разметка одной сноски (новой или изменённой). st — id стилей «Текст сноски» и «Знак сноски» в пакете (или null)
export function footnoteXml(note, st = {}) {
  const pPr = st.text ? `<w:pPr><w:pStyle w:val="${esc(st.text)}"/></w:pPr>` : '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';
  const refR = st.ref ? `<w:rPr><w:rStyle w:val="${esc(st.ref)}"/></w:rPr>` : '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>';
  const rPr = st.text ? '' : '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
  const paras = String(note.text || '').split('\n').map((t, i) => `<w:p>${pPr}${i === 0 ? `<w:r>${refR}<w:footnoteRef/></w:r><w:r>${rPr}<w:t xml:space="preserve"> </w:t></w:r>` : ''}${t ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : ''}</w:p>`).join('');
  return `<w:footnote w:id="${esc(note.id)}">${paras}</w:footnote>`;
}

const SEPARATORS = '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

// footnotes.xml: original — исходная часть (строка) или null; edited — id изменённых в Пере сносок; used — id со ссылками в тексте
export function footnotesPartXml(list, ns, { original = null, changed = new Set(), styles = {} } = {}) {
  const orig = new Map();
  let seps = SEPARATORS;
  let root = `<w:footnotes ${ns}>`;
  if (original) {
    const m = /<w:footnotes\b[^>]*>/.exec(original);
    if (m) root = m[0];
    const sepList = [];
    for (const f of original.match(/<w:footnote\b[^>]*?(?:\/>|>[\s\S]*?<\/w:footnote>)/g) || []) {
      const head = /^<w:footnote\b[^>]*>/.exec(f)[0];
      const type = (/w:type="([^"]*)"/.exec(head) || [])[1];
      if (type && type !== 'normal') sepList.push(f); else orig.set(idOf(head), f);
    }
    if (sepList.length) seps = sepList.join('');
  }
  const body = list.map((n) => (orig.has(String(n.id)) && !changed.has(String(n.id)) ? orig.get(String(n.id)) : footnoteXml(n, styles))).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${root}${seps}${body}</w:footnotes>`;
}

// Новый id сноски: больше всех известных (в списке и в ссылках текста)
function nextId(doc) {
  let max = 0;
  for (const n of doc.attrs.footnotes || []) max = Math.max(max, +n.id || 0);
  for (const r of footnoteRefs(doc)) max = Math.max(max, +r.id || 0);
  return String(max + 1);
}

export function insertFootnoteTr(state) {
  const id = nextId(state.doc);
  const ref = state.schema.nodes.opaque.create({ kind: 'footnote', xml: `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>` });
  const at = state.selection.to; // Word вставляет знак сноски после выделения
  const tr = state.tr.insert(at, ref);
  tr.setDocAttribute('footnotes', [...(state.doc.attrs.footnotes || []), { id, text: '', changed: true }]);
  tr.setDocAttribute('footnotesEdited', true);
  tr.setSelection(TextSelection.create(tr.doc, at + 1));
  return { tr, id };
}

export function setFootnoteTextTr(state, id, text) {
  const list = state.doc.attrs.footnotes || [];
  const has = list.some((n) => String(n.id) === String(id));
  const next = has ? list.map((n) => (String(n.id) === String(id) ? { ...n, text, changed: true } : n)) : [...list, { id: String(id), text, changed: true }];
  return state.tr.setDocAttribute('footnotes', next).setDocAttribute('footnotesEdited', true);
}

export const Footnotes = Extension.create({
  name: 'footnotes',
  addStorage() {
    return { onOpen: null }; // App: открыть панель сносок на сноске id
  },
  addCommands() {
    return {
      insertFootnote: () => ({ state, dispatch }) => {
        const { tr, id } = insertFootnoteTr(state);
        if (dispatch) { dispatch(tr.scrollIntoView()); if (this.storage.onOpen) setTimeout(() => this.storage.onOpen(id), 0); }
        return true;
      },
      setFootnoteText: (id, text) => ({ state, dispatch }) => { if (dispatch) dispatch(setFootnoteTextTr(state, id, text)); return true; },
      deleteFootnote: (id) => ({ state, dispatch }) => {
        const r = footnoteRefs(state.doc).find((x) => String(x.id) === String(id));
        if (!r) return false;
        if (dispatch) dispatch(state.tr.delete(r.pos, r.pos + 1).setDocAttribute('footnotesEdited', true));
        return true;
      },
      // к ссылке на сноску id или к следующей/предыдущей (dir = 1 / -1)
      goToFootnote: (target) => ({ state, dispatch }) => {
        const refs = footnoteRefs(state.doc);
        const at = state.selection.from;
        const r = typeof target === 'number' && Math.abs(target) === 1
          ? (target > 0 ? refs.find((x) => x.pos >= at + (state.selection.empty ? 0 : 1) && x.pos !== at - 1) : [...refs].reverse().find((x) => x.pos < at - 1))
          : refs.find((x) => String(x.id) === String(target));
        if (!r) return false;
        if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, r.pos + 1)).scrollIntoView());
        return true;
      },
    };
  },
  addKeyboardShortcuts() {
    return { 'Mod-Alt-f': () => this.editor.commands.insertFootnote() };
  },
});
