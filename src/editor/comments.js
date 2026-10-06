// Примечания: подсветка комментируемых фрагментов и команды (создать, изменить текст, удалить).
//
// Тексты — в doc.attrs.comments: [{ id, author, initials, date, text }] (text — абзацы через \n);
// commentsEdited — примечания менялись в Пере (при сохранении comments.xml пишется заново).
// В тексте примечание — значки opaque: commentRange (начало/конец диапазона, невидимые) и comment (ссылка).
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export const commentsKey = new PluginKey('comments');

// Дата примечания как пишет Word: местное время с суффиксом Z (в w:date Word хранит местное время, не UTC)
export function wordDate(d = new Date()) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}Z`;
}
// Показ даты примечания — без пересчёта часового пояса (см. wordDate)
export function fmtWordDate(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(s || '');
  return m ? `${m[3]}.${m[2]}.${m[1]}, ${m[4]}:${m[5]}` : '';
}
const idOf = (xml) => { const m = /w:id="(-?\d+)"/.exec(xml || ''); return m ? m[1] : null; };

// Позиции примечаний в документе: id → { start, end, ref }
export function commentAnchors(doc) {
  const out = {};
  doc.descendants((node, pos) => {
    if (node.type.name !== 'opaque') return true;
    const { kind, xml } = node.attrs;
    if (kind !== 'commentRange' && kind !== 'comment') return false;
    const id = idOf(xml);
    if (id == null) return false;
    const a = (out[id] = out[id] || {});
    if (kind === 'comment') a.ref = pos;
    else if (/commentRangeStart/.test(xml)) a.start = pos + 1;
    else a.end = pos;
    return false;
  });
  return out;
}

function decorate(doc, active) {
  const decos = [];
  for (const [id, a] of Object.entries(commentAnchors(doc))) {
    if (a.start != null && a.end != null && a.end > a.start) {
      decos.push(Decoration.inline(a.start, a.end, { class: id === active ? 'comment-range active' : 'comment-range', 'data-cid': id }));
    }
  }
  return DecorationSet.create(doc, decos);
}

export const Comments = Extension.create({
  name: 'comments',

  addProseMirrorPlugins() {
    return [new Plugin({
      key: commentsKey,
      state: {
        init: (_c, state) => ({ active: null, deco: decorate(state.doc, null) }),
        apply(tr, old, _o, state) {
          const meta = tr.getMeta(commentsKey);
          const active = meta && 'active' in meta ? meta.active : old.active;
          if (!tr.docChanged && active === old.active) return old;
          return { active, deco: decorate(state.doc, active) };
        },
      },
      props: {
        decorations: (state) => commentsKey.getState(state).deco,
        // щелчок в комментируемом фрагменте — сделать его примечание активным
        handleClick(view, pos) {
          const anchors = commentAnchors(view.state.doc);
          const hit = Object.entries(anchors).find(([, a]) => a.start != null && a.end != null && pos >= a.start && pos <= a.end);
          const active = hit ? hit[0] : null;
          if (active !== commentsKey.getState(view.state).active) view.dispatch(view.state.tr.setMeta(commentsKey, { active }));
          return false;
        },
      },
    })];
  },

  addCommands() {
    const setComments = (tr, list) => tr.setDocAttribute('comments', list).setDocAttribute('commentsEdited', true);
    return {
      // Новое примечание к выделенному фрагменту (без выделения — к месту курсора)
      addComment: ({ author, initials, text = '' }) => ({ state, tr, dispatch }) => {
        const list = state.doc.attrs.comments || [];
        let max = -1;
        for (const c of list) max = Math.max(max, +c.id);
        for (const id of Object.keys(commentAnchors(state.doc))) max = Math.max(max, +id);
        const id = String(max + 1);
        const { from, to } = state.selection;
        const op = state.schema.nodes.opaque;
        const mk = (kind, xml) => op.create({ kind, xml });
        if (dispatch) {
          // сначала конец (позиции до него не сдвигаются), потом начало
          tr.insert(to, [mk('commentRange', `<w:commentRangeEnd w:id="${id}"/>`), mk('comment', `<w:r><w:commentReference w:id="${id}"/></w:r>`)]);
          tr.insert(from, mk('commentRange', `<w:commentRangeStart w:id="${id}"/>`));
          setComments(tr, [...list, { id, author, initials, date: wordDate(), text }]);
          tr.setMeta(commentsKey, { active: id });
          dispatch(tr);
        }
        return id;
      },
      updateComment: (id, text) => ({ state, tr, dispatch }) => {
        const list = (state.doc.attrs.comments || []).map((c) => (c.id === id ? { ...c, text } : c));
        if (dispatch) dispatch(setComments(tr, list));
        return true;
      },
      // Удалить примечание: значки в тексте и запись (текст документа остаётся)
      deleteComment: (id) => ({ state, tr, dispatch }) => {
        const del = [];
        state.doc.descendants((node, pos) => {
          if (node.type.name === 'opaque' && (node.attrs.kind === 'commentRange' || node.attrs.kind === 'comment') && idOf(node.attrs.xml) === id) del.push(pos);
        });
        if (dispatch) {
          for (const pos of del.reverse()) tr.delete(pos, pos + 1);
          setComments(tr, (state.doc.attrs.comments || []).filter((c) => c.id !== id));
          tr.setMeta(commentsKey, { active: null });
          dispatch(tr);
        }
        return true;
      },
      // Перейти к примечанию: выделить фрагмент
      goToComment: (id) => ({ state, tr, dispatch, view }) => {
        const a = commentAnchors(state.doc)[id];
        if (!a) return false;
        const from = a.start ?? a.ref;
        const to = a.end ?? from;
        if (dispatch) {
          dispatch(tr.setSelection(TextSelection.create(tr.doc, from, Math.max(from, to))).setMeta(commentsKey, { active: id }).scrollIntoView());
          setTimeout(() => {
            const dom = view && view.domAtPos(from).node;
            const el = dom && (dom.nodeType === 1 ? dom : dom.parentElement);
            if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
          }, 0);
        }
        return true;
      },
    };
  },
});

// Примечания в порядке следования в документе (для панели)
export function orderedComments(state) {
  const anchors = commentAnchors(state.doc);
  const list = state.doc.attrs.comments || [];
  const at = (c) => { const a = anchors[c.id] || {}; return a.start ?? a.ref ?? Infinity; };
  return [...list].sort((x, y) => at(x) - at(y));
}

// ---------- OOXML ----------
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

// comments.xml из списка примечаний (простой текст, абзацы — по \n)
export function commentsXml(list, ns) {
  const body = list.map((c) => {
    const paras = String(c.text || '').split('\n').map((t, i) => `<w:p>${i === 0 ? '<w:r><w:annotationRef/></w:r>' : ''}${t ? `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : ''}</w:p>`).join('');
    return `<w:comment w:id="${esc(c.id)}" w:author="${esc(c.author || 'Пользователь')}" w:date="${esc(c.date || '')}" w:initials="${esc(c.initials || '')}">${paras}</w:comment>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments ${ns}>${body}</w:comments>`;
}

// Разбор comments.xml (xmldom-документ) → список примечаний
export function readComments(doc) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const out = [];
  for (const c of Array.from(doc.getElementsByTagNameNS(W, 'comment'))) {
    const get = (n) => c.getAttributeNS(W, n) || c.getAttribute('w:' + n) || '';
    const paras = Array.from(c.getElementsByTagNameNS(W, 'p')).map((p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent).join(''));
    out.push({ id: get('id'), author: get('author'), initials: get('initials'), date: get('date'), text: paras.join('\n') });
  }
  return out;
}
