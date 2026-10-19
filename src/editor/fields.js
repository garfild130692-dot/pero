// Поля, закладки, перекрёстные ссылки и оглавление — как в Word.
//
// Сложное поле в модели — значки opaque kind 'field' (fldChar begin / instrText / fldChar separate / fldChar end),
// результат между separate и end — обычный текст (импорт так и делает). Закладка — пара значков kind 'bookmark'
// (w:bookmarkStart w:name w:id / w:bookmarkEnd w:id). Обновление поля (F9) заменяет текст результата: PAGE, NUMPAGES,
// DATE/TIME (+ \@ «формат»), REF/PAGEREF (закладка), TOC — оглавление целиком (абзацы между началом и концом поля).
// Номера страниц — из разбиения на экране (pagination.js), поэтому после изменения оглавления его обновляют ещё раз.
import { Extension } from '@tiptap/core';
import { Fragment, Slice } from '@tiptap/pm/model';
import { Plugin, TextSelection } from '@tiptap/pm/state';
import { paginationKey } from './pagination.js';
import { parseSectPr, sectionAt } from './page.js';

// символы, недопустимые в XML 1.0, выбрасываются (как в export.js) — иначе текст из чужого файла ломает часть
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// ---------- Разбор значков ----------
const isOpaque = (n, kind) => n && n.type.name === 'opaque' && n.attrs.kind === kind;

// Часть сложного поля: { char: 'begin'|'separate'|'end' } или { instr }
export function fieldPart(node) {
  if (!isOpaque(node, 'field')) return null;
  const xml = node.attrs.xml || '';
  const ch = /fldCharType="(begin|separate|end)"/.exec(xml);
  if (ch) return { char: ch[1] };
  const m = /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/.exec(xml) || /<w:instrText\b[^>]*\/>/.exec(xml);
  return m ? { instr: unesc(m[1] || '') } : null;
}

// Все поля документа: сложные { instr, begin, sep, end, depth } и простые { simple, pos, instr }
export function scanFields(doc) {
  const out = [];
  const stack = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'opaque') {
      if (node.attrs.kind === 'fieldSimple') {
        const m = /w:instr="([^"]*)"/.exec(node.attrs.xml || '');
        out.push({ simple: true, pos, instr: unesc(m ? m[1] : ''), node });
        return false;
      }
      const p = fieldPart(node);
      if (!p) return false;
      if (p.char === 'begin') stack.push({ instr: '', begin: pos, sep: null, end: null, depth: stack.length });
      else if (p.instr != null && stack.length && stack[stack.length - 1].sep == null) stack[stack.length - 1].instr += p.instr;
      else if (p.char === 'separate' && stack.length) stack[stack.length - 1].sep = pos;
      else if (p.char === 'end' && stack.length) { const f = stack.pop(); f.end = pos; out.push(f); }
      return false;
    }
    return true;
  });
  return out.sort((a, b) => (a.begin ?? a.pos) - (b.begin ?? b.pos));
}

export const fieldName = (instr) => String(instr || '').trim().split(/\s+/)[0].toUpperCase();

// Закладки: имя → { id, start, end } (позиции значков начала и конца)
export function bookmarks(doc) {
  const byId = new Map();
  const out = new Map();
  doc.descendants((node, pos) => {
    if (!isOpaque(node, 'bookmark')) return true;
    const xml = node.attrs.xml || '';
    const id = (/w:id="([^"]*)"/.exec(xml) || [])[1];
    if (/bookmarkStart/.test(xml)) {
      const name = unesc((/w:name="([^"]*)"/.exec(xml) || [])[1] || '');
      const b = { id, name, start: pos, end: pos };
      byId.set(id, b);
      if (name && !out.has(name)) out.set(name, b);
    } else if (byId.has(id)) byId.get(id).end = pos;
    return false;
  });
  return out;
}

export const BOOKMARK_NAME = /^[\p{L}][\p{L}\p{N}_]{0,39}$/u; // как в Word: с буквы, без пробелов, до 40 знаков

function nextBookmarkId(doc) {
  let max = 0;
  doc.descendants((node) => {
    if (isOpaque(node, 'bookmark')) { const id = +((/w:id="(\d+)"/.exec(node.attrs.xml) || [])[1] || 0); if (id > max) max = id; }
    return node.isBlock;
  });
  return max + 1;
}

export const bookmarkNodes = (schema, name, id) => [
  schema.nodes.opaque.create({ kind: 'bookmark', xml: `<w:bookmarkStart w:id="${id}" w:name="${esc(name)}"/>` }),
  schema.nodes.opaque.create({ kind: 'bookmark', xml: `<w:bookmarkEnd w:id="${id}"/>` }),
];

// Узлы сложного поля с результатом text
export function fieldNodes(schema, instr, text, marks = []) {
  const o = (xml) => schema.nodes.opaque.create({ kind: 'field', xml }, null, marks);
  const nodes = [
    o('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'),
    o(`<w:r><w:instrText xml:space="preserve"> ${esc(instr.trim()).replace(/&quot;/g, '"')} </w:instrText></w:r>`),
    o('<w:r><w:fldChar w:fldCharType="separate"/></w:r>'),
  ];
  if (text) nodes.push(schema.text(text, marks));
  nodes.push(o('<w:r><w:fldChar w:fldCharType="end"/></w:r>'));
  return nodes;
}

// ---------- Значения полей ----------
const MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
const MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
const DAYS = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
const DAYS_SHORT = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
const pad = (n, w = 2) => String(n).padStart(w, '0');

// Формат даты Word (\@ "dd.MM.yyyy"): d dd ddd dddd M MM MMM MMMM yy yyyy H HH h hh m mm s ss AM/PM 'текст'
export function formatDate(date, pic) {
  const genitive = /d/.test(pic.replace(/'[^']*'/g, '')); // «11 октября» — с числом месяц в родительном
  return pic.replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm/g, (t, lit) => {
    if (lit !== undefined) return lit;
    const h12 = date.getHours() % 12 || 12;
    switch (t) {
      case 'yyyy': return String(date.getFullYear());
      case 'yy': return pad(date.getFullYear() % 100);
      case 'MMMM': return (genitive ? MONTHS_GEN : MONTHS)[date.getMonth()];
      case 'MMM': return MONTHS[date.getMonth()].slice(0, 3);
      case 'MM': return pad(date.getMonth() + 1);
      case 'M': return String(date.getMonth() + 1);
      case 'dddd': return DAYS[date.getDay()];
      case 'ddd': return DAYS_SHORT[date.getDay()];
      case 'dd': return pad(date.getDate());
      case 'd': return String(date.getDate());
      case 'HH': return pad(date.getHours());
      case 'H': return String(date.getHours());
      case 'hh': return pad(h12);
      case 'h': return String(h12);
      case 'mm': return pad(date.getMinutes());
      case 'm': return String(date.getMinutes());
      case 'ss': return pad(date.getSeconds());
      case 's': return String(date.getSeconds());
      default: return date.getHours() < 12 ? (t === 'AM/PM' ? 'AM' : 'am') : (t === 'AM/PM' ? 'PM' : 'pm');
    }
  });
}

// Форматы «Вставка → Дата и время» (как в русском Word)
export const DATE_FORMATS = ['dd.MM.yyyy', 'dddd, d MMMM yyyy \'г.\'', 'd MMMM yyyy \'г.\'', 'dd.MM.yy', 'yyyy-MM-dd', 'd-MMM-yy',
  'dd.MM.yyyy H:mm', 'dd.MM.yyyy H:mm:ss', 'H:mm', 'H:mm:ss', 'MMMM yy', 'MMM-yy'];

const switchArg = (instr, sw) => {
  const m = new RegExp(`\\\\${sw}\\s+(?:"([^"]*)"|(\\S+))`).exec(instr);
  return m ? (m[1] ?? m[2]) : null;
};
const fieldArg = (instr) => {
  const m = /^\s*\S+\s+(?:"([^"]*)"|([^\\\s]\S*))/.exec(instr);
  return m ? (m[1] ?? m[2]) : null;
};

const toRoman = (n) => {
  const R = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of R) while (n >= v) { s += r; n -= v; }
  return s;
};
// \* Arabic / ROMAN / roman / ALPHABETIC / alphabetic для чисел
function numFormat(n, instr) {
  const f = switchArg(instr, '\\*');
  if (f === 'ROMAN') return toRoman(n).toUpperCase();
  if (f === 'roman') return toRoman(n);
  if (f === 'ALPHABETIC' || f === 'alphabetic') {
    const s = String.fromCharCode(64 + ((n - 1) % 26) + 1).repeat(Math.floor((n - 1) / 26) + 1);
    return f === 'ALPHABETIC' ? s : s.toLowerCase();
  }
  return String(n);
}

export const REF_ERROR = 'Ошибка! Источник ссылки не найден.';

// Номер страницы позиции — по разбиению на экране
export function pageAt(state, pos) {
  const p = paginationKey.getState(state);
  if (!p) return 1;
  return 1 + p.breaks.filter((b) => b <= pos).length;
}

// Значение поля или null (не умеем — оставить как есть). ctx: { state, marks: bookmarks, now }
export function fieldValue(instr, pos, ctx) {
  const name = fieldName(instr);
  const p = paginationKey.getState(ctx.state);
  switch (name) {
    case 'PAGE': return numFormat(pageAt(ctx.state, pos), instr);
    case 'NUMPAGES': return numFormat(p ? p.total : 1, instr);
    case 'DATE': case 'TIME':
      return formatDate(ctx.now, switchArg(instr, '@') || (name === 'TIME' ? 'H:mm' : 'dd.MM.yyyy'));
    case 'CREATEDATE': case 'SAVEDATE': {
      // даты из свойств документа (docProps/core.xml); нет — не трогаем
      const v = (ctx.state.doc.attrs.props || {})[name === 'CREATEDATE' ? 'created' : 'modified'];
      const d = v ? new Date(v) : null;
      return d && !Number.isNaN(+d) ? formatDate(d, switchArg(instr, '@') || 'dd.MM.yyyy H:mm') : null;
    }
    case 'PRINTDATE': return null;
    case 'AUTHOR': case 'TITLE': case 'SUBJECT': case 'KEYWORDS': case 'COMMENTS': case 'LASTSAVEDBY': {
      const k = { AUTHOR: 'author', TITLE: 'title', SUBJECT: 'subject', KEYWORDS: 'keywords', COMMENTS: 'comments', LASTSAVEDBY: 'lastModifiedBy' }[name];
      return (ctx.state.doc.attrs.props || {})[k] || '';
    }
    case 'DOCPROPERTY': {
      const k = { title: 'title', subject: 'subject', author: 'author', keywords: 'keywords', comments: 'comments' }[String(fieldArg(instr) || '').toLowerCase()];
      return k ? (ctx.state.doc.attrs.props || {})[k] || '' : null;
    }
    case 'FILENAME': {
      const f = ctx.file || {};
      if (!f.name) return null;
      return /\\p\b/.test(instr) && f.path ? f.path : f.name;
    }
    case 'NUMWORDS': return String((ctx.state.doc.textBetween(0, ctx.state.doc.content.size, ' ', ' ').match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) || []).length);
    case 'NUMCHARS': return String(ctx.state.doc.textBetween(0, ctx.state.doc.content.size, ' ', ' ').replace(/\s/g, '').length);
    case 'REF': case 'PAGEREF': {
      const b = ctx.marks.get(fieldArg(instr));
      if (!b) return REF_ERROR;
      if (name === 'PAGEREF') {
        if (/\\p\b/.test(instr)) return pageAt(ctx.state, b.start) === pageAt(ctx.state, pos) ? 'на этой странице' : b.start < pos ? 'выше' : 'ниже';
        return numFormat(pageAt(ctx.state, b.start), instr);
      }
      if (/\\p\b/.test(instr)) return b.start < pos ? 'выше' : 'ниже';
      return ctx.state.doc.textBetween(b.start, b.end, ' ', '').replace(/\s+$/, '');
    }
    default: return null;
  }
}

// ---------- Оглавление ----------
// Параметры поля TOC: уровни \o "1-3", гиперссылки \h
export function tocOptions(instr) {
  const o = switchArg(instr, 'o');
  const m = o && /(\d)\s*-\s*(\d)/.exec(o);
  return { from: m ? +m[1] : 1, to: m ? +m[2] : 3, links: /\\h\b/.test(instr) };
}

// Стили «toc N» / «TOC Heading»: из документа (по имени) или новые
function tocStyle(styles, name, make) {
  const s = styles.find((x) => String(x.name || '').toLowerCase() === name.toLowerCase());
  if (s) return { styles, id: s.id };
  const def = make();
  return { styles: [...styles, def], id: def.id };
}

// Абзацы оглавления для заголовков heads [{ text, level, anchor, page }]
function tocParagraphs(state, instr, heads, styles, width) {
  const { schema } = state;
  const opt = tocOptions(instr);
  // как у стилей «toc N» Word: правая табуляция чуть левее поля (9345 твипов при ширине текста 9355)
  const right = [{ pos: Math.floor(width * 4) / 4 - 0.5, align: 'right', leader: 'dot' }];
  let st = styles;
  const idOf = (level) => {
    const r = tocStyle(st, `toc ${level}`, () => ({ id: `TOC${level}`, name: `toc ${level}`, para: { spaceAfter: 5, indentLeft: (level - 1) * 11 }, tabs: right, origin: 'pero' }));
    st = r.styles;
    return r.id;
  };
  const begin = fieldNodes(schema, instr, '').slice(0, 3);
  const end = fieldNodes(schema, instr, '').slice(-1);
  const paras = [];
  if (!heads.length) {
    paras.push(schema.nodes.paragraph.create(null, [...begin, schema.text('Элементы оглавления не найдены.'), ...end]));
    return { paras, styles: st };
  }
  heads.forEach((h, i) => {
    const link = opt.links ? [schema.marks.link.create({ href: `#${h.anchor}`, class: 'plain' })] : [];
    const content = [];
    if (i === 0) content.push(...begin);
    if (h.text) content.push(schema.text(h.text, link));
    content.push(schema.nodes.tab.create(null, null, link));
    content.push(...fieldNodes(schema, `PAGEREF ${h.anchor} \\h`, String(h.page), link));
    if (i === heads.length - 1) content.push(...end);
    paras.push(schema.nodes.paragraph.create({ styleId: idOf(h.level), tabs: right }, content));
  });
  return { paras, styles: st };
}

const randDigits = (n) => Array.from({ length: n }, (_, i) => (i === 0 ? 1 + Math.floor(Math.random() * 9) : Math.floor(Math.random() * 10))).join('');

// Заголовки для оглавления + закладки _Toc у тех, где их нет (вставляются в tr; позиции — исходного state).
// skip — [from, to] самого оглавления; cache — общий для нескольких оглавлений в одном tr
function tocCache(state) {
  const anchorAt = new Map(); // позиция заголовка → имя _Toc закладки
  for (const [name, b] of bookmarks(state.doc)) if (/^_Toc/.test(name)) anchorAt.set(state.doc.resolve(b.start).before(), name);
  return { anchorAt, id: nextBookmarkId(state.doc) };
}
function collectHeads(state, tr, opt, skip, cache = tocCache(state)) {
  const heads = [];
  const { anchorAt } = cache;
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'heading') return node.isBlock && node.type.name !== 'paragraph';
    if (skip && pos >= skip[0] && pos < skip[1]) return false;
    const level = node.attrs.level;
    const text = node.textBetween(0, node.content.size, ' ', '').replace(/\t/g, ' ').trim();
    if (level < opt.from || level > opt.to || !text) return false;
    let anchor = anchorAt.get(pos);
    if (!anchor) {
      anchor = `_Toc${randDigits(9)}`;
      anchorAt.set(pos, anchor);
      const [s, e] = bookmarkNodes(state.schema, anchor, cache.id++);
      tr.insert(tr.mapping.map(pos + node.content.size + 1), e);
      tr.insert(tr.mapping.map(pos + 1), s);
    }
    heads.push({ text, level, anchor, page: pageAt(state, pos + 1) });
    return false;
  });
  return heads;
}

const textWidthPt = (state) => {
  const g = parseSectPr(sectionAt(state).sectPr);
  return (g.w - g.left - g.right) / 20;
};

// Заменить поле TOC f (из scanFields) новым оглавлением
function rebuildToc(state, tr, f, cache) {
  const opt = tocOptions(f.instr);
  const $b = state.doc.resolve(f.begin);
  const $e = state.doc.resolve(f.end);
  const from = $b.parentOffset === 0 ? $b.before() : f.begin;
  const to = $e.parentOffset + 1 === $e.parent.content.size ? $e.after() : f.end + 1;
  const heads = collectHeads(state, tr, opt, [from, to], cache);
  const { paras, styles } = tocParagraphs(state, f.instr, heads, tr.doc.attrs.styles || [], textWidthPt(state));
  const a = tr.mapping.map(from);
  const b = tr.mapping.map(to);
  if (from === $b.before() && to === $e.after()) tr.replaceWith(a, b, paras);
  else tr.replace(a, b, new Slice(Fragment.from(paras), from === $b.before() ? 0 : 1, to === $e.after() ? 0 : 1));
  if (styles !== (tr.doc.attrs.styles || [])) tr.setDocAttribute('styles', styles);
}

// Обновить поля в [from, to] (null — во всём документе). Возвращает tr или null
export function updateFieldsTr(state, range = null, now = new Date(), file = null) {
  const fields = scanFields(state.doc);
  const ctx = { state, marks: bookmarks(state.doc), now, file };
  const inRange = (f) => {
    if (!range) return true;
    const a = f.begin ?? f.pos; const b = f.end ?? f.pos;
    return b >= range[0] && a <= range[1];
  };
  const tr = state.tr;
  // оглавления — целиком; поля внутри них (PAGEREF) пересоздаются вместе с оглавлением
  const tocs = fields.filter((f) => !f.simple && fieldName(f.instr) === 'TOC' && inRange(f));
  const insideToc = (f) => tocs.some((t) => (f.begin ?? f.pos) > t.begin && (f.end ?? f.pos) < t.end);
  const simple = fields.filter((f) => !(fieldName(f.instr) === 'TOC') && inRange(f) && !insideToc(f) && (f.simple || (f.sep != null && f.depth === 0)));
  // сначала оглавления (с конца документа), потом прочие поля; все позиции — исходного state через tr.mapping
  const cache = tocCache(state);
  for (const f of [...tocs].reverse()) rebuildToc(state, tr, f, cache);
  const edits = [];
  for (const f of simple) {
    const v = fieldValue(f.instr, f.begin ?? f.pos, ctx);
    if (v != null) edits.push([f, v]);
  }
  for (const [f, v] of edits.reverse()) {
    if (f.simple) {
      const xml = f.node.attrs.xml.replace(/(<w:fldSimple\b[^>]*>)[\s\S]*(<\/w:fldSimple>)/, (m0, a, b) => {
        const rPr = (/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m0) || [''])[0];
        return `${a}<w:r>${rPr}<w:t xml:space="preserve">${esc(v)}</w:t></w:r>${b}`;
      });
      if (xml !== f.node.attrs.xml || f.node.attrs.text !== v) tr.setNodeMarkup(tr.mapping.map(f.pos), null, { ...f.node.attrs, xml, text: v });
      continue;
    }
    const a = f.sep + 1; const b = f.end;
    if (state.doc.resolve(a).parent !== state.doc.resolve(b).parent) continue; // результат на несколько абзацев — не трогаем
    if (state.doc.textBetween(a, b, '\n', '') === v) continue;
    // оформление результата — как у первого символа результата (или у начала поля)
    let marks = null;
    state.doc.nodesBetween(a, b, (n) => { if (!marks && n.isText) marks = n.marks; });
    marks = marks || state.doc.nodeAt(f.begin).marks;
    const ma = tr.mapping.map(a); const mb = tr.mapping.map(b);
    if (v) tr.replaceWith(ma, mb, state.schema.text(v, marks)); else tr.delete(ma, mb);
  }
  return tr.docChanged ? tr : null;
}

// Вставить оглавление на месте текущего абзаца (пустой абзац — заменяется)
export function insertTocTr(state, instr = 'TOC \\o "1-3" \\h \\z \\u') {
  const { schema } = state;
  const tr = state.tr;
  const $f = state.selection.$from;
  let depth = $f.depth;
  while (depth > 0 && !$f.node(depth).isTextblock) depth--;
  const block = $f.node(depth);
  const at = depth > 0 ? $f.before(depth) : $f.pos;
  const opt = tocOptions(instr);
  const heads = collectHeads(state, tr, opt, null);
  let styles = state.doc.attrs.styles || [];
  const th = tocStyle(styles, 'TOC Heading', () => ({ id: 'TOCHeading', name: 'TOC Heading', para: { spaceBefore: 12, spaceAfter: 6 }, run: { fontFamily: 'Calibri Light', fontSize: 16, color: '#2f5496' }, origin: 'pero' }));
  styles = th.styles;
  const r = tocParagraphs(state, instr, heads, styles, textWidthPt(state));
  styles = r.styles;
  const title = schema.nodes.paragraph.create({ styleId: th.id }, schema.text('Оглавление'));
  const nodes = [title, ...r.paras];
  const empty = block && block.isTextblock && block.content.size === 0 && depth > 0;
  const pos = tr.mapping.map(at);
  if (empty) tr.replaceWith(pos, tr.mapping.map(at + block.nodeSize), nodes);
  else tr.insert(pos, nodes);
  tr.setDocAttribute('styles', styles);
  tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1)));
  return tr;
}

// Удалить оглавление под курсором (или первое в документе)
export function removeTocTr(state) {
  const tocs = scanFields(state.doc).filter((f) => !f.simple && fieldName(f.instr) === 'TOC');
  if (!tocs.length) return null;
  const at = state.selection.from;
  const f = tocs.find((t) => at >= t.begin && at <= t.end) || tocs[0];
  const $b = state.doc.resolve(f.begin); const $e = state.doc.resolve(f.end);
  const tr = state.tr.delete($b.before(), $e.after());
  // заголовок «Оглавление» перед ним (стиль TOC Heading) — тоже
  const prev = $b.before() > 0 ? state.doc.resolve($b.before()).nodeBefore : null;
  const st = prev && prev.attrs && prev.attrs.styleId && (state.doc.attrs.styles || []).find((s) => s.id === prev.attrs.styleId);
  if (st && /^toc heading$/i.test(st.name || '')) tr.delete($b.before() - prev.nodeSize, $b.before());
  return tr;
}

// ---------- Закладки и перекрёстные ссылки ----------
export function addBookmarkTr(state, name) {
  const marks = bookmarks(state.doc);
  const tr = state.tr;
  const old = marks.get(name);
  if (old) { // то же имя — переносится (как в Word)
    tr.delete(old.end, old.end + 1);
    tr.delete(old.start, old.start + 1);
  }
  const id = old ? old.id : nextBookmarkId(state.doc);
  const [s, e] = bookmarkNodes(state.schema, name, id);
  const { from, to } = state.selection;
  tr.insert(tr.mapping.map(to, 1), e);
  tr.insert(tr.mapping.map(from, -1), s);
  return tr;
}

export function removeBookmarkTr(state, name) {
  const b = bookmarks(state.doc).get(name);
  if (!b) return null;
  const tr = state.tr;
  if (b.end !== b.start) tr.delete(b.end, b.end + 1);
  tr.delete(b.start, b.start + 1);
  return tr;
}

// Заголовки документа для перекрёстных ссылок: [{ pos, level, text, ref (закладка _Ref, если уже есть) }]
export function headingList(doc) {
  const marks = bookmarks(doc);
  const refAt = new Map();
  for (const [name, b] of marks) if (/^_Ref/.test(name)) refAt.set(doc.resolve(b.start).before(), name);
  const out = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'heading') return node.isBlock && node.type.name !== 'paragraph';
    const text = node.textBetween(0, node.content.size, ' ', '').trim();
    if (text) out.push({ pos, level: node.attrs.level, text, ref: refAt.get(pos) || null });
    return false;
  });
  return out;
}

// Вставить перекрёстную ссылку. target: { bookmark } | { heading: pos }; what: 'text' | 'page' | 'above'; link — \h
export function insertCrossRefTr(state, target, what, link) {
  const tr = state.tr;
  let name = target.bookmark;
  if (target.heading != null) {
    const h = headingList(state.doc).find((x) => x.pos === target.heading);
    if (!h) return null;
    name = h.ref;
    if (!name) {
      // закладка _Ref вокруг текста заголовка, как делает Word
      name = `_Ref${randDigits(9)}`;
      const node = state.doc.nodeAt(h.pos);
      const [s, e] = bookmarkNodes(state.schema, name, nextBookmarkId(state.doc));
      tr.insert(h.pos + node.content.size + 1, e);
      tr.insert(h.pos + 1, s);
    }
  }
  const instr = `${what === 'page' ? 'PAGEREF' : 'REF'} ${name}${what === 'above' ? ' \\p' : ''}${link ? ' \\h' : ''}`;
  const st2 = state.apply(tr);
  const at = tr.mapping.map(state.selection.from);
  const value = fieldValue(instr, at, { state: st2, marks: bookmarks(st2.doc), now: new Date() }) || '';
  const marks = state.storedMarks || state.selection.$from.marks();
  const nodes = fieldNodes(state.schema, instr, value, marks.filter((m) => m.type.name !== 'link'));
  tr.replaceWith(at, tr.mapping.map(state.selection.to), nodes);
  return tr;
}

// Поле «Дата и время»: fixed — просто текст
export function insertDateTr(state, pic, asField, now = new Date()) {
  const text = formatDate(now, pic);
  const marks = state.storedMarks || state.selection.$from.marks();
  const timeOnly = !/[dMy]/.test(pic.replace(/'[^']*'/g, ''));
  const nodes = asField ? fieldNodes(state.schema, `${timeOnly ? 'TIME' : 'DATE'} \\@ "${pic}"`, text, marks) : [state.schema.text(text, marks)];
  return state.tr.replaceWith(state.selection.from, state.selection.to, nodes);
}

export const Fields = Extension.create({
  name: 'fields',
  addStorage() {
    return { file: null }; // { name, path } открытого файла — для поля FILENAME (задаёт App)
  },
  addCommands() {
    return {
      // F9: поля в выделении (без выделения — поле под курсором; Ctrl+A, F9 — все)
      updateFields: (all = false) => ({ state, dispatch }) => {
        const { from, to } = state.selection;
        const tr = updateFieldsTr(state, all ? null : [from, to], new Date(), this.storage.file);
        if (!tr) return false;
        if (dispatch) dispatch(tr);
        return true;
      },
      updateToc: () => ({ state, dispatch }) => {
        const tocs = scanFields(state.doc).filter((f) => !f.simple && fieldName(f.instr) === 'TOC');
        if (!tocs.length) return false;
        const tr = updateFieldsTr(state, [tocs[0].begin, tocs[tocs.length - 1].end]);
        if (!tr) return false;
        if (dispatch) dispatch(tr);
        return true;
      },
      insertToc: (instr) => ({ state, dispatch }) => { if (dispatch) dispatch(insertTocTr(state, instr).scrollIntoView()); return true; },
      removeToc: () => ({ state, dispatch }) => { const tr = removeTocTr(state); if (!tr) return false; if (dispatch) dispatch(tr); return true; },
      addBookmark: (name) => ({ state, dispatch }) => { if (dispatch) dispatch(addBookmarkTr(state, name)); return true; },
      removeBookmark: (name) => ({ state, dispatch }) => { const tr = removeBookmarkTr(state, name); if (!tr) return false; if (dispatch) dispatch(tr); return true; },
      goToBookmark: (name) => ({ state, dispatch }) => {
        const b = bookmarks(state.doc).get(name);
        if (!b) return false;
        if (dispatch) dispatch(state.tr.setSelection(TextSelection.create(state.doc, b.start + 1, Math.max(b.start + 1, b.end))).scrollIntoView());
        return true;
      },
      insertCrossRef: (target, what, link) => ({ state, dispatch }) => {
        const tr = insertCrossRefTr(state, target, what, link);
        if (!tr) return false;
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
      insertDate: (pic, asField) => ({ state, dispatch }) => { if (dispatch) dispatch(insertDateTr(state, pic, asField).scrollIntoView()); return true; },
    };
  },
  addProseMirrorPlugins() {
    // Ctrl+щелчок по ссылке на закладку (#имя — пункты оглавления, перекрёстные ссылки) — переход к ней
    return [new Plugin({
      props: {
        handleClick: (view, pos, event) => {
          if (!(event.ctrlKey || event.metaKey)) return false;
          const a = event.target.closest && event.target.closest('a[href^="#"]');
          if (!a) return false;
          return this.editor.commands.goToBookmark(decodeURIComponent(a.getAttribute('href').slice(1)));
        },
      },
    })];
  },
  addKeyboardShortcuts() {
    return { F9: () => { this.editor.commands.updateFields(this.editor.state.selection.from === 0 || isAll(this.editor.state)); return true; } };
  },
});

const isAll = (state) => state.selection.from <= 1 && state.selection.to >= state.doc.content.size - 1;
