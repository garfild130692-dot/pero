// Импорт .docx → JSON-документ редактора (формат ProseMirror/TipTap).
// Собственный разбор OOXML: никаких сторонних конвертеров, только распаковка zip и XML.
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { HIGHLIGHTS, headingDefaults, normalizeColor } from '../editor/font.js';
import { WORD_FALLBACK_PARA, paraBase, paraDiff, twipsToPt } from '../editor/paragraph.js';
import { CELL_SIDES, NO_BORDERS, TABLE_SIDES, TWIPS_PER_PX, borderFromOoxml } from '../editor/table.js';
import { EMU_PER_PX, bytesToBase64, mimeOf } from '../editor/image.js';
import { EMPTY_HF, readHfPart } from '../editor/page.js';
import { readComments } from '../editor/comments.js';
import { layoutAttrs, readLayoutProps } from '../editor/paraLayout.js';
import { applyTabLayer, readTabs } from '../editor/tabs.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---------- XML-хелперы ----------
const isEl = (n) => n && n.nodeType === 1;
const ln = (n) => n.localName || String(n.nodeName).split(':').pop();

function children(el, name) {
  const out = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (isEl(n) && (!name || ln(n) === name)) out.push(n);
  }
  return out;
}
const child = (el, name) => children(el, name)[0] || null;

function wAttr(el, name) {
  if (!el) return null;
  const v = el.getAttributeNS ? el.getAttributeNS(W, name) : null;
  if (v !== null && v !== '') return v;
  const v2 = el.getAttribute('w:' + name);
  return v2 === '' ? null : v2;
}

// <w:b/>, <w:b w:val="true"/> — включено; <w:b w:val="0|false|off"/> — выключено
function onOff(el) {
  if (!el) return null;
  const v = wAttr(el, 'val');
  if (v === null) return true;
  return !['0', 'false', 'off', 'none'].includes(String(v).toLowerCase());
}

const serializer = new XMLSerializer();
const xmlOf = (node) => serializer.serializeToString(node);
// Видимый текст фрагмента (w:t внутри) — для полей
function textOf(el) {
  let out = '';
  const walk = (n) => { for (const c of children(n)) { if (ln(c) === 't') out += c.textContent; else walk(c); } };
  walk(el);
  return out;
}

function parseXml(text) {
  return new DOMParser().parseFromString(text, 'text/xml');
}

// ---------- Тема оформления: шрифты заголовков (major) и основного текста (minor) ----------
function parseTheme(doc) {
  const theme = { major: null, minor: null };
  if (!doc) return theme;
  const find = (el, name) => {
    if (!el) return null;
    for (const c of children(el)) {
      if (ln(c) === name) return c;
      const deep = find(c, name);
      if (deep) return deep;
    }
    return null;
  };
  const scheme = find(doc.documentElement, 'fontScheme');
  for (const key of ['major', 'minor']) {
    const latin = child(child(scheme, key + 'Font'), 'latin');
    theme[key] = (latin && latin.getAttribute('typeface')) || null;
  }
  return theme;
}

// ---------- Стили ----------
function parseStyles(doc, theme) {
  const map = {};
  const defaults = { marks: {} };
  if (!doc) return { map, defaults, defaultPara: null };
  const root = doc.documentElement;
  let defaultPara = null;
  for (const st of children(root, 'style')) {
    const id = wAttr(st, 'styleId');
    if (!id) continue;
    const pPr = child(st, 'pPr');
    const rPr = child(st, 'rPr');
    map[id] = {
      type: wAttr(st, 'type'),
      name: (wAttr(child(st, 'name'), 'val') || '').toLowerCase(),
      displayName: wAttr(child(st, 'name'), 'val') || id,
      basedOn: wAttr(child(st, 'basedOn'), 'val'),
      jc: wAttr(child(pPr, 'jc'), 'val'),
      numPr: readNumPr(child(pPr, 'numPr')),
      outlineLvl: wAttr(child(pPr, 'outlineLvl'), 'val'),
      marks: readRunMarks(rPr, theme),
      para: readParaProps(pPr),
      layout: readLayoutProps(pPr, child, wAttr), // «не отрывать», границы, заливка…
      tabs: child(pPr, 'tabs') ? readTabs(child(pPr, 'tabs'), children, wAttr) : null, // позиции табуляции
      tblBorders: readBorders(child(child(st, 'tblPr'), 'tblBorders'), TABLE_SIDES), // у стилей таблиц
    };
    const isDefault = ['1', 'true', 'on'].includes(String(wAttr(st, 'default')).toLowerCase());
    if (map[id].type === 'paragraph' && isDefault) defaultPara = id;
  }
  const rPrDefault = child(child(child(root, 'docDefaults'), 'rPrDefault'), 'rPr');
  defaults.marks = readRunMarks(rPrDefault, theme);
  defaults.para = readParaProps(child(child(child(root, 'docDefaults'), 'pPrDefault'), 'pPr'));
  defaults.tabs = readTabs(child(child(child(child(root, 'docDefaults'), 'pPrDefault'), 'pPr'), 'tabs'), children, wAttr);
  defaults.layout = readLayoutProps(child(child(child(root, 'docDefaults'), 'pPrDefault'), 'pPr'), child, wAttr);
  return { map, defaults, defaultPara };
}

// Свойства стиля с учётом цепочки basedOn (родитель → потомок); key — 'marks' или 'para'
function resolvedStyle(styles, styleId, cache, key) {
  if (!styleId || !styles[styleId]) return {};
  const ck = key + ':' + styleId;
  if (cache[ck]) return cache[ck];
  const chain = [];
  for (let id = styleId; id && styles[id] && chain.length < 10 && !chain.includes(id); id = styles[id].basedOn) chain.push(id);
  const out = {};
  for (const id of chain.reverse()) Object.assign(out, styles[id][key]);
  cache[ck] = out;
  return out;
}
const resolvedStyleMarks = (styles, id, cache) => resolvedStyle(styles, id, cache, 'marks');
const resolvedStylePara = (styles, id, cache) => resolvedStyle(styles, id, cache, 'para');
const resolvedStyleLayout = (styles, id, cache) => resolvedStyle(styles, id, cache, 'layout');
// Позиции табуляции стиля с учётом basedOn: слои применяются от корня (clear убирает позицию родителя)
function resolvedStyleTabs(styles, styleId, base = []) {
  const chain = [];
  for (let id = styleId; id && styles[id] && chain.length < 10 && !chain.includes(id); id = styles[id].basedOn) chain.push(id);
  return chain.reverse().reduce((list, id) => applyTabLayer(list, styles[id].tabs), base);
}

// ---------- Границы таблиц и ячеек ----------
// <w:tblBorders>/<w:tcBorders> → { сторона: граница } только для заданных сторон (start/end = left/right)
function readBorders(el, sides) {
  const out = {};
  if (!el) return out;
  for (const side of sides) {
    const alt = { left: 'start', right: 'end' }[side];
    const b = child(el, side) || (alt && child(el, alt));
    if (b) out[side] = borderFromOoxml(wAttr(b, 'val'), wAttr(b, 'sz'), wAttr(b, 'color'));
  }
  return out;
}

// ---------- Свойства абзаца: интервалы и отступы ----------
// Возвращает только заданные в pPr ключи (в pt; межстрочный auto — множитель)
function readParaProps(pPr) {
  const out = {};
  if (!pPr) return out;
  const sp = child(pPr, 'spacing');
  if (sp) {
    const before = twipsToPt(wAttr(sp, 'before'));
    const after = twipsToPt(wAttr(sp, 'after'));
    // «авто»-интервалы HTML (beforeAutospacing) Word рисует как 14 pt
    if (['1', 'true', 'on'].includes(String(wAttr(sp, 'beforeAutospacing')).toLowerCase())) out.spaceBefore = 14;
    else if (before !== null) out.spaceBefore = before;
    if (['1', 'true', 'on'].includes(String(wAttr(sp, 'afterAutospacing')).toLowerCase())) out.spaceAfter = 14;
    else if (after !== null) out.spaceAfter = after;
    const line = +wAttr(sp, 'line');
    if (line > 0) {
      const rule = wAttr(sp, 'lineRule') || 'auto';
      if (rule === 'exact' || rule === 'atLeast') { out.lineRule = rule; out.lineSpacing = line / 20; }
      else { out.lineRule = 'auto'; out.lineSpacing = line / 240; }
    }
  }
  const ind = child(pPr, 'ind');
  if (ind) {
    const left = twipsToPt(wAttr(ind, 'left') ?? wAttr(ind, 'start'));
    const right = twipsToPt(wAttr(ind, 'right') ?? wAttr(ind, 'end'));
    const first = twipsToPt(wAttr(ind, 'firstLine'));
    const hanging = twipsToPt(wAttr(ind, 'hanging'));
    if (left !== null) out.indentLeft = left;
    if (right !== null) out.indentRight = right;
    if (hanging !== null) out.indentFirst = -hanging; // выступ приоритетнее отступа первой строки
    else if (first !== null) out.indentFirst = first;
  }
  return out;
}

// Уровень заголовка по стилю. Учитываем локализованные styleId (в русском Word
// заголовки часто имеют id "1", "2"…), поэтому смотрим на внутреннее имя "heading N".
function headingLevel(styles, styleId) {
  let id = styleId;
  for (let i = 0; i < 8 && id && styles[id]; i++) {
    const s = styles[id];
    const m = /^heading (\d)$/.exec(s.name);
    if (m) return Math.min(+m[1], 6);
    if (s.name === 'title') return 1;
    if (s.name === 'subtitle') return 2;
    id = s.basedOn;
  }
  const m2 = /^heading(\d)$/i.exec(styleId || '');
  if (m2) return Math.min(+m2[1], 6);
  return 0;
}

function styleProp(styles, styleId, prop) {
  let id = styleId;
  for (let i = 0; i < 8 && id && styles[id]; i++) {
    if (styles[id][prop]) return styles[id][prop];
    id = styles[id].basedOn;
  }
  return null;
}

// ---------- Нумерация ----------
function readNumPr(numPr) {
  if (!numPr) return null;
  const numId = wAttr(child(numPr, 'numId'), 'val');
  if (!numId || numId === '0') return null;
  return { numId, ilvl: +(wAttr(child(numPr, 'ilvl'), 'val') || 0) };
}

function parseNumbering(doc) {
  const abstract = {};
  const nums = {};
  if (!doc) return { abstract, nums };
  const root = doc.documentElement;
  for (const a of children(root, 'abstractNum')) {
    const levels = {};
    for (const lvl of children(a, 'lvl')) {
      const il = +(wAttr(lvl, 'ilvl') || 0);
      levels[il] = {
        fmt: wAttr(child(lvl, 'numFmt'), 'val') || 'decimal',
        start: +(wAttr(child(lvl, 'start'), 'val') || 1),
      };
    }
    abstract[wAttr(a, 'abstractNumId')] = levels;
  }
  for (const n of children(root, 'num')) {
    nums[wAttr(n, 'numId')] = wAttr(child(n, 'abstractNumId'), 'val');
  }
  return { abstract, nums };
}

function listInfo(numbering, numId, ilvl) {
  const abs = numbering.abstract[numbering.nums[numId]];
  const lvl = abs && (abs[ilvl] || abs[0]);
  const fmt = lvl ? lvl.fmt : 'bullet';
  return { type: fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList', start: lvl ? lvl.start : 1 };
}

// Первый потомок с локальным именем name на любой глубине
function findDeep(el, name) {
  for (const c of children(el)) {
    if (ln(c) === name) return c;
    const d = findDeep(c, name);
    if (d) return d;
  }
  return null;
}
function countDeep(el, name) {
  let n = 0;
  for (const c of children(el)) n += (ln(c) === name ? 1 : 0) + countDeep(c, name);
  return n;
}

// Простой рисунок: <w:drawing> с одним wp:inline/wp:anchor, внутри — ровно одна картинка pic:pic.
// Фигуры, диаграммы, надписи, группы — не рисунки в этом смысле (остаются значками).
function readPicture(drawing) {
  const frame = findDeep(drawing, 'inline') || findDeep(drawing, 'anchor');
  if (!frame) return null;
  const data = findDeep(frame, 'graphicData');
  if (!data || !/\/picture$/.test(data.getAttribute('uri') || '') || countDeep(data, 'pic') !== 1) return null;
  const blip = findDeep(data, 'blip');
  const rid = blip && (blip.getAttributeNS(R, 'embed') || blip.getAttribute('r:embed'));
  if (!rid) return null;
  const ext = child(frame, 'extent');
  const docPr = child(frame, 'docPr');
  return {
    rid,
    cx: +(ext && ext.getAttribute('cx')) || 0,
    cy: +(ext && ext.getAttribute('cy')) || 0,
    alt: (docPr && (docPr.getAttribute('descr') || '')) || '',
    floating: ln(frame) === 'anchor',
  };
}

// ---------- Связи (гиперссылки) ----------
function parseRels(doc) {
  const map = {};
  if (!doc) return map;
  for (const rel of children(doc.documentElement, 'Relationship')) {
    map[rel.getAttribute('Id')] = { target: rel.getAttribute('Target'), type: rel.getAttribute('Type') };
  }
  return map;
}

// ---------- Форматирование символов ----------
// Гарнитура из w:rFonts. Шрифт темы (asciiTheme/hAnsiTheme) приоритетнее явного — так считает Word.
// Кириллица в Word отображается шрифтом hAnsi, латиница — ascii; обычно они совпадают.
function readFont(rFonts, theme) {
  if (!rFonts) return null;
  for (const attr of ['hAnsiTheme', 'asciiTheme']) {
    const t = wAttr(rFonts, attr);
    if (t) {
      const f = t.startsWith('major') ? theme.major : t.startsWith('minor') ? theme.minor : null;
      if (f) return f;
    }
  }
  return wAttr(rFonts, 'hAnsi') || wAttr(rFonts, 'ascii') || null;
}

// Свойства символов. Для цвета/выделения null означает явное «нет» (перекрывает стиль), отсутствие ключа — «не задано».
function readRunMarks(rPr, theme = {}) {
  const marks = {};
  if (!rPr) return marks;
  const b = onOff(child(rPr, 'b'));
  if (b !== null) marks.bold = b;
  const i = onOff(child(rPr, 'i'));
  if (i !== null) marks.italic = i;
  const s = onOff(child(rPr, 'strike'));
  if (s !== null) marks.strike = s;
  const u = child(rPr, 'u');
  if (u) {
    marks.underline = onOff(u);
    const kind = wAttr(u, 'val');
    if (marks.underline && kind && kind !== 'single') marks.underlineKind = kind;
    const uc = wAttr(u, 'color');
    if (marks.underline && uc && uc.toLowerCase() !== 'auto') marks.underlineColor = normalizeColor(uc);
  }
  const caps = onOff(child(rPr, 'caps'));
  if (caps !== null) marks.caps = caps;
  const sc = onOff(child(rPr, 'smallCaps'));
  if (sc !== null) marks.smallCaps = sc;
  const sp = wAttr(child(rPr, 'spacing'), 'val');
  if (sp != null && +sp) marks.spacing = Math.round((+sp / 20) * 4) / 4; // твипы → pt (до 0,25)

  const font = readFont(child(rPr, 'rFonts'), theme);
  if (font) marks.fontFamily = font;
  const sz = +wAttr(child(rPr, 'sz'), 'val');
  if (sz > 0) marks.fontSize = sz / 2;
  const color = child(rPr, 'color');
  if (color) {
    const v = wAttr(color, 'val');
    marks.color = !v || v.toLowerCase() === 'auto' ? null : normalizeColor(v);
  }
  const hl = child(rPr, 'highlight');
  if (hl) marks.highlight = HIGHLIGHTS[wAttr(hl, 'val')] || null;
  const shd = child(rPr, 'shd');
  if (shd) {
    // заливка: обычно fill; при узоре solid цвет берётся из color
    const v = wAttr(shd, 'val');
    const fill = v === 'solid' ? wAttr(shd, 'color') : wAttr(shd, 'fill');
    marks.shading = fill && fill.toLowerCase() !== 'auto' && v !== 'nil' ? normalizeColor(fill) : null;
  }
  const va = wAttr(child(rPr, 'vertAlign'), 'val');
  if (va) marks.vertAlign = va === 'superscript' || va === 'subscript' ? va : null;
  return marks;
}

// plainLink — гиперссылка без стиля «Гиперссылка» (как пункты оглавления Word): на экране и в файле — без синего подчёркивания
function marksToArray(m, link, plainLink = false) {
  const out = [];
  if (link) out.push({ type: 'link', attrs: plainLink ? { href: link, class: 'plain' } : { href: link } });
  if (m.bold) out.push({ type: 'bold' });
  if (m.italic) out.push({ type: 'italic' });
  if (m.strike) out.push({ type: 'strike' });
  if (m.underline) out.push(m.underlineKind || m.underlineColor ? { type: 'underline', attrs: { kind: m.underlineKind || null, color: m.underlineColor || null } } : { type: 'underline' });
  if (m.fontFamily || m.fontSize || m.color || m.spacing || m.caps || m.smallCaps) {
    const attrs = { fontFamily: m.fontFamily || null, fontSize: m.fontSize || null, color: m.color || null };
    if (m.spacing) attrs.spacing = m.spacing;
    if (m.caps) attrs.caps = true;
    if (m.smallCaps) attrs.smallCaps = true;
    out.push({ type: 'textStyle', attrs });
  }
  // выделение цветом рисуется поверх заливки — если есть оба, видно выделение
  const bg = m.highlight || m.shading;
  if (bg) out.push({ type: 'highlight', attrs: { color: bg } });
  if (m.vertAlign === 'superscript') out.push({ type: 'superscript' });
  if (m.vertAlign === 'subscript') out.push({ type: 'subscript' });
  return out;
}

// ---------- Конвертер ----------
class Importer {
  constructor(ctx) {
    this.styles = ctx.styles;
    this.defaultPara = ctx.defaultPara;
    this.theme = ctx.theme;
    this.docDefaults = ctx.docDefaults;
    this.docPara = ctx.docPara;
    this.docLayout = ctx.docLayout;
    this.docTabs = ctx.docTabs;
    this.usedStyles = {};  // пользовательские стили абзацев, на которые ссылается документ
    this.rev = null;       // текущая правка (внутри w:ins / w:del)
    this.fontBase = null;  // шрифт/размер, от которых считается прямое форматирование в текущем абзаце
    this.styleCache = {};
    this.numbering = ctx.numbering;
    this.rels = ctx.rels;
    this.media = ctx.media || {}; // rId → { src: data:-URL, target: путь в пакете }
    this.warnings = {};   // что теряется при сохранении
    this.preserved = {};  // что Перо не редактирует, но сохраняет как есть
    this.bodySectPr = null;
  }

  warn(key) {
    this.warnings[key] = (this.warnings[key] || 0) + 1;
  }

  // Фрагмент, который сохраняется как есть: значок в тексте с исходным XML
  opaque(push, kind, xml, extra = {}) {
    this.preserved[kind] = (this.preserved[kind] || 0) + 1;
    push({ type: 'opaque', attrs: { kind, xml, text: '', font: null, ...extra } });
  }

  // Содержимое абзаца → массив inline-узлов
  inlines(p, paraMarks, isHeading) {
    const out = [];
    const push = (node) => {
      const last = out[out.length - 1];
      if (node.type === 'text' && last && last.type === 'text' &&
          JSON.stringify(last.marks || []) === JSON.stringify(node.marks || [])) {
        last.text += node.text;
      } else out.push(node);
    };

    const walk = (el, link) => {
      for (const n of children(el)) {
        const name = ln(n);
        if (name === 'r') this.run(n, paraMarks, link, push, isHeading);
        else if (name === 'hyperlink') {
          const rid = n.getAttributeNS(R, 'id') || n.getAttribute('r:id');
          const anchor = wAttr(n, 'anchor');
          const href = rid && this.rels[rid] ? this.rels[rid].target : anchor ? '#' + anchor : null;
          walk(n, href);
        } else if (['smartTag', 'customXml', 'dir', 'bdo'].includes(name)) walk(n, link);
        else if (['ins', 'moveTo', 'del', 'moveFrom'].includes(name)) {
          // исправления: вставленный/удалённый текст с меткой правки (автор, дата, id)
          const prev = this.rev;
          this.rev = { type: name === 'ins' || name === 'moveTo' ? 'insertion' : 'deletion', attrs: { id: wAttr(n, 'id'), author: wAttr(n, 'author') || '', date: wAttr(n, 'date') || '' } };
          walk(n, link);
          this.rev = prev;
        }
        else if (name === 'sdt') { this.warn('элементы управления содержимым'); walk(child(n, 'sdtContent'), link); }
        else if (name === 'bookmarkStart' || name === 'bookmarkEnd') this.opaque(push, 'bookmark', xmlOf(n));
        else if (name === 'commentRangeStart' || name === 'commentRangeEnd') this.opaque(push, 'commentRange', xmlOf(n));
        else if (name === 'permStart' || name === 'permEnd') this.opaque(push, 'marker', xmlOf(n));
        else if (name === 'oMath' || name === 'oMathPara') this.opaque(push, 'math', xmlOf(n));
        else if (name === 'fldSimple') this.opaque(push, 'fieldSimple', xmlOf(n), { text: textOf(n) });
        else if (name === 'AlternateContent') this.opaque(push, 'drawing', xmlOf(n));
        else if (!['pPr', 'proofErr'].includes(name)) this.opaque(push, 'other', xmlOf(n)); // неизвестное — сохраняем
      }
    };
    walk(p, null);
    return out;
  }

  run(r, paraMarks, link, push, isHeading) {
    const rPr = child(r, 'rPr');
    // символьный стиль (кроме «Гиперссылки» — её вид задаёт сама ссылка)
    const rStyle = wAttr(child(rPr, 'rStyle'), 'val');
    const charMarks = rStyle && this.styles[rStyle] && this.styles[rStyle].name !== 'hyperlink'
      ? resolvedStyleMarks(this.styles, rStyle, this.styleCache) : {};
    const m = { ...paraMarks, ...charMarks, ...readRunMarks(rPr, this.theme) };
    // совпадающее с умолчаниями (документа или стиля абзаца) — не прямое форматирование;
    // у заголовков умолчания свои (шрифт, размер, цвет нашего вида заголовка) — isHeading = уровень заголовка
    if (!isHeading) {
      const base = this.fontBase || this.docDefaults;
      if (m.fontFamily === base.fontFamily) delete m.fontFamily;
      if (m.fontSize === base.fontSize) delete m.fontSize;
    } else {
      const hd = headingDefaults(isHeading);
      if (m.fontFamily === hd.fontFamily) delete m.fontFamily;
      if (m.fontSize === hd.fontSize) delete m.fontSize;
      if (m.color && m.color.toLowerCase() === hd.color) delete m.color;
    }
    const marks = marksToArray(m, link, !!link && link.startsWith('#') && !(rStyle && this.styles[rStyle] && this.styles[rStyle].name === 'hyperlink'));
    if (this.rev) marks.push({ type: this.rev.type, attrs: this.rev.attrs });
    if (child(rPr, 'rPrChange')) this.warn('исправления форматирования');
    const text = (t) => t && push(marks.length ? { type: 'text', text: t, marks } : { type: 'text', text: t });
    // Отдельный «прогон» для сохраняемого элемента: те же свойства символов (rPr) + сам элемент
    const runXml = (n) => {
      const nr = r.cloneNode(false);
      if (rPr) nr.appendChild(rPr.cloneNode(true));
      nr.appendChild(n.cloneNode(true));
      return xmlOf(nr);
    };
    for (const n of children(r)) {
      switch (ln(n)) {
        case 't': case 'delText': text(n.textContent); break;
        case 'tab': push(marks.length ? { type: 'tab', marks } : { type: 'tab' }); break;
        case 'ptab': {
          // абсолютная табуляция (колонтитулы): выравнивание относительно полей
          const ptab = { alignment: wAttr(n, 'alignment') || 'left', relativeTo: wAttr(n, 'relativeTo') || 'margin', leader: wAttr(n, 'leader') || 'none' };
          push(marks.length ? { type: 'tab', attrs: { ptab }, marks } : { type: 'tab', attrs: { ptab } });
          break;
        }
        case 'noBreakHyphen': text('‑'); break;
        case 'softHyphen': text('\u00ad'); break;
        case 'br': {
          const type = wAttr(n, 'type');
          if (type === 'page') this.opaque(push, 'pageBreak', runXml(n));
          else if (type === 'column') this.opaque(push, 'columnBreak', runXml(n));
          else push({ type: 'hardBreak' });
          break;
        }
        case 'cr': push({ type: 'hardBreak' }); break;
        case 'drawing': {
          const pic = readPicture(n);
          const media = pic && this.media[pic.rid];
          if (media) {
            this.preserved.picture = (this.preserved.picture || 0) + 1;
            push({
              type: 'image',
              attrs: {
                src: media.src,
                width: pic.cx ? pic.cx / EMU_PER_PX : null,
                height: pic.cy ? pic.cy / EMU_PER_PX : null,
                alt: pic.alt,
                xml: runXml(n),
                target: media.target,
                floating: pic.floating,
              },
            });
          } else this.opaque(push, 'drawing', runXml(n));
          break;
        }
        case 'AlternateContent': this.opaque(push, 'drawing', runXml(n)); break;
        case 'pict': case 'object': this.opaque(push, 'object', runXml(n)); break;
        case 'footnoteReference': this.opaque(push, 'footnote', runXml(n)); break;
        case 'endnoteReference': this.opaque(push, 'endnote', runXml(n)); break;
        case 'commentReference': this.opaque(push, 'comment', runXml(n)); break;
        case 'fldChar': case 'instrText': case 'delInstrText': this.opaque(push, 'field', runXml(n)); break;
        case 'sym': {
          // символ шрифта Symbol/Wingdings: код в области F000–F0FF
          const code = parseInt(wAttr(n, 'char') || '', 16);
          this.opaque(push, 'symbol', runXml(n), { text: isFinite(code) ? String.fromCharCode(code) : '', font: wAttr(n, 'font') });
          break;
        }
        case 'rPr': case 'lastRenderedPageBreak': break;
        default: this.opaque(push, 'other', runXml(n)); // неизвестное — сохраняем как есть
      }
    }
  }

  // Определение пользовательского стиля абзаца для doc.attrs.styles (свойства — с учётом basedOn, отличия от «Обычного»)
  styleDef(id) {
    if (this.usedStyles[id]) return this.usedStyles[id];
    const marks = resolvedStyleMarks(this.styles, id, this.styleCache);
    const run = {};
    if (marks.fontFamily && marks.fontFamily !== this.docDefaults.fontFamily) run.fontFamily = marks.fontFamily;
    if (marks.fontSize && marks.fontSize !== this.docDefaults.fontSize) run.fontSize = marks.fontSize;
    if (marks.color) run.color = marks.color;
    for (const k of ['bold', 'italic', 'underline']) if (marks[k]) run[k] = true;
    const diff = paraDiff(resolvedStylePara(this.styles, id, this.styleCache), paraBase(this.docPara));
    const para = Object.fromEntries(Object.entries(diff).filter(([, v]) => v != null));
    const jc = styleProp(this.styles, id, 'jc');
    const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc] || null;
    const tabs = resolvedStyleTabs(this.styles, id, applyTabLayer([], this.docTabs));
    const def = { id, name: this.styles[id].displayName, para, run, align, origin: 'file', modified: false, ...(tabs.length ? { tabs } : {}) };
    this.usedStyles[id] = def;
    return def;
  }

  paragraph(p) {
    const pPr = child(p, 'pPr');
    if (child(pPr, 'pPrChange')) this.warn('исправления форматирования');
    const styleId = wAttr(child(pPr, 'pStyle'), 'val') || this.defaultPara;
    const level = headingLevel(this.styles, styleId);
    let numPr = readNumPr(child(pPr, 'numPr'));
    if (!numPr && styleId) {
      // нумерация задана стилем («Маркированный список 2» и т.п.) — уровень берём из имени стиля
      const sNum = styleProp(this.styles, styleId, 'numPr');
      const sm = /^list (?:bullet|number) (\d)$/.exec((this.styles[styleId] || {}).name || '');
      if (sNum) numPr = { ...sNum, ilvl: sNum.ilvl || (sm ? +sm[1] - 1 : 0) };
    }
    if (level) numPr = null;

    // Пользовательский стиль (не «Обычный», не заголовок, не список) — абзац ссылается на него (styleId),
    // свойства стиля не переносятся в прямое форматирование: изменение стиля в Пере меняет все его абзацы.
    const st = this.styles[styleId];
    const custom = !level && !numPr && st && st.type === 'paragraph' && styleId !== this.defaultPara &&
      !['list paragraph', 'title', 'subtitle'].includes(st.name);
    const def = custom ? this.styleDef(styleId) : null;

    const directJc = wAttr(child(pPr, 'jc'), 'val');
    const jc = directJc || (custom ? null : styleProp(this.styles, styleId, 'jc'));
    const textAlign = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc] || null;
    const styleMarks = custom ? {} : { ...resolvedStyleMarks(this.styles, styleId, this.styleCache) };
    this.fontBase = def ? { fontFamily: def.run.fontFamily || this.docDefaults.fontFamily, fontSize: def.run.fontSize || this.docDefaults.fontSize } : null;
    const content = this.inlines(p, styleMarks, level);
    this.fontBase = null;

    // Интервалы и отступы: у заголовка — только прямые (вид заголовка задаёт наш стиль);
    // у элемента списка отступы задаёт список, интервалы стиля (List Paragraph) не переносим;
    // у абзаца пользовательского стиля — только прямые (отличия от стиля); у обычного — стиль (с basedOn) + прямые.
    // В атрибуты попадают только отличия от базы.
    const direct = readParaProps(pPr);
    const style = custom ? {} : resolvedStylePara(this.styles, styleId, this.styleCache);
    let values;
    let base = paraBase(this.docPara, { headingLevel: level, inList: !!numPr });
    if (numPr) {
      values = { lineRule: style.lineRule, lineSpacing: style.lineSpacing, ...direct };
      delete values.indentLeft; delete values.indentRight; delete values.indentFirst;
    } else if (this.cell && !level) {
      // абзац в ячейке — как Word: умолчания документа ← стиль таблицы ← стиль абзаца ← прямые;
      // вид ячейки по умолчанию (styles.css) — как у «Сетки таблицы»: без интервалов, одинарный
      values = { ...this.docPara, ...this.cell.para, ...style, ...direct };
      base = { ...base, spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1 };
    } else values = { ...style, ...direct };
    if (def) Object.assign(base, def.para);
    const para = paraDiff(values, base);

    // Положение на странице, границы, заливка: умолчания документа ← стиль (с basedOn) ← прямые
    // Позиции табуляции: умолчания документа ← стиль ← прямые (действующий список целиком)
    const tabs = applyTabLayer(resolvedStyleTabs(this.styles, styleId, applyTabLayer([], this.docTabs)),
      child(pPr, 'tabs') ? readTabs(child(pPr, 'tabs'), children, wAttr) : null);
    if (tabs.length) para.tabs = tabs;
    const layout = layoutAttrs({ ...(this.docLayout || {}), ...resolvedStyleLayout(this.styles, styleId, this.styleCache), ...readLayoutProps(pPr, child, wAttr) }, !!level);
    const node = level
      ? { type: 'heading', attrs: { level, textAlign, ...para, ...layout, ...(styleId && (this.styles[styleId] || {}).name !== `heading ${level}` ? { fileStyle: { id: styleId, level } } : {}) }, content }
      : { type: 'paragraph', attrs: { textAlign, ...para, ...layout, ...(custom ? { styleId } : {}) }, content };
    // разрыв раздела после абзаца — исходный XML параметров раздела
    const sect = child(pPr, 'sectPr');
    if (sect) node.attrs.sectPr = xmlOf(sect);
    if (!content.length) delete node.content;
    return { node, numPr };
  }

  table(tbl) {
    const tblPr = child(tbl, 'tblPr');
    // Границы по умолчанию: стиль таблицы (с basedOn) ← прямые tblBorders; не заданное нигде — «нет» (TableNormal)
    const styleId = wAttr(child(tblPr, 'tblStyle'), 'val');
    const chain = [];
    for (let id = styleId; id && this.styles[id] && chain.length < 10 && !chain.includes(id); id = this.styles[id].basedOn) chain.push(id);
    const borders = { ...NO_BORDERS };
    for (const id of chain.reverse()) Object.assign(borders, this.styles[id].tblBorders);
    Object.assign(borders, readBorders(child(tblPr, 'tblBorders'), TABLE_SIDES));
    // Свойства абзаца из стиля таблицы («Сетка таблицы»: без интервалов, одинарный) — для абзацев её ячеек
    const tablePara = {};
    for (const id of chain) Object.assign(tablePara, this.styles[id].para);
    const jc = wAttr(child(tblPr, 'jc'), 'val');
    const align = { center: 'center', right: 'right', end: 'right' }[jc] || null;
    // Ширины столбцов из сетки таблицы (твипы → px, как хранит TipTap)
    const grid = children(child(tbl, 'tblGrid'), 'gridCol').map((g) => +wAttr(g, 'w') || 0);

    const rows = [];
    const origins = {}; // колонка → ячейка, начавшая вертикальное объединение
    for (const tr of children(tbl, 'tr')) {
      const cells = [];
      let col = 0;
      for (const tc of children(tr, 'tc')) {
        const tcPr = child(tc, 'tcPr');
        const span = +(wAttr(child(tcPr, 'gridSpan'), 'val') || 1);
        const vm = child(tcPr, 'vMerge');
        const vmVal = vm ? wAttr(vm, 'val') || 'continue' : null;
        const cellBorders = readBorders(child(tcPr, 'tcBorders'), CELL_SIDES);
        if (vmVal === 'continue' && origins[col]) {
          const o = origins[col];
          o.attrs.rowspan += 1;
          // нижнюю границу объединённой ячейки Word берёт у последней ячейки объединения
          if (cellBorders.bottom) o.attrs.borders = { ...(o.attrs.borders || {}), bottom: cellBorders.bottom };
        } else {
          const outerCell = this.cell;
          this.cell = { para: tablePara };
          let content = this.blocks(tc);
          this.cell = outerCell;
          if (!content.length) content = [{ type: 'paragraph' }];
          const widths = grid.slice(col, col + span);
          const colwidth = widths.length === span && widths.every((w) => w > 0) ? widths.map((w) => w / TWIPS_PER_PX) : null;
          const shd = child(tcPr, 'shd');
          const fill = shd && (wAttr(shd, 'val') === 'solid' ? wAttr(shd, 'color') : wAttr(shd, 'fill'));
          const va = wAttr(child(tcPr, 'vAlign'), 'val');
          const cell = {
            type: 'tableCell',
            attrs: {
              colspan: span,
              rowspan: 1,
              colwidth,
              align: null,
              borders: Object.keys(cellBorders).length ? cellBorders : null,
              background: fill && fill.toLowerCase() !== 'auto' ? normalizeColor(fill) : null,
              verticalAlign: va === 'center' || va === 'bottom' ? va : null,
            },
            content,
          };
          cells.push(cell);
          if (vmVal === 'restart') origins[col] = cell;
          else delete origins[col];
        }
        col += span;
      }
      const trPr = child(tr, 'trPr');
      const th = child(trPr, 'trHeight');
      const height = twipsToPt(wAttr(th, 'val'));
      const rowAttrs = {
        height: height || null,
        heightRule: height && wAttr(th, 'hRule') === 'exact' ? 'exact' : null,
        repeatHeader: !!onOff(child(trPr, 'tblHeader')),
        cantSplit: !!onOff(child(trPr, 'cantSplit')),
      };
      if (cells.length) rows.push({ type: 'tableRow', attrs: rowAttrs, content: cells });
    }
    return rows.length ? { type: 'table', attrs: { borders, align }, content: rows } : null;
  }

  // Блочное содержимое тела документа / ячейки: абзацы, списки, таблицы
  blocks(container) {
    const out = [];
    let stack = []; // [{list, level, numId}]

    const closeLists = () => { stack = []; };

    const addListItem = (para, numPr) => {
      const { numId, ilvl } = numPr;
      const info = listInfo(this.numbering, numId, ilvl);
      while (stack.length && stack[stack.length - 1].level > ilvl) stack.pop();
      let top = stack[stack.length - 1];
      if (top && top.level === ilvl && (top.list.type !== info.type || (ilvl === 0 && top.numId !== numId))) {
        stack.pop();
        top = stack[stack.length - 1];
      }
      if (!top || top.level < ilvl) {
        const list = { type: info.type, content: [] };
        if (info.type === 'orderedList') list.attrs = { start: info.start };
        if (top) {
          const items = top.list.content;
          items[items.length - 1].content.push(list);
        } else {
          out.push(list);
        }
        stack.push({ list, level: ilvl, numId });
        top = stack[stack.length - 1];
      }
      top.list.content.push({ type: 'listItem', content: [para] });
    };

    for (const el of children(container)) {
      const name = ln(el);
      if (name === 'p') {
        const { node, numPr } = this.paragraph(el);
        if (numPr) addListItem(node.type === 'paragraph' ? node : { ...node, type: 'paragraph' }, numPr);
        else { closeLists(); out.push(node); }
      } else if (name === 'tbl') {
        closeLists();
        const t = this.table(el);
        if (t) out.push(t);
      } else if (name === 'sdt') {
        closeLists();
        out.push(...this.blocks(child(el, 'sdtContent')));
      } else if (name === 'customXml') {
        closeLists();
        out.push(...this.blocks(el));
      } else if (name === 'sectPr') {
        this.bodySectPr = xmlOf(el); // параметры последнего раздела — сохраняем как есть
      } else if (!['tcPr', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'commentRangeStart', 'commentRangeEnd'].includes(name)) {
        this.warn('прочие элементы');
      }
    }
    return out;
  }
}

/**
 * @param {ArrayBuffer|Uint8Array} data — содержимое .docx
 * @returns {Promise<{doc: object, warnings: Record<string, number>}>}
 */
export async function importDocx(data) {
  const zip = await JSZip.loadAsync(data);
  const read = async (path) => {
    const f = zip.file(path);
    return f ? parseXml(await f.async('string')) : null;
  };
  const documentXml = await read('word/document.xml');
  if (!documentXml) throw new Error('Это не документ Word: внутри нет word/document.xml');

  const theme = parseTheme(await read('word/theme/theme1.xml'));
  const { map: styles, defaults, defaultPara } = parseStyles(await read('word/styles.xml'), theme);
  // Шрифт/размер документа: docDefaults, поверх — стиль абзаца по умолчанию («Обычный»).
  // Если не задано нигде — так же, как Word: Times New Roman 10 pt.
  const normal = { ...defaults.marks, ...resolvedStyleMarks(styles, defaultPara, {}) };
  const docDefaults = { fontFamily: normal.fontFamily || 'Times New Roman', fontSize: normal.fontSize || 10 };
  // Абзац по умолчанию: как у Word без настроек ← pPrDefault ← стиль «Обычный»
  const docPara = { ...WORD_FALLBACK_PARA, ...defaults.para, ...resolvedStylePara(styles, defaultPara, {}) };
  // Картинки документа: связь типа image → файл в пакете → data:-URL (для показа и пересохранения)
  const rels = parseRels(await read('word/_rels/document.xml.rels'));
  // Шаг позиций табуляции по умолчанию (settings.xml); нет — как Word без настроек: 720 твипов
  const settings = await read('word/settings.xml');
  const dts = settings ? +wAttr(child(settings.documentElement, 'defaultTabStop'), 'val') : NaN;
  const defaultTab = dts > 0 ? dts / 20 : 36;
  const media = {};
  for (const [rid, rel] of Object.entries(rels)) {
    if (!/\/image$/.test(rel.type || '') || !rel.target || /^[a-z]+:/i.test(rel.target)) continue; // внешние ссылки — нет
    const target = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target.replace(/^\.\//, '');
    const f = zip.file(target);
    if (f) media[rid] = { src: `data:${mimeOf(target)};base64,${bytesToBase64(await f.async('uint8array'))}`, target };
  }
  const imp = new Importer({
    media,
    styles,
    defaultPara,
    theme,
    docDefaults,
    docPara,
    docLayout: defaults.layout,
    docTabs: defaults.tabs,
    numbering: parseNumbering(await read('word/numbering.xml')),
    rels,
  });
  const body = child(documentXml.documentElement, 'body');
  let content = imp.blocks(body);
  if (!content.length) content = [{ type: 'paragraph' }];
  // Колонтитулы для показа: первые ссылки «по умолчанию» в порядке разделов (разделы без своих наследуют их).
  // Сами части колонтитулов сохраняются вместе с пакетом без изменений.
  const sects = [];
  const walkSect = (n) => { if (n.attrs && n.attrs.sectPr) sects.push(n.attrs.sectPr); (n.content || []).forEach(walkSect); };
  content.forEach(walkSect);
  if (imp.bodySectPr) sects.push(imp.bodySectPr);
  // варианты: «по умолчанию», первой страницы (w:titlePg раздела), чётных (w:evenAndOddHeaders в settings.xml)
  const hf = { ...EMPTY_HF, own: { ...EMPTY_HF.own } };
  hf.evenOdd = !!(settings && child(settings.documentElement, 'evenAndOddHeaders') && onOff(child(settings.documentElement, 'evenAndOddHeaders')));
  for (const type of ['default', 'first', 'even']) {
    for (const kind of ['header', 'footer']) {
      for (const sx of sects) {
        const ref = new RegExp(`<w:${kind}Reference\\b[^>]*>`, 'g');
        const tag = (sx.match(ref) || []).find((t) => new RegExp(`w:type="${type}"`).test(t));
        const rid = tag && (/r:id="([^"]+)"/.exec(tag) || [])[1];
        const rel = rid && rels[rid];
        if (!rel) continue;
        const part = await read(rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target);
        if (part) {
          const paras = readHfPart(part);
          if (type === 'default') hf[kind] = paras;
          else hf[type] = { ...(hf[type] || { header: null, footer: null }), [kind]: paras };
        }
        break;
      }
    }
  }
  if (hf.header || hf.footer || hf.first || hf.even) imp.preserved.header = 1;
  // колонки текста (w:cols w:num > 1) сохраняются в sectPr, но на экране и при печати — одной колонкой
  if (sects.some((x) => +((/<w:cols\b[^>]*w:num="(\d+)"/.exec(x || '') || [])[1] || 1) > 1)) imp.preserved.columns = 1;
  // тексты примечаний (связь типа comments)
  const cRel = Object.values(rels).find((r) => /\/comments$/.test(r.type || ''));
  const cPart = cRel && await read(cRel.target.startsWith('/') ? cRel.target.slice(1) : 'word/' + cRel.target);
  const commentList = cPart ? readComments(cPart) : [];
  return {
    doc: {
      type: 'doc',
      attrs: { ...docDefaults, para: docPara, defaultTab, sectPr: imp.bodySectPr, hf, styles: Object.values(imp.usedStyles), comments: commentList },
      content,
    },
    warnings: imp.warnings,
    preserved: imp.preserved,
  };
}
