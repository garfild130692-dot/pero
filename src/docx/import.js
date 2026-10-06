// Импорт .docx → JSON-документ редактора (формат ProseMirror/TipTap).
// Собственный разбор OOXML: никаких сторонних конвертеров, только распаковка zip и XML.
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { HIGHLIGHTS, normalizeColor } from '../editor/font.js';

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
      basedOn: wAttr(child(st, 'basedOn'), 'val'),
      jc: wAttr(child(pPr, 'jc'), 'val'),
      numPr: readNumPr(child(pPr, 'numPr')),
      outlineLvl: wAttr(child(pPr, 'outlineLvl'), 'val'),
      marks: readRunMarks(rPr, theme),
    };
    const isDefault = ['1', 'true', 'on'].includes(String(wAttr(st, 'default')).toLowerCase());
    if (map[id].type === 'paragraph' && isDefault) defaultPara = id;
  }
  const rPrDefault = child(child(child(root, 'docDefaults'), 'rPrDefault'), 'rPr');
  defaults.marks = readRunMarks(rPrDefault, theme);
  return { map, defaults, defaultPara };
}

// Свойства символов стиля с учётом цепочки basedOn (родитель → потомок)
function resolvedStyleMarks(styles, styleId, cache) {
  if (!styleId || !styles[styleId]) return {};
  if (cache[styleId]) return cache[styleId];
  const chain = [];
  for (let id = styleId; id && styles[id] && chain.length < 10 && !chain.includes(id); id = styles[id].basedOn) chain.push(id);
  const marks = {};
  for (const id of chain.reverse()) Object.assign(marks, styles[id].marks);
  cache[styleId] = marks;
  return marks;
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
  if (u) marks.underline = onOff(u);

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

function marksToArray(m, link) {
  const out = [];
  if (link) out.push({ type: 'link', attrs: { href: link } });
  if (m.bold) out.push({ type: 'bold' });
  if (m.italic) out.push({ type: 'italic' });
  if (m.strike) out.push({ type: 'strike' });
  if (m.underline) out.push({ type: 'underline' });
  if (m.fontFamily || m.fontSize || m.color) {
    out.push({ type: 'textStyle', attrs: { fontFamily: m.fontFamily || null, fontSize: m.fontSize || null, color: m.color || null } });
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
    this.styleCache = {};
    this.numbering = ctx.numbering;
    this.rels = ctx.rels;
    this.warnings = {};
  }

  warn(key) {
    this.warnings[key] = (this.warnings[key] || 0) + 1;
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
        } else if (['ins', 'smartTag', 'fldSimple', 'customXml', 'dir', 'bdo'].includes(name)) walk(n, link);
        else if (name === 'sdt') walk(child(n, 'sdtContent'), link);
        else if (name === 'oMath' || name === 'oMathPara') this.warn('формулы');
        else if (name === 'del') this.warn('исправления (удалённый текст)');
        // pPr, bookmarkStart/End, proofErr, commentRangeStart и т.п. — пропускаем
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
    // совпадающее с умолчаниями документа — не прямое форматирование (у заголовков свои шрифт и размер)
    if (!isHeading) {
      if (m.fontFamily === this.docDefaults.fontFamily) delete m.fontFamily;
      if (m.fontSize === this.docDefaults.fontSize) delete m.fontSize;
    }
    const marks = marksToArray(m, link);
    const text = (t) => t && push(marks.length ? { type: 'text', text: t, marks } : { type: 'text', text: t });
    for (const n of children(r)) {
      switch (ln(n)) {
        case 't': text(n.textContent); break;
        case 'tab': text('\t'); break;
        case 'noBreakHyphen': text('‑'); break;
        case 'br':
          if (wAttr(n, 'type') === 'page') this.warn('разрывы страниц');
          push({ type: 'hardBreak' });
          break;
        case 'cr': push({ type: 'hardBreak' }); break;
        case 'drawing': case 'pict': case 'object': this.warn('изображения и объекты'); break;
        case 'footnoteReference': case 'endnoteReference': this.warn('сноски'); break;
        case 'commentReference': this.warn('комментарии'); break;
        default: break;
      }
    }
  }

  paragraph(p) {
    const pPr = child(p, 'pPr');
    const styleId = wAttr(child(pPr, 'pStyle'), 'val') || this.defaultPara;
    const jc = wAttr(child(pPr, 'jc'), 'val') || styleProp(this.styles, styleId, 'jc');
    const textAlign = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc] || null;
    const level = headingLevel(this.styles, styleId);
    const styleMarks = level ? {} : resolvedStyleMarks(this.styles, styleId, this.styleCache);
    const content = this.inlines(p, styleMarks, !!level);
    const node = level
      ? { type: 'heading', attrs: { level, textAlign }, content }
      : { type: 'paragraph', attrs: { textAlign }, content };
    if (!content.length) delete node.content;
    let numPr = readNumPr(child(pPr, 'numPr'));
    if (!numPr && styleId) {
      // нумерация задана стилем («Маркированный список 2» и т.п.) — уровень берём из имени стиля
      const sNum = styleProp(this.styles, styleId, 'numPr');
      const sm = /^list (?:bullet|number) (\d)$/.exec((this.styles[styleId] || {}).name || '');
      if (sNum) numPr = { ...sNum, ilvl: sNum.ilvl || (sm ? +sm[1] - 1 : 0) };
    }
    return { node, numPr: level ? null : numPr };
  }

  table(tbl) {
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
        if (vmVal === 'continue' && origins[col]) {
          origins[col].attrs.rowspan += 1;
        } else {
          let content = this.blocks(tc);
          if (!content.length) content = [{ type: 'paragraph' }];
          const cell = { type: 'tableCell', attrs: { colspan: span, rowspan: 1, colwidth: null }, content };
          cells.push(cell);
          if (vmVal === 'restart') origins[col] = cell;
          else delete origins[col];
        }
        col += span;
      }
      if (cells.length) rows.push({ type: 'tableRow', content: cells });
    }
    return rows.length ? { type: 'table', content: rows } : null;
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
        if (children(el, 'headerReference').length || children(el, 'footerReference').length) this.warn('колонтитулы');
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
  const imp = new Importer({
    styles,
    defaultPara,
    theme,
    docDefaults,
    numbering: parseNumbering(await read('word/numbering.xml')),
    rels: parseRels(await read('word/_rels/document.xml.rels')),
  });
  const body = child(documentXml.documentElement, 'body');
  let content = imp.blocks(body);
  if (!content.length) content = [{ type: 'paragraph' }];
  if (zip.file('word/comments.xml')) imp.warn('комментарии');
  return { doc: { type: 'doc', attrs: docDefaults, content }, warnings: imp.warnings };
}
