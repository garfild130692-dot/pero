// Экспорт JSON-документа редактора → .docx (OOXML), собственная генерация.
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { DEFAULT_FONT, DEFAULT_SIZE, highlightName, normalizeColor, parseSize } from '../editor/font.js';
import { NEW_DOC_PARA, ptToTwips } from '../editor/paragraph.js';
import { CELL_SIDES, GRID_BORDERS, TABLE_SIDES, TWIPS_PER_PX, borderToOoxml } from '../editor/table.js';
import { EMU_PER_PX, dataUrlToBytes, extOf } from '../editor/image.js';
import { setHfRefs, writeHfPart } from '../editor/page.js';
import { commentsXml } from '../editor/comments.js';

const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

// Определение пользовательского стиля абзаца (из doc.attrs.styles); basedOn — «Обычный» файла
function customStyleXml(st, normalId) {
  const r = st.run || {};
  const p = st.para || {};
  let rPr = '';
  if (r.fontFamily) rPr += `<w:rFonts w:ascii="${esc(r.fontFamily)}" w:hAnsi="${esc(r.fontFamily)}" w:eastAsia="${esc(r.fontFamily)}" w:cs="${esc(r.fontFamily)}"/>`;
  if (r.bold) rPr += '<w:b/><w:bCs/>';
  if (r.italic) rPr += '<w:i/><w:iCs/>';
  if (r.color) rPr += `<w:color w:val="${r.color.slice(1).toUpperCase()}"/>`;
  if (r.fontSize) rPr += `<w:sz w:val="${Math.round(r.fontSize * 2)}"/><w:szCs w:val="${Math.round(r.fontSize * 2)}"/>`;
  if (r.underline) rPr += '<w:u w:val="single"/>';
  const jc = { center: 'center', right: 'right', justify: 'both' }[st.align];
  const pPr = spacingXml(p) + indXml(p) + (jc ? `<w:jc w:val="${jc}"/>` : '');
  return `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${esc(st.id)}"><w:name w:val="${esc(st.name || st.id)}"/>` +
    `${normalId ? `<w:basedOn w:val="${esc(normalId)}"/>` : ''}<w:qFormat/>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:style>`;
}

// Разметка нового встроенного рисунка (как пишет Word для «Вставка → Рисунок»)
function drawingXml({ rid, cx, cy, id, name, alt }) {
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  return `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${esc(name)}" descr="${esc(alt || '')}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="0" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

// Исходный рисунок с новым размером и замещающим текстом (остальное — как было)
function patchDrawing(xml, cx, cy, alt) {
  let out = xml.replace(/(<[\w]+:extent\b[^>]*?)cx="\d+"([^>]*?)cy="\d+"/, `$1cx="${cx}"$2cy="${cy}"`);
  out = out.replace(/(<[\w]+:xfrm\b[^>]*>\s*(?:<[\w]+:off\b[^>]*\/>\s*)?<[\w]+:ext\b[^>]*?)cx="\d+"([^>]*?)cy="\d+"/, `$1cx="${cx}"$2cy="${cy}"`);
  out = out.replace(/(<[\w]+:docPr\b)([^>]*?)(\/?>)/, (all, tag, attrs, end) => {
    const rest = attrs.replace(/\sdescr="[^"]*"/, '');
    return `${tag}${rest}${alt ? ` descr="${esc(alt)}"` : ''}${end}`;
  });
  return out;
}

// <w:spacing> и <w:ind> из параметров абзаца (в pt); пустая строка, если задавать нечего
function spacingXml(a) {
  let attrs = '';
  if (a.spaceBefore != null) attrs += ` w:before="${ptToTwips(a.spaceBefore)}"`;
  if (a.spaceAfter != null) attrs += ` w:after="${ptToTwips(a.spaceAfter)}"`;
  if (a.lineSpacing != null) {
    const rule = a.lineRule || 'auto';
    const line = rule === 'auto' ? Math.round(a.lineSpacing * 240) : ptToTwips(a.lineSpacing);
    attrs += ` w:line="${line}" w:lineRule="${rule}"`;
  }
  return attrs ? `<w:spacing${attrs}/>` : '';
}
function indXml(a) {
  let attrs = '';
  if (a.indentLeft != null) attrs += ` w:left="${ptToTwips(a.indentLeft)}"`;
  if (a.indentRight != null) attrs += ` w:right="${ptToTwips(a.indentRight)}"`;
  if (a.indentFirst != null) {
    attrs += a.indentFirst < 0 ? ` w:hanging="${ptToTwips(-a.indentFirst)}"` : ` w:firstLine="${ptToTwips(a.indentFirst)}"`;
  }
  return attrs ? `<w:ind${attrs}/>` : '';
}

// Параметры страницы A4 как у русского Word по умолчанию (в твипах, 1 см = 567)
export const PAGE = { w: 11906, h: 16838, top: 1134, right: 850, bottom: 1134, left: 1701 };
const TEXT_WIDTH = PAGE.w - PAGE.left - PAGE.right;

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const esc = (s) =>
  String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c])
    // удаляем символы, недопустимые в XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

// Идентификаторы стилей, на которые ссылается экспорт. В исходном файле они могут быть другими
// (русский Word: «1» вместо Heading1) — тогда берутся из его styles.xml по внутреннему имени стиля.
const OUR_STYLE_IDS = {
  Heading1: 'Heading1', Heading2: 'Heading2', Heading3: 'Heading3', Heading4: 'Heading4', Heading5: 'Heading5', Heading6: 'Heading6',
  ListParagraph: 'ListParagraph', Hyperlink: 'Hyperlink', TableGrid: 'TableGrid',
};
const STYLE_NAMES = {
  Heading1: 'heading 1', Heading2: 'heading 2', Heading3: 'heading 3', Heading4: 'heading 4', Heading5: 'heading 5', Heading6: 'heading 6',
  ListParagraph: 'list paragraph', Hyperlink: 'hyperlink', TableGrid: 'table grid',
};

const DEFAULT_SECTPR = `<w:sectPr><w:pgSz w:w="${PAGE.w}" w:h="${PAGE.h}"/>` +
  `<w:pgMar w:top="${PAGE.top}" w:right="${PAGE.right}" w:bottom="${PAGE.bottom}" w:left="${PAGE.left}" w:header="708" w:footer="708" w:gutter="0"/>` +
  '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

class Exporter {
  constructor({ styleIds = OUR_STYLE_IDS, numBase = 1, relTargets = null, imgBase = 0, sect = (x) => x, customStyles = [] } = {}) {
    this.sect = sect;             // преобразование XML sectPr (подстановка ссылок на свои колонтитулы)
    this.customStyles = new Set(customStyles); // id пользовательских стилей документа
    this.imgBase = imgBase;       // номера картинок Пера, уже занятые в исходном пакете
    this.relTargets = relTargets; // rId → путь части в исходном пакете (если сохраняем в него)
    this.images = [];             // новые картинки: { rid, path, bytes, ext }
    this.links = [];      // гиперссылки → relationships
    this.orderedNums = []; // экземпляры нумерованных списков: { start, ilvl }
    this.styleIds = styleIds;
    this.numBase = numBase; // первый свободный numId (в исходном numbering.xml уже могут быть свои)
  }

  linkId(href) {
    let i = this.links.indexOf(href);
    if (i < 0) { this.links.push(href); i = this.links.length - 1; }
    return 'rIdPeroL' + (i + 1);
  }

  run(node) {
    if (node.type === 'opaque') return node.attrs.xml || ''; // сохраняемый фрагмент — как был
    if (node.type === 'image') return this.image(node);
    if (node.type === 'hardBreak') return '<w:r><w:br/></w:r>';
    if (node.type !== 'text') return '';
    const marks = node.marks || [];
    const has = (t) => marks.some((m) => m.type === t);
    const attrs = (t) => (marks.find((m) => m.type === t) || {}).attrs || {};
    const link = marks.find((m) => m.type === 'link');
    const ts = attrs('textStyle');
    const font = ts.fontFamily;
    const size = parseSize(ts.fontSize);
    const color = normalizeColor(ts.color);
    const bg = has('highlight') ? normalizeColor(attrs('highlight').color) || '#ffff00' : null;
    const hl = bg && highlightName(bg);
    // порядок элементов rPr задан схемой (CT_RPr): rStyle, rFonts, b, bCs, i, iCs, strike, color, sz, szCs,
    // highlight, u, shd, vertAlign
    let rPr = '';
    if (link) rPr += `<w:rStyle w:val="${this.styleIds.Hyperlink}"/>`;
    if (font) rPr += `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}" w:cs="${esc(font)}"/>`;
    if (has('bold')) rPr += '<w:b/><w:bCs/>';
    if (has('italic')) rPr += '<w:i/><w:iCs/>';
    if (has('strike')) rPr += '<w:strike/>';
    if (color) rPr += `<w:color w:val="${color.slice(1).toUpperCase()}"/>`;
    if (size) rPr += `<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>`;
    // стандартный цвет выделения — w:highlight, любой другой — заливка w:shd
    if (hl) rPr += `<w:highlight w:val="${hl}"/>`;
    if (has('underline')) rPr += '<w:u w:val="single"/>';
    if (bg && !hl) rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg.slice(1).toUpperCase()}"/>`;
    if (has('superscript')) rPr += '<w:vertAlign w:val="superscript"/>';
    else if (has('subscript')) rPr += '<w:vertAlign w:val="subscript"/>';
    const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
    const parts = node.text.split('\t');
    let inner = '';
    parts.forEach((p, i) => {
      if (i > 0) inner += '<w:tab/>';
      if (p) inner += `<w:t xml:space="preserve">${esc(p)}</w:t>`;
    });
    const r = `<w:r>${rPrXml}${inner}</w:r>`;
    if (link) {
      const href = link.attrs.href || '';
      if (href.startsWith('#')) return `<w:hyperlink w:anchor="${esc(href.slice(1))}">${r}</w:hyperlink>`;
      return `<w:hyperlink r:id="${this.linkId(href)}">${r}</w:hyperlink>`;
    }
    return r;
  }

  image(node) {
    const a = node.attrs;
    const cx = Math.max(1, Math.round((a.width || 96) * EMU_PER_PX));
    const cy = Math.max(1, Math.round((a.height || 96) * EMU_PER_PX));
    // рисунок из этого же пакета: исходная разметка (обтекание, обрезка, эффекты) с новым размером
    const rid = a.xml && (/r:embed="([^"]+)"/.exec(a.xml) || [])[1];
    if (rid && this.relTargets && this.relTargets[rid] === a.target) return patchDrawing(a.xml, cx, cy, a.alt);
    // новый рисунок (или пришёл из другого документа) — своя часть в word/media
    const data = dataUrlToBytes(a.src);
    if (!data) return '';
    const n = this.imgBase + this.images.length + 1;
    const ext = extOf(data.mime);
    const img = { rid: `rIdPeroImg${n}`, path: `media/pero_image${n}.${ext}`, bytes: data.bytes, ext, mime: data.mime === 'image/jpeg' ? 'image/jpeg' : `image/${ext}` };
    this.images.push(img);
    return drawingXml({ rid: img.rid, cx, cy, id: 90000 + n, name: `Рисунок ${n}`, alt: a.alt });
  }

  inlines(content = []) {
    // соседние текстовые узлы с одной ссылкой склеиваем внутри одной w:hyperlink не обязательно —
    // Word корректно читает несколько подряд идущих гиперссылок.
    return content.map((n) => this.run(n)).join('');
  }

  paragraph(node, { numPr, style } = {}) {
    const attrs = node.attrs || {};
    let pPr = '';
    const pStyle = style || (node.type === 'heading' ? this.styleIds['Heading' + Math.min(Math.max(attrs.level || 1, 1), 6)]
      : (attrs.styleId && this.customStyles.has(attrs.styleId) ? attrs.styleId : null));
    if (pStyle) pPr += `<w:pStyle w:val="${pStyle}"/>`;
    if (numPr) pPr += `<w:numPr><w:ilvl w:val="${numPr.ilvl}"/><w:numId w:val="${numPr.numId}"/></w:numPr>`;
    // порядок в pPr по схеме: pStyle, numPr, spacing, ind, jc
    pPr += spacingXml(attrs) + indXml(attrs);
    const jc = { center: 'center', right: 'right', justify: 'both' }[attrs.textAlign];
    if (jc) pPr += `<w:jc w:val="${jc}"/>`;
    if (attrs.sectPr) pPr += this.sect(attrs.sectPr); // разрыв раздела — в конце pPr по схеме
    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${this.inlines(node.content)}</w:p>`;
  }

  list(node, ilvl) {
    let numId;
    if (node.type === 'bulletList') numId = this.numBase;
    else {
      this.orderedNums.push({ start: (node.attrs && node.attrs.start) || 1, ilvl });
      numId = this.numBase + this.orderedNums.length; // numBase — маркированный, дальше — нумерованные
    }
    let xml = '';
    for (const item of node.content || []) {
      let first = true;
      for (const block of item.content || []) {
        if (block.type === 'bulletList' || block.type === 'orderedList') {
          xml += this.list(block, Math.min(ilvl + 1, 8));
        } else if (block.type === 'paragraph' || block.type === 'heading') {
          xml += first
            ? this.paragraph(block, { numPr: { numId, ilvl }, style: this.styleIds.ListParagraph })
            : this.paragraph(block, { style: this.styleIds.ListParagraph });
          first = false;
        } else {
          xml += this.block(block);
        }
      }
    }
    return xml;
  }

  table(node) {
    const rows = node.content || [];
    // 1) раскладка ячеек по сетке с учётом colspan/rowspan
    const grid = []; // grid[r][c] = { cell, origin:true } | { cont:true, colspan } | { covered:true }
    let cols = 0;
    rows.forEach((row, r) => {
      grid[r] = grid[r] || [];
      let c = 0;
      for (const cell of row.content || []) {
        while (grid[r][c]) c++;
        const cs = (cell.attrs && cell.attrs.colspan) || 1;
        const rs = Math.min((cell.attrs && cell.attrs.rowspan) || 1, rows.length - r);
        for (let rr = r; rr < r + rs; rr++) {
          grid[rr] = grid[rr] || [];
          for (let cc = c; cc < c + cs; cc++) {
            if (rr === r && cc === c) grid[rr][cc] = { cell, origin: true, colspan: cs, vMerge: rs > 1 };
            else if (cc === c) grid[rr][cc] = { cell, cont: true, colspan: cs };
            else grid[rr][cc] = { covered: true };
          }
        }
        c += cs;
      }
      cols = Math.max(cols, grid[r].length);
    });
    cols = Math.max(cols, 1);

    // 2) ширины колонок: из colwidth (px) или поровну
    const widths = new Array(cols).fill(null);
    rows.forEach((row, r) => {
      grid[r].forEach((g, c) => {
        if (g && g.origin && g.cell.attrs && Array.isArray(g.cell.attrs.colwidth)) {
          g.cell.attrs.colwidth.forEach((px, k) => { if (px && c + k < cols) widths[c + k] = Math.round(px * TWIPS_PER_PX); });
        }
      });
    });
    const known = widths.filter(Boolean).reduce((a, b) => a + b, 0);
    const unknown = widths.filter((w) => !w).length;
    const rest = unknown ? Math.max(Math.floor((TEXT_WIDTH - known) / unknown), 600) : 0;
    const W = widths.map((w) => w || rest);

    // tblPr по схеме: tblStyle, tblW, jc, tblBorders, tblLook. Стиль TableGrid оставляем ради интервалов абзацев
    // в ячейках (после 0, одинарный); границы пишем явно — они перекрывают стиль.
    const ta = node.attrs || {};
    const tb = { ...GRID_BORDERS, ...(ta.borders || {}) };
    let xml = `<w:tbl><w:tblPr><w:tblStyle w:val="${this.styleIds.TableGrid}"/><w:tblW w:w="0" w:type="auto"/>` +
      (ta.align === 'center' || ta.align === 'right' ? `<w:jc w:val="${ta.align}"/>` : '') +
      `<w:tblBorders>${TABLE_SIDES.map((s) => borderToOoxml(s, tb[s])).join('')}</w:tblBorders>` +
      '<w:tblLook w:val="04A0"/></w:tblPr>';
    xml += `<w:tblGrid>${W.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

    grid.forEach((cells, r) => {
      // trPr по схеме: trHeight, tblHeader
      const ra = (rows[r] && rows[r].attrs) || {};
      let trPr = '';
      if (ra.height) trPr += `<w:trHeight w:val="${ptToTwips(ra.height)}"${ra.heightRule === 'exact' ? ' w:hRule="exact"' : ''}/>`;
      if (ra.repeatHeader) trPr += '<w:tblHeader/>';
      xml += `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}`;
      for (let c = 0; c < cols; c++) {
        const g = cells[c];
        if (!g || g.covered) {
          if (!g) xml += `<w:tc><w:tcPr><w:tcW w:w="${W[c]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
          continue;
        }
        const width = W.slice(c, c + g.colspan).reduce((a, b) => a + b, 0);
        let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
        if (g.colspan > 1) tcPr += `<w:gridSpan w:val="${g.colspan}"/>`;
        if (g.cont) tcPr += '<w:vMerge/>';
        else if (g.vMerge) tcPr += '<w:vMerge w:val="restart"/>';
        // оформление: у продолжения объединённой ячейки — как у начала (Word рисует каждую часть по её tcPr)
        // порядок: tcBorders, shd, vAlign
        const ca = g.cell.attrs || {};
        const cb = ca.borders || {};
        const sides = CELL_SIDES.filter((s) => cb[s]);
        if (sides.length) tcPr += `<w:tcBorders>${sides.map((s) => borderToOoxml(s, cb[s])).join('')}</w:tcBorders>`;
        if (ca.background) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${ca.background.slice(1).toUpperCase()}"/>`;
        if (ca.verticalAlign === 'center' || ca.verticalAlign === 'bottom') tcPr += `<w:vAlign w:val="${ca.verticalAlign}"/>`;
        let body = g.cont ? '<w:p/>' : this.blocks(g.cell.content);
        if (!body.endsWith('</w:p>')) body += '<w:p/>'; // ячейка обязана заканчиваться абзацем
        xml += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${body}</w:tc>`;
      }
      xml += '</w:tr>';
    });
    return xml + '</w:tbl>';
  }

  block(node) {
    switch (node.type) {
      case 'paragraph':
      case 'heading':
        return this.paragraph(node);
      case 'bulletList':
      case 'orderedList':
        return this.list(node, 0);
      case 'table':
        return this.table(node);
      default:
        // неизвестный блок — сохраняем хотя бы текст
        return node.content ? this.blocks(node.content) : '';
    }
  }

  blocks(content = []) {
    return content.map((n) => this.block(n)).join('');
  }
}

// ---------- Служебные части пакета ----------
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

function coreXml() {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Pero</Application></Properties>`;

function headingStyle(n) {
  const size = { 1: 32, 2: 26, 3: 24, 4: 22, 5: 22, 6: 22 }[n];
  const color = n <= 2 ? '2F5496' : '1F3763';
  return `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 240 : 40}" w:after="0" w:line="259" w:lineRule="auto"/><w:outlineLvl w:val="${n - 1}"/></w:pPr>` +
    // шрифт заголовков задаём явно: темы оформления (theme1.xml) мы не пишем
    `<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:eastAsia="Calibri Light" w:cs="Calibri Light"/>${n >= 4 ? '<w:i/>' : ''}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
}

// Шрифт и размер документа по умолчанию — из атрибутов документа
const stylesXml = (font, size, para) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}" w:cs="${esc(font)}"/><w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>${spacingXml(para)}${indXml(para)}</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
${[1, 2, 3, 4, 5, 6].map(headingStyle).join('\n')}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;

// Наши определения нумерации. absBase/numBase — первые свободные идентификаторы (при слиянии с исходным файлом).
function numberingParts(orderedNums, absBase = 0, numBase = 1) {
  // маркеры Word: точка (Symbol U+F0B7), «o» (Courier New), квадрат (Wingdings U+F0A7)
  const BULLETS = [['\uF0B7', 'Symbol'], ['o', 'Courier New'], ['\uF0A7', 'Wingdings']];
  const bulletLvls = Array.from({ length: 9 }, (_, i) => {
    const [text, font] = BULLETS[i % 3];
    return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr></w:lvl>`;
  }).join('');
  const FMT = ['decimal', 'lowerLetter', 'lowerRoman'];
  const decimalLvls = Array.from({ length: 9 }, (_, i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${FMT[i % 3]}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="${i % 3 === 2 ? 'right' : 'left'}"/>` +
    `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="${i % 3 === 2 ? 180 : 360}"/></w:pPr></w:lvl>`
  ).join('');
  let nums = `<w:num w:numId="${numBase}"><w:abstractNumId w:val="${absBase}"/></w:num>`;
  orderedNums.forEach((n, k) => {
    nums += `<w:num w:numId="${numBase + k + 1}"><w:abstractNumId w:val="${absBase + 1}"/>` +
      `<w:lvlOverride w:ilvl="${n.ilvl}"><w:startOverride w:val="${n.start}"/></w:lvlOverride></w:num>`;
  });
  const abstracts =
    `<w:abstractNum w:abstractNumId="${absBase}"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="${absBase + 1}"><w:multiLevelType w:val="hybridMultilevel"/>${decimalLvls}</w:abstractNum>`;
  return { abstracts, nums };
}

function numberingXml(orderedNums) {
  const { abstracts, nums } = numberingParts(orderedNums);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${NS}>
${abstracts}
${nums}
</w:numbering>`;
}

// ---------- Сохранность исходного файла ----------
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relType = (t) => `${REL}/${t}`;

// Путь части относительно папки, где лежит документ (Target в .rels — относительный)
function resolvePart(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dir ? dir.split('/') : []).concat(target.split('/'));
  const out = [];
  for (const x of parts) { if (x === '..') out.pop(); else if (x && x !== '.') out.push(x); }
  return out.join('/');
}

// Прочитать из исходного пакета то, что нужно для слияния
async function readSource(zip) {
  const text = async (path) => { const f = zip.file(path); return f ? f.async('string') : null; };
  const rootRels = (await text('_rels/.rels')) || '';
  const m = /<Relationship\b[^>]*Type="[^"]*\/officeDocument"[^>]*>/.exec(rootRels);
  const target = m && /Target="([^"]+)"/.exec(m[0]);
  const mainPath = target ? resolvePart('', target[1]) : 'word/document.xml';
  const dir = mainPath.includes('/') ? mainPath.slice(0, mainPath.lastIndexOf('/')) : '';
  const file = mainPath.slice(dir.length ? dir.length + 1 : 0);
  const relsPath = `${dir ? dir + '/' : ''}_rels/${file}.rels`;

  const docXml = (await text(mainPath)) || '';
  const rootTag = (/<w:document\b[^>]*>/.exec(docXml) || [`<w:document ${NS}>`])[0];
  const bodyAt = docXml.indexOf('<w:body');
  const preBody = bodyAt > 0 ? docXml.slice(docXml.indexOf(rootTag) + rootTag.length, bodyAt) : ''; // напр. w:background
  const rootNs = [...rootTag.matchAll(/xmlns:([\w.-]+)="([^"]*)"/g)].map((x) => [x[1], x[2]]);

  const rels = (await text(relsPath)) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const partOf = (type) => {
    const r = new RegExp(`<Relationship\\b[^>]*Type="${relType(type).replace(/[.\/]/g, '\\$&')}"[^>]*>`).exec(rels);
    const t = r && /Target="([^"]+)"/.exec(r[0]);
    return t ? resolvePart(dir, t[1]) : null;
  };
  // все связи документа: rId → путь части (для проверки, что рисунок — из этого пакета)
  const relTargets = {};
  for (const m2 of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m2[0]);
    const t = /Target="([^"]+)"/.exec(m2[0]);
    if (id && t && !/TargetMode="External"/.test(m2[0])) relTargets[id[1]] = resolvePart(dir, t[1]);
  }
  const stylesPath = partOf('styles');
  const numberingPath = partOf('numbering');
  return {
    mainPath, dir, relsPath, rootTag, preBody, rootNs, rels, relTargets, stylesPath, numberingPath,
    styles: stylesPath ? await text(stylesPath) : null,
    numbering: numberingPath ? await text(numberingPath) : null,
    contentTypes: (await text('[Content_Types].xml')) || CONTENT_TYPES,
    core: await text('docProps/core.xml'),
  };
}

// Сопоставить нужные экспорту стили со стилями исходного файла; недостающие — дописать
function mergeStyles(stylesXml, fallback) {
  const doc = new DOMParser().parseFromString(stylesXml, 'text/xml');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const byName = {};
  const ids = new Set();
  const defaults = {};
  for (const st of Array.from(doc.getElementsByTagNameNS(W, 'style'))) {
    const id = st.getAttributeNS(W, 'styleId');
    const type = st.getAttributeNS(W, 'type');
    ids.add(id);
    const nameEl = st.getElementsByTagNameNS(W, 'name')[0];
    const name = nameEl ? (nameEl.getAttributeNS(W, 'val') || '').toLowerCase() : '';
    if (name && !byName[name]) byName[name] = id;
    if (['1', 'true', 'on'].includes(String(st.getAttributeNS(W, 'default')).toLowerCase())) defaults[type] = id;
  }
  const styleIds = {};
  let add = '';
  for (const [key, name] of Object.entries(STYLE_NAMES)) {
    if (byName[name]) { styleIds[key] = byName[name]; continue; }
    let id = OUR_STYLE_IDS[key];
    while (ids.has(id)) id += 'Pero';
    ids.add(id);
    styleIds[key] = id;
    // определение из наших стилей; basedOn — на стили по умолчанию исходного файла (или без него)
    let xml = fallback[key].replace(`w:styleId="${OUR_STYLE_IDS[key]}"`, `w:styleId="${id}"`);
    const base = { Normal: defaults.paragraph, DefaultParagraphFont: defaults.character, TableNormal: defaults.table };
    xml = xml.replace(/<w:basedOn w:val="(\w+)"\/>/, (all, b) => (base[b] ? `<w:basedOn w:val="${base[b]}"/>` : ''));
    add += xml;
  }
  return { styleIds, defaults, xml: add ? stylesXml.replace(/<\/w:styles>\s*$/, `${add}</w:styles>`) : stylesXml };
}

// Наши определения стилей (из stylesXml), по ключу
function ourStyleDefs() {
  const all = stylesXml(DEFAULT_FONT, DEFAULT_SIZE, NEW_DOC_PARA);
  const out = {};
  for (const key of Object.keys(OUR_STYLE_IDS)) {
    const m = new RegExp(`<w:style [^>]*w:styleId="${OUR_STYLE_IDS[key]}"[\\s\\S]*?</w:style>`).exec(all);
    out[key] = m[0];
  }
  return out;
}

const maxId = (xml, re) => Math.max(-1, ...[...(xml || '').matchAll(re)].map((x) => +x[1]));

// Убрать локальные объявления пространств имён, которые уже есть в корне документа
function stripRootNs(xml, rootNs) {
  let out = xml;
  for (const [pfx, uri] of rootNs) out = out.split(` xmlns:${pfx}="${uri}"`).join('');
  return out;
}

// Типы содержимого для картинок (по расширению)
function addImageTypes(ct, images) {
  let out = ct;
  for (const img of images) {
    if (new RegExp(`<Default\\b[^>]*Extension="${img.ext}"`, 'i').test(out)) continue;
    out = out.replace('</Types>', `<Default Extension="${img.ext}" ContentType="${img.mime}"/></Types>`);
  }
  return out;
}

function addOverride(ct, partName, type) {
  if (ct.includes(`PartName="/${partName}"`)) return ct;
  return ct.replace('</Types>', `<Override PartName="/${partName}" ContentType="${type}"/></Types>`);
}

/**
 * @param {object} doc — JSON документа редактора (editor.getJSON())
 * @param {{ source?: ArrayBuffer|Uint8Array }} [options] — исходный .docx: всё, что Перо не меняет
 *   (колонтитулы, сноски, комментарии, картинки, тема, настройки…), переносится из него без изменений
 * @returns {Promise<Uint8Array>} — содержимое .docx
 */
export async function exportDocx(doc, { source } = {}) {
  const da = doc.attrs || {};
  // Свои колонтитулы (созданы в Пере): части header_pero.xml/footer_pero.xml и ссылки на них во всех разделах.
  // Колонтитулы из файла (own = false) не трогаем — они уже в пакете.
  const hf = da.hf || {};
  const own = hf.own || {};
  const hfIds = {};
  if (own.header) hfIds.header = hf.header && hf.header.length ? 'rIdPeroH' : null;
  if (own.footer) hfIds.footer = hf.footer && hf.footer.length ? 'rIdPeroF' : null;
  const sect = (x) => (own.header || own.footer ? setHfRefs(x, hfIds) : x);
  const sectPr = sect(da.sectPr || DEFAULT_SECTPR);
  const hfParts = [
    hfIds.header && { id: 'rIdPeroH', file: 'header_pero.xml', xml: writeHfPart('hdr', hf.header, NS), type: 'header' },
    hfIds.footer && { id: 'rIdPeroF', file: 'footer_pero.xml', xml: writeHfPart('ftr', hf.footer, NS), type: 'footer' },
  ].filter(Boolean);
  const hfRels = hfParts.map((p) => `<Relationship Id="${p.id}" Type="${relType(p.type)}" Target="${p.file}"/>`).join('');
  const addHfTypes = (ct, prefix) => hfParts.reduce((acc, p) => addOverride(acc, prefix + p.file,
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${p.type}+xml`), ct);
  const customs = da.styles || [];
  if (!source) {
    const ex = new Exporter({ sect, customStyles: customs.map((s) => s.id) });
    let body = ex.blocks(doc.content || []);
    if (!body) body = '<w:p/>';
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${body}${sectPr}</w:body></w:document>`;
    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${relType('styles')}" Target="styles.xml"/>
<Relationship Id="rId2" Type="${relType('numbering')}" Target="numbering.xml"/>
${ex.links.map((href, i) => `<Relationship Id="rIdPeroL${i + 1}" Type="${relType('hyperlink')}" Target="${esc(href)}" TargetMode="External"/>`).join('\n')}
${ex.images.map((img) => `<Relationship Id="${img.rid}" Type="${relType('image')}" Target="${img.path}"/>`).join('\n')}
${hfRels}
${(da.comments || []).length ? `<Relationship Id="rIdPeroC" Type="${relType('comments')}" Target="comments.xml"/>` : ''}
</Relationships>`;
    const zip = new JSZip();
    if ((da.comments || []).length) zip.file('word/comments.xml', commentsXml(da.comments, NS));
    let ct0 = addHfTypes(addImageTypes(CONTENT_TYPES, ex.images), 'word/');
    if ((da.comments || []).length) ct0 = addOverride(ct0, 'word/comments.xml', COMMENTS_TYPE);
    zip.file('[Content_Types].xml', ct0);
    for (const p of hfParts) zip.file('word/' + p.file, p.xml);
    for (const img of ex.images) zip.file('word/' + img.path, img.bytes);
    zip.file('_rels/.rels', ROOT_RELS);
    zip.file('docProps/core.xml', coreXml());
    zip.file('docProps/app.xml', APP_XML);
    zip.file('word/document.xml', documentXml);
    zip.file('word/styles.xml', stylesXml(da.fontFamily || DEFAULT_FONT, parseSize(da.fontSize) || DEFAULT_SIZE, { ...NEW_DOC_PARA, ...(da.para || {}) })
      .replace('</w:styles>', `${customs.map((s) => customStyleXml(s, 'Normal')).join('')}</w:styles>`));
    zip.file('word/numbering.xml', numberingXml(ex.orderedNums));
    zip.file('word/_rels/document.xml.rels', docRels);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  // Исходный пакет: меняем только документ, стили (дописываем недостающие), нумерацию (добавляем свою)
  // и связи (добавляем гиперссылки). Всё остальное остаётся байт в байт.
  const zip = await JSZip.loadAsync(source);
  const src = await readSource(zip);
  let ct = src.contentTypes;
  let rels = src.rels;
  const prefix = src.dir ? src.dir + '/' : '';

  let styleIds = OUR_STYLE_IDS;
  if (src.styles) {
    const merged = mergeStyles(src.styles, ourStyleDefs());
    styleIds = merged.styleIds;
    let xml = merged.xml;
    const normalId = merged.defaults.paragraph || null;
    // пользовательские стили: изменённые в Пере — переписать определение, новые — дописать, прочие — как в файле
    for (const st of customs) {
      const re = new RegExp(`<w:style\\b[^>]*w:styleId="${st.id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>[\\s\\S]*?</w:style>`);
      if (re.test(xml)) { if (st.modified) xml = xml.replace(re, customStyleXml(st, normalId)); }
      else xml = xml.replace(/<\/w:styles>\s*$/, `${customStyleXml(st, normalId)}</w:styles>`);
    }
    // «Обычный» изменён в Пере: умолчания документа целиком в docDefaults, из стиля «Обычный» — убрать шрифт/размер/интервалы
    if (da.defaultsModified) {
      const ours = stylesXml(da.fontFamily || DEFAULT_FONT, parseSize(da.fontSize) || DEFAULT_SIZE, { ...NEW_DOC_PARA, ...(da.para || {}) });
      const dd = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(ours)[0];
      xml = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.test(xml) ? xml.replace(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/, dd) : xml.replace(/(<w:styles\b[^>]*>)/, `$1${dd}`);
      if (normalId) {
        const re = new RegExp(`(<w:style\\b[^>]*w:styleId="${normalId}"[^>]*>)([\\s\\S]*?)(</w:style>)`);
        xml = xml.replace(re, (all, a, body, c) => a + body
          .replace(/<w:rFonts\b[^>]*\/>|<w:sz\b[^>]*\/>|<w:szCs\b[^>]*\/>|<w:spacing\b[^>]*\/>|<w:ind\b[^>]*\/>/g, '')
          .replace(/<w:rPr>\s*<\/w:rPr>|<w:pPr>\s*<\/w:pPr>/g, '') + c);
      }
    }
    zip.file(src.stylesPath, xml);
  } else {
    zip.file(prefix + 'styles.xml', stylesXml(da.fontFamily || DEFAULT_FONT, parseSize(da.fontSize) || DEFAULT_SIZE, { ...NEW_DOC_PARA, ...(da.para || {}) }));
    rels = rels.replace('</Relationships>', `<Relationship Id="rIdPeroS" Type="${relType('styles')}" Target="styles.xml"/></Relationships>`);
    ct = addOverride(ct, prefix + 'styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
  }

  const absBase = maxId(src.numbering, /w:abstractNumId="(\d+)"/g) + 1;
  const numBase = Math.max(1, maxId(src.numbering, /<w:num\b[^>]*w:numId="(\d+)"/g) + 1);
  // картинки, добавленные Пером при прошлых сохранениях, остаются; новые нумеруются дальше
  const imgBase = Math.max(0, ...Object.keys(zip.files).map((n) => +((/\/media\/pero_image(\d+)\.\w+$/.exec(n) || [])[1] || 0)),
    ...[...src.rels.matchAll(/Id="rIdPeroImg(\d+)"/g)].map((m) => +m[1]));
  const ex = new Exporter({ styleIds, numBase, relTargets: src.relTargets, imgBase, sect, customStyles: customs.map((s) => s.id) });
  let body = ex.blocks(doc.content || []);
  if (!body) body = '<w:p/>';
  body = stripRootNs(body + sectPr, src.rootNs);
  zip.file(src.mainPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${src.rootTag}${src.preBody}<w:body>${body}</w:body></w:document>`);

  const { abstracts, nums } = numberingParts(ex.orderedNums, absBase, numBase);
  if (src.numbering) {
    // по схеме все abstractNum идут перед num; num — перед numIdMacAtCleanup
    let x = src.numbering;
    const firstNum = x.search(/<w:num[\s>]/);
    x = firstNum >= 0 ? x.slice(0, firstNum) + abstracts + x.slice(firstNum) : x.replace('</w:numbering>', abstracts + '</w:numbering>');
    const tail = x.search(/<w:numIdMacAtCleanup\b|<\/w:numbering>/);
    x = x.slice(0, tail) + nums + x.slice(tail);
    zip.file(src.numberingPath, x);
  } else {
    zip.file(prefix + 'numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${NS}>${abstracts}${nums}</w:numbering>`);
    rels = rels.replace('</Relationships>', `<Relationship Id="rIdPeroN" Type="${relType('numbering')}" Target="numbering.xml"/></Relationships>`);
    ct = addOverride(ct, prefix + 'numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  }

  // гиперссылки Пера; исходные связи (картинки, колонтитулы, сноски…) не трогаем
  const linkRels = ex.links.map((href, i) => `<Relationship Id="rIdPeroL${i + 1}" Type="${relType('hyperlink')}" Target="${esc(href)}" TargetMode="External"/>`).join('');
  const imgRels = ex.images.map((img) => `<Relationship Id="${img.rid}" Type="${relType('image')}" Target="${img.path}"/>`).join('');
  // ссылки Пера пересоздаются заново; картинки прошлых сохранений остаются (на них ссылаются рисунки)
  rels = rels.replace(/<Relationship\b[^>]*Id="rIdPero(?:L\d+|H|F)"[^>]*\/>/g, '').replace('</Relationships>', linkRels + imgRels + hfRels + '</Relationships>');
  for (const img of ex.images) zip.file(prefix + img.path, img.bytes);
  for (const name of ['header_pero.xml', 'footer_pero.xml']) if (!hfParts.some((p) => p.file === name)) zip.remove(prefix + name);
  for (const p of hfParts) zip.file(prefix + p.file, p.xml);
  ct = addHfTypes(ct, prefix);

  // Примечания менялись в Пере: comments.xml — заново (простой текст); расширенные части примечаний
  // (ответы, отметки «решено», id абзацев) больше не соответствуют — убираем их вместе со связями и типами
  if (da.commentsEdited) {
    const list = da.comments || [];
    const cm = /<Relationship\b[^>]*Type="[^"]*\/comments"[^>]*Target="([^"]+)"[^>]*\/>|<Relationship\b[^>]*Target="([^"]+)"[^>]*Type="[^"]*\/comments"[^>]*\/>/.exec(rels);
    const cPath = cm ? resolvePart(src.dir, cm[1] || cm[2]) : prefix + 'comments.xml';
    for (const kind of ['commentsExtended', 'commentsIds', 'commentsExtensible']) {
      const re = new RegExp(`<Relationship\\b[^>]*Type="[^"]*/${kind}"[^>]*/>`, 'g');
      for (const m of rels.match(re) || []) {
        const t = /Target="([^"]+)"/.exec(m);
        if (t) { const part = resolvePart(src.dir, t[1]); zip.remove(part); ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${part.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}"[^>]*/>`), ''); }
      }
      rels = rels.replace(re, '');
    }
    if (list.length) {
      zip.file(cPath, commentsXml(list, NS));
      if (!cm) rels = rels.replace('</Relationships>', `<Relationship Id="rIdPeroC" Type="${relType('comments')}" Target="comments.xml"/></Relationships>`);
      ct = addOverride(ct, cPath, COMMENTS_TYPE);
    } else if (cm) {
      zip.remove(cPath);
      rels = rels.replace(cm[0], '');
      ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${cPath.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}"[^>]*/>`), '');
    }
  }
  ct = addImageTypes(ct, ex.images);
  zip.file(src.relsPath, rels);
  zip.file('[Content_Types].xml', ct);
  if (src.core) {
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', src.core.replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${now}$2`));
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
