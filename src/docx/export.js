// Экспорт JSON-документа редактора → .docx (OOXML), собственная генерация.
import JSZip from 'jszip';

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

class Exporter {
  constructor() {
    this.links = [];      // гиперссылки → relationships
    this.orderedNums = []; // экземпляры нумерованных списков: { start, ilvl }
  }

  linkId(href) {
    let i = this.links.indexOf(href);
    if (i < 0) { this.links.push(href); i = this.links.length - 1; }
    return 'rIdL' + (i + 1);
  }

  run(node) {
    if (node.type === 'hardBreak') return '<w:r><w:br/></w:r>';
    if (node.type !== 'text') return '';
    const marks = node.marks || [];
    const has = (t) => marks.some((m) => m.type === t);
    const link = marks.find((m) => m.type === 'link');
    // порядок элементов rPr задан схемой: rStyle, b, i, strike, color, u
    let rPr = '';
    if (link) rPr += '<w:rStyle w:val="Hyperlink"/>';
    if (has('bold')) rPr += '<w:b/><w:bCs/>';
    if (has('italic')) rPr += '<w:i/><w:iCs/>';
    if (has('strike')) rPr += '<w:strike/>';
    if (has('underline')) rPr += '<w:u w:val="single"/>';
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

  inlines(content = []) {
    // соседние текстовые узлы с одной ссылкой склеиваем внутри одной w:hyperlink не обязательно —
    // Word корректно читает несколько подряд идущих гиперссылок.
    return content.map((n) => this.run(n)).join('');
  }

  paragraph(node, { numPr, style } = {}) {
    const attrs = node.attrs || {};
    let pPr = '';
    const pStyle = style || (node.type === 'heading' ? 'Heading' + (attrs.level || 1) : null);
    if (pStyle) pPr += `<w:pStyle w:val="${pStyle}"/>`;
    if (numPr) pPr += `<w:numPr><w:ilvl w:val="${numPr.ilvl}"/><w:numId w:val="${numPr.numId}"/></w:numPr>`;
    const jc = { center: 'center', right: 'right', justify: 'both' }[attrs.textAlign];
    if (jc) pPr += `<w:jc w:val="${jc}"/>`;
    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${this.inlines(node.content)}</w:p>`;
  }

  list(node, ilvl) {
    let numId;
    if (node.type === 'bulletList') numId = 1;
    else {
      this.orderedNums.push({ start: (node.attrs && node.attrs.start) || 1, ilvl });
      numId = 1 + this.orderedNums.length; // 1 — маркированный, дальше — нумерованные
    }
    let xml = '';
    for (const item of node.content || []) {
      let first = true;
      for (const block of item.content || []) {
        if (block.type === 'bulletList' || block.type === 'orderedList') {
          xml += this.list(block, Math.min(ilvl + 1, 8));
        } else if (block.type === 'paragraph' || block.type === 'heading') {
          xml += first
            ? this.paragraph(block, { numPr: { numId, ilvl }, style: 'ListParagraph' })
            : this.paragraph(block, { style: 'ListParagraph' });
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
            else if (cc === c) grid[rr][cc] = { cont: true, colspan: cs };
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
          g.cell.attrs.colwidth.forEach((px, k) => { if (px && c + k < cols) widths[c + k] = Math.round(px * 15); });
        }
      });
    });
    const known = widths.filter(Boolean).reduce((a, b) => a + b, 0);
    const unknown = widths.filter((w) => !w).length;
    const rest = unknown ? Math.max(Math.floor((TEXT_WIDTH - known) / unknown), 600) : 0;
    const W = widths.map((w) => w || rest);

    const border = (s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    let xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
      `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
      '<w:tblLook w:val="04A0"/></w:tblPr>';
    xml += `<w:tblGrid>${W.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

    grid.forEach((cells) => {
      xml += '<w:tr>';
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
    `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 240 : 40}" w:after="0"/><w:outlineLvl w:val="${n - 1}"/></w:pPr>` +
    `<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>${n >= 4 ? '<w:i/>' : ''}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
${[1, 2, 3, 4, 5, 6].map(headingStyle).join('\n')}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;

function numberingXml(orderedNums) {
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
  let nums = '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  orderedNums.forEach((n, k) => {
    nums += `<w:num w:numId="${k + 2}"><w:abstractNumId w:val="1"/>` +
      `<w:lvlOverride w:ilvl="${n.ilvl}"><w:startOverride w:val="${n.start}"/></w:lvlOverride></w:num>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${NS}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${decimalLvls}</w:abstractNum>
${nums}
</w:numbering>`;
}

/**
 * @param {object} doc — JSON документа редактора (editor.getJSON())
 * @returns {Promise<Uint8Array>} — содержимое .docx
 */
export async function exportDocx(doc) {
  const ex = new Exporter();
  let body = ex.blocks(doc.content || []);
  if (!body) body = '<w:p/>';
  const sectPr = `<w:sectPr><w:pgSz w:w="${PAGE.w}" w:h="${PAGE.h}"/>` +
    `<w:pgMar w:top="${PAGE.top}" w:right="${PAGE.right}" w:bottom="${PAGE.bottom}" w:left="${PAGE.left}" w:header="708" w:footer="708" w:gutter="0"/>` +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${body}${sectPr}</w:body></w:document>`;

  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
${ex.links.map((href, i) => `<Relationship Id="rIdL${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(href)}" TargetMode="External"/>`).join('\n')}
</Relationships>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('docProps/core.xml', coreXml());
  zip.file('docProps/app.xml', APP_XML);
  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', numberingXml(ex.orderedNums));
  zip.file('word/_rels/document.xml.rels', docRels);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
