// Документы Word 97–2003 (.doc, [MS-DOC]) — своим кодом (этап 19.4). Из двоичного документа строится пакет
// WordprocessingML (document.xml, styles.xml, numbering.xml, колонтитулы, сноски, примечания, рисунки) — дальше его
// читает обычный импорт .docx, так что стили, таблицы, списки и всё прочее разбираются одним кодом.
//
// Устройство .doc: составной файл [MS-CFB]; поток WordDocument начинается с FIB (где что лежит), текст — куски (piece
// table, Clx в потоке 0Table/1Table), свойства символов и абзацев — страницы FKP (CHPX/PAPX) с наборами sprm, стили —
// STSH, разделы — PlcfSed + SEPX, списки — PlfLst/PlfLfo, колонтитулы — PlcfHdd, сноски — PlcffndRef/Txt, рисунки — поток
// Data (PICF + OfficeArt). Разметка: 0x0D — конец абзаца, 0x07 — конец ячейки/строки таблицы, 0x13/0x14/0x15 — поле.
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { readCfb } from '../formats/cfb.js';
import { watermarkRunXml } from '../editor/design.js';

// ---------- чтение байтов ----------
const u8 = (d, o) => d[o];
const u16 = (d, o) => d[o] | (d[o + 1] << 8);
const i16 = (d, o) => { const v = u16(d, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (d, o) => (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0;
const i32 = (d, o) => u32(d, o) | 0;
const utf16 = (d, o, n) => { let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(u16(d, o + i * 2)); return s; };
const hex6 = (v) => v.toString(16).padStart(6, '0').toUpperCase();
// экранирование для XML; символы, недопустимые в XML 1.0 (управляющие, U+FFFE/FFFF, одиночные суррогаты), — выбрасываются
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
// ссылки HYPERLINK из файла: только веб, почта и закладки (file:, UNC, javascript: в Word сработали бы при щелчке)
const safeUrl = (u) => /^(https?:\/\/|mailto:)/i.test(String(u).trim());
// COLORREF (r, g, b, fAuto) → 'RRGGBB' | 'auto'
const colorref = (d, o) => (d[o + 3] === 0xFF ? 'auto' : hex6((d[o] << 16) | (d[o + 1] << 8) | d[o + 2]));

export const isDoc = (bytes) => { try { const c = readCfb(bytes); const w = c.stream('WordDocument'); return !!w && u16(w, 0) === 0xA5EC; } catch { return false; } };

// Цвета ico (старые 16 цветов Word) и имена выделения OOXML
const ICO = [null, '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
const HIGHLIGHT = [null, 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];
const KUL = { 1: 'single', 2: 'words', 3: 'double', 4: 'dotted', 6: 'thick', 7: 'dash', 9: 'dotDash', 10: 'dotDotDash', 11: 'wave', 20: 'dottedHeavy', 23: 'dashedHeavy', 25: 'dashDotHeavy', 26: 'dashDotDotHeavy', 27: 'wavyHeavy', 39: 'dashLong', 43: 'wavyDouble', 55: 'dashLongHeavy' };
const BRC = { 1: 'single', 2: 'thick', 3: 'double', 5: 'hairline', 6: 'dotted', 7: 'dashed', 8: 'dotDash', 9: 'dotDotDash', 10: 'triple', 11: 'thinThickSmallGap', 12: 'thickThinSmallGap', 13: 'thinThickThinSmallGap', 14: 'thinThickMediumGap', 15: 'thickThinMediumGap', 16: 'thinThickThinMediumGap', 17: 'thinThickLargeGap', 18: 'thickThinLargeGap', 19: 'thinThickThinLargeGap', 20: 'wave', 21: 'doubleWave', 22: 'dashSmallGap', 23: 'dashDotStroked', 24: 'threeDEmboss', 25: 'threeDEngrave', 26: 'outset', 27: 'inset' };
const NFC = { 0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 6: 'cardinalText', 7: 'ordinalText', 22: 'decimalZero', 23: 'bullet', 58: 'russianLower', 59: 'russianUpper', 255: 'none' };
const SHD_PAT = { 0: 'clear', 1: 'solid', 2: 'pct5', 3: 'pct10', 4: 'pct20', 5: 'pct25', 6: 'pct30', 7: 'pct40', 8: 'pct50', 9: 'pct60', 10: 'pct70', 11: 'pct75', 12: 'pct80', 13: 'pct90' };
// встроенные стили (sti) — внутренние имена OOXML; по ним импорт узнаёт стили (в файле имя может быть локализованным).
// Сверено с Word: tests/out/stimap.mjs — порядок стилей .doc и .doc.docx (Word сам пересохранил) совпадает, 0 противоречий
const STI = {
  0: 'Normal', 1: 'heading 1', 2: 'heading 2', 3: 'heading 3', 4: 'heading 4', 5: 'heading 5', 6: 'heading 6',
  7: 'heading 7', 8: 'heading 8', 9: 'heading 9', 10: 'index 1', 11: 'index 2', 12: 'index 3', 13: 'index 4',
  14: 'index 5', 15: 'index 6', 16: 'index 7', 17: 'index 8', 18: 'index 9', 19: 'toc 1', 20: 'toc 2', 21: 'toc 3',
  22: 'toc 4', 23: 'toc 5', 24: 'toc 6', 25: 'toc 7', 26: 'toc 8', 27: 'toc 9', 28: 'Normal Indent',
  29: 'footnote text', 30: 'annotation text', 31: 'header', 32: 'footer', 33: 'index heading', 34: 'caption',
  35: 'table of figures', 36: 'envelope address', 37: 'envelope return', 38: 'footnote reference',
  39: 'annotation reference', 40: 'line number', 41: 'page number', 42: 'endnote reference', 43: 'endnote text',
  44: 'table of authorities', 45: 'macro', 46: 'toa heading', 47: 'List', 48: 'List Bullet', 49: 'List Number',
  50: 'List 2', 51: 'List 3', 52: 'List 4', 53: 'List 5', 54: 'List Bullet 2', 55: 'List Bullet 3',
  56: 'List Bullet 4', 57: 'List Bullet 5', 58: 'List Number 2', 59: 'List Number 3', 60: 'List Number 4',
  61: 'List Number 5', 62: 'Title', 63: 'Closing', 64: 'Signature', 65: 'Default Paragraph Font', 66: 'Body Text',
  67: 'Body Text Indent', 68: 'List Continue', 69: 'List Continue 2', 70: 'List Continue 3', 71: 'List Continue 4',
  72: 'List Continue 5', 73: 'Message Header', 74: 'Subtitle', 75: 'Salutation', 76: 'Date',
  77: 'Body Text First Indent', 78: 'Body Text First Indent 2', 79: 'Note Heading', 80: 'Body Text 2',
  81: 'Body Text 3', 82: 'Body Text Indent 2', 83: 'Body Text Indent 3', 84: 'Block Text', 85: 'Hyperlink',
  86: 'FollowedHyperlink', 87: 'Strong', 88: 'Emphasis', 89: 'Document Map', 90: 'Plain Text', 91: 'E-mail Signature',
  92: 'HTML Top of Form', 93: 'HTML Bottom of Form', 94: 'Normal (Web)', 95: 'HTML Acronym', 96: 'HTML Address',
  97: 'HTML Cite', 98: 'HTML Code', 99: 'HTML Definition', 100: 'HTML Keyboard', 101: 'HTML Preformatted',
  102: 'HTML Sample', 103: 'HTML Typewriter', 104: 'HTML Variable', 105: 'Normal Table', 106: 'annotation subject',
  107: 'No List', 153: 'Balloon Text', 154: 'Table Grid', 156: 'Placeholder Text', 157: 'No Spacing',
  158: 'Light Shading', 159: 'Light List', 160: 'Light Grid', 161: 'Medium Shading 1', 162: 'Medium Shading 2',
  163: 'Medium List 1', 164: 'Medium List 2', 165: 'Medium Grid 1', 166: 'Medium Grid 2', 167: 'Medium Grid 3',
  168: 'Dark List', 169: 'Colorful Shading', 170: 'Colorful List', 171: 'Colorful Grid', 172: 'Light Shading Accent 1',
  173: 'Light List Accent 1', 174: 'Light Grid Accent 1', 175: 'Medium Shading 1 Accent 1',
  176: 'Medium Shading 2 Accent 1', 177: 'Medium List 1 Accent 1', 178: 'Revision', 179: 'List Paragraph',
  180: 'Quote', 181: 'Intense Quote', 182: 'Medium List 2 Accent 1', 183: 'Medium Grid 1 Accent 1',
  184: 'Medium Grid 2 Accent 1', 185: 'Medium Grid 3 Accent 1', 186: 'Dark List Accent 1',
  187: 'Colorful Shading Accent 1', 188: 'Colorful List Accent 1', 189: 'Colorful Grid Accent 1',
  190: 'Light Shading Accent 2', 191: 'Light List Accent 2', 192: 'Light Grid Accent 2',
  193: 'Medium Shading 1 Accent 2', 194: 'Medium Shading 2 Accent 2', 195: 'Medium List 1 Accent 2',
  196: 'Medium List 2 Accent 2', 197: 'Medium Grid 1 Accent 2', 198: 'Medium Grid 2 Accent 2',
  199: 'Medium Grid 3 Accent 2', 200: 'Dark List Accent 2', 201: 'Colorful Shading Accent 2',
  202: 'Colorful List Accent 2', 203: 'Colorful Grid Accent 2', 204: 'Light Shading Accent 3',
  205: 'Light List Accent 3', 206: 'Light Grid Accent 3', 207: 'Medium Shading 1 Accent 3',
  208: 'Medium Shading 2 Accent 3', 209: 'Medium List 1 Accent 3', 210: 'Medium List 2 Accent 3',
  211: 'Medium Grid 1 Accent 3', 212: 'Medium Grid 2 Accent 3', 213: 'Medium Grid 3 Accent 3',
  214: 'Dark List Accent 3', 215: 'Colorful Shading Accent 3', 216: 'Colorful List Accent 3',
  217: 'Colorful Grid Accent 3', 218: 'Light Shading Accent 4', 219: 'Light List Accent 4', 220: 'Light Grid Accent 4',
  221: 'Medium Shading 1 Accent 4', 222: 'Medium Shading 2 Accent 4', 223: 'Medium List 1 Accent 4',
  224: 'Medium List 2 Accent 4', 225: 'Medium Grid 1 Accent 4', 226: 'Medium Grid 2 Accent 4',
  227: 'Medium Grid 3 Accent 4', 228: 'Dark List Accent 4', 229: 'Colorful Shading Accent 4',
  230: 'Colorful List Accent 4', 231: 'Colorful Grid Accent 4', 232: 'Light Shading Accent 5',
  233: 'Light List Accent 5', 234: 'Light Grid Accent 5', 235: 'Medium Shading 1 Accent 5',
  236: 'Medium Shading 2 Accent 5', 237: 'Medium List 1 Accent 5', 238: 'Medium List 2 Accent 5',
  239: 'Medium Grid 1 Accent 5', 240: 'Medium Grid 2 Accent 5', 241: 'Medium Grid 3 Accent 5',
  242: 'Dark List Accent 5', 243: 'Colorful Shading Accent 5', 244: 'Colorful List Accent 5',
  245: 'Colorful Grid Accent 5', 246: 'Light Shading Accent 6', 247: 'Light List Accent 6', 248: 'Light Grid Accent 6',
  249: 'Medium Shading 1 Accent 6', 250: 'Medium Shading 2 Accent 6', 251: 'Medium List 1 Accent 6',
  252: 'Medium List 2 Accent 6', 253: 'Medium Grid 1 Accent 6', 254: 'Medium Grid 2 Accent 6',
  255: 'Medium Grid 3 Accent 6', 256: 'Dark List Accent 6', 257: 'Colorful Shading Accent 6',
  258: 'Colorful List Accent 6', 259: 'Colorful Grid Accent 6', 260: 'Subtle Emphasis', 261: 'Intense Emphasis',
  262: 'Subtle Reference', 263: 'Intense Reference', 264: 'Book Title', 265: 'Bibliography', 266: 'TOC Heading',
  374: 'Unresolved Mention',
};

// Prm0 куска ([MS-DOC] 2.9.215): isprm → код sprm (все с операндом в 1 байт; прочие isprm недопустимы).
// Выписано из спецификации (tests/out/specs/doc_fixes.md); isprm совпадают с номерами sprm Word 6.
const PRM0 = {
  0x04: 0x2602, 0x05: 0x2461, 0x07: 0x2405, 0x08: 0x2406, 0x09: 0x2407, 0x0C: 0x260A, 0x0D: 0x2470, 0x0E: 0x240C, 0x0F: 0x2471,
  0x18: 0x2416, 0x19: 0x2417, 0x1D: 0x261B, 0x25: 0x2423, 0x2C: 0x242A, 0x32: 0x2430, 0x33: 0x2431, 0x35: 0x2433, 0x36: 0x2434,
  0x37: 0x2435, 0x38: 0x2436, 0x39: 0x2437, 0x3A: 0x2438, 0x78: 0x2640, 0x7E: 0x2443,
  0x41: 0x0800, 0x42: 0x0801, 0x43: 0x0802, 0x47: 0x0806, 0x4B: 0x080A, 0x4D: 0x2A0C, 0x4E: 0x0858, 0x4F: 0x2859, 0x50: 0x0811,
  0x51: 0x0818, 0x53: 0x2A33, 0x55: 0x0835, 0x56: 0x0836, 0x57: 0x0837, 0x58: 0x0838, 0x59: 0x0839, 0x5A: 0x083A, 0x5B: 0x083B,
  0x5C: 0x083C, 0x5E: 0x2A3E, 0x62: 0x2A42, 0x68: 0x2A48, 0x73: 0x2A53, 0x74: 0x0854, 0x75: 0x0855, 0x76: 0x0856, 0x7B: 0x2A90,
  0x7C: 0x2A86,
};

// ---------- sprm ----------
// размер операнда по spra (биты 13–15): 0,1 — 1; 2 — 2; 3 — 4; 4,5 — 2; 6 — переменный; 7 — 3
function sprms(g, from = 0, to = g.length) {
  const out = [];
  let p = from;
  while (p + 2 <= to) {
    const sprm = u16(g, p); const spra = sprm >> 13; p += 2;
    let n;
    if (spra === 6) {
      if (sprm === 0xD608 || sprm === 0xD606) { n = u16(g, p) + 1; } // sprmTDefTable: cb (u16) = остаток + 1
      else if (sprm === 0xC615) { n = g[p] === 255 ? 0 : g[p] + 1; } // sprmPChgTabs (сложный вид не встречается в PAPX)
      else { n = g[p] + 1; }
      out.push({ sprm, at: spra === 6 && sprm !== 0xD608 && sprm !== 0xD606 ? p + 1 : p + 2, len: sprm === 0xD608 || sprm === 0xD606 ? n - 2 : n - 1 });
      p += n;
      continue;
    }
    n = [1, 1, 2, 4, 2, 2, 0, 3][spra];
    out.push({ sprm, at: p, len: n });
    p += n;
  }
  return out.map((x) => ({ ...x, g }));
}
const opU8 = (s) => s.g[s.at];
const opU16 = (s) => u16(s.g, s.at);
const opI16 = (s) => i16(s.g, s.at);

// DTTM (дата-время Word: минуты, часы, день, месяц, год с 1900) → ISO, как w:date
const dttm = (v) => { if (!v) return '2000-01-01T00:00:00Z'; const min = v & 0x3F; const hr = (v >> 6) & 0x1F; const dom = (v >> 11) & 0x1F; const mon = (v >> 16) & 0x0F; const yr = 1900 + ((v >> 20) & 0x1FF); const z = (n) => String(n).padStart(2, '0'); return `${yr}-${z(mon)}-${z(dom)}T${z(hr)}:${z(min)}:00Z`; };

// ---------- STTB ----------
function sttb(d, fc, lcb) {
  if (!lcb) return { strs: [], extra: [] };
  let p = fc; const ext = u16(d, p) === 0xFFFF; if (ext) p += 2;
  const n = u16(d, p); const cbExtra = u16(d, p + 2); p += 4;
  const strs = []; const extra = [];
  const end = Math.min(fc + lcb, d.length);
  for (let i = 0; i < n && p < end; i++) {
    if (ext) { const cch = u16(d, p); strs.push(utf16(d, p + 2, cch)); p += 2 + cch * 2; } else { const cch = d[p]; strs.push(String.fromCharCode(...d.subarray(p + 1, p + 1 + cch))); p += 1 + cch; }
    extra.push(d.subarray(p, p + cbExtra)); p += cbExtra;
  }
  return { strs, extra };
}
// PLC: (n + 1) CP и n записей по cb байт
function plc(d, fc, lcb, cb) {
  // длина из FIB — не больше потока (враждебный lcb иначе даёт сотни миллионов итераций)
  if (!lcb || lcb < 4 || fc + lcb > d.length) return { cps: [], data: [] };
  const n = Math.floor(cb ? (lcb - 4) / (4 + cb) : (lcb - 4) / 4);
  const cps = []; for (let i = 0; i <= n; i++) cps.push(i32(d, fc + i * 4));
  const data = []; for (let i = 0; i < n; i++) data.push(d.subarray(fc + (n + 1) * 4 + i * cb, fc + (n + 1) * 4 + (i + 1) * cb));
  return { cps, data };
}

// Набор свойств OLE ([MS-OLEPS]): первый раздел → { pid: строка | Date | число }
function propertySet(d) {
  const out = {};
  if (!d || d.length < 48) return out;
  const off = u32(d, 44); if (off + 8 > d.length) return out;
  const n = u32(d, off + 4);
  const items = []; for (let i = 0; i < n; i++) items.push([u32(d, off + 8 + i * 8), u32(d, off + 12 + i * 8)]);
  let cp = 1252;
  const cpItem = items.find(([pid]) => pid === 1); if (cpItem) cp = u16(d, off + cpItem[1] + 4);
  let dec; try { dec = new TextDecoder(cp === 65001 ? 'utf-8' : `windows-${cp}`); } catch { dec = new TextDecoder('windows-1252'); }
  for (const [pid, o] of items) {
    const at = off + o; if (at + 4 > d.length) continue;
    const type = u32(d, at);
    if (type === 0x1E) { const cch = u32(d, at + 4); const s = dec.decode(d.subarray(at + 8, at + 8 + cch)).replace(/\0+$/, ''); if (s) out[pid] = s; }
    else if (type === 0x1F) { const cch = u32(d, at + 4); const s = utf16(d, at + 8, cch).replace(/\0+$/, ''); if (s) out[pid] = s; }
    else if (type === 0x40) { const lo = u32(d, at + 4); const hi = u32(d, at + 8); const ms = (hi * 4294967296 + lo) / 10000 - 11644473600000; if (ms > 0 && hi) out[pid] = new Date(Math.round(ms / 1000) * 1000); }
    else if (type === 0x03) out[pid] = i32(d, at + 4);
  }
  return out;
}

// Формула из файла → проверенный OMML: разбор XML, корень — m:oMath / m:oMathPara, внутри только элементы m: и w:, без
// инструкций обработки; затем сериализация заново (разметка из файла не может закрыть абзац и вставить своё). Иначе — null.
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function safeOmml(frag) {
  let doc; let failed = false;
  try {
    doc = new DOMParser({ onError: () => { failed = true; } }).parseFromString(`<r xmlns:m="${M_NS}" xmlns:w="${W_NS}">${frag}</r>`, 'text/xml');
  } catch { return null; }
  if (failed || !doc || !doc.documentElement) return null;
  const kids = [...doc.documentElement.childNodes].filter((n) => n.nodeType === 1);
  if (kids.length !== 1 || !/^m:oMath(Para)?$/.test(kids[0].nodeName)) return null;
  const ok = (n) => {
    if (n.nodeType === 7 || n.nodeType === 10) return false; // инструкции обработки, DOCTYPE
    if (n.nodeType === 1 && (!/^(m|w):/.test(n.nodeName) || (n.namespaceURI !== M_NS && n.namespaceURI !== W_NS))) return false;
    return [...(n.childNodes || [])].every(ok);
  };
  if (!ok(kids[0])) return null;
  return new XMLSerializer().serializeToString(kids[0]).replace(/\s+xmlns:(m|w)="[^"]*"/g, '');
}

// ---------- OfficeArt ([MS-ODRAW]) ----------
// записи: { type, inst, ver, at (данные), len, kids (у контейнеров) } — все подряд, с вложенными
function artRecords(d, a, b) {
  const out = [];
  const walk = (x, y, parent, depth) => {
    y = Math.min(y, d.length); // границы из файла — не дальше буфера
    if (depth > 64) return; // глубже Word не вкладывает; враждебная вложенность — не до переполнения стека
    while (x + 8 <= y) {
      const vi = u16(d, x); const type = u16(d, x + 2); const len = u32(d, x + 4);
      if (type < 0xF000 || type > 0xF200 || x + 8 + len > y) { x++; continue; } // в DggInfo перед рисунком — байт dgglbl
      const r = { type, inst: vi >> 4, ver: vi & 0xF, at: x + 8, len, kids: [] };
      out.push(r); if (parent) parent.kids.push(r);
      if (r.ver === 0xF) walk(x + 8, x + 8 + len, r, depth + 1);
      x += 8 + len;
    }
  };
  walk(a, b, null, 0);
  return out;
}
// свойства фигуры (OfficeArtFOPT и вторичные): pid → { val, str (сложные — строки UTF-16) }
function shapeOpt(d, sp) {
  const m = new Map();
  for (const k of sp.kids) {
    if (k.type !== 0xF00B && k.type !== 0xF121 && k.type !== 0xF122) continue;
    const end = Math.min(k.at + k.len, d.length);
    const n = Math.min(k.inst, Math.floor((end - k.at) / 6)); let cx = k.at + n * 6;
    for (let i = 0; i < n; i++) {
      const id = u16(d, k.at + i * 6); const val = u32(d, k.at + i * 6 + 2);
      const pid = id & 0x3FFF; const complex = !!(id & 0x8000);
      const e = { val };
      if (complex) {
        const len = Math.min(val, Math.max(0, end - cx)); // длина сложного свойства — не дальше записи
        e.raw = d.subarray(cx, cx + len); e.str = utf16(d, cx, Math.floor(len / 2)).replace(/\0[\s\S]*$/, ''); cx += len;
      }
      if (!m.has(pid)) m.set(pid, e);
    }
  }
  return m;
}
// OfficeArtCOLORREF (r, g, b, флаги) → 'RRGGBB'; цвета схемы/системы не разбираются
const artColor = (v) => ((v >>> 24) & 0x18 ? null : hex6(((v & 0xFF) << 16) | (v & 0xFF00) | ((v >> 16) & 0xFF)));
// обрезка: доли 16.16 со знаком (сверху, снизу, слева, справа)
const cropOf = (o) => { const f = (pid) => (o.has(pid) ? (o.get(pid).val | 0) / 65536 : 0); return { t: f(0x0100), b: f(0x0101), l: f(0x0102), r: f(0x0103) }; };
const cropXml = (o) => {
  const c = cropOf(o);
  if (!c.t && !c.b && !c.l && !c.r) return '';
  const p = (v) => Math.round(v * 100000);
  return `<a:srcRect${c.l ? ` l="${p(c.l)}"` : ''}${c.t ? ` t="${p(c.t)}"` : ''}${c.r ? ` r="${p(c.r)}"` : ''}${c.b ? ` b="${p(c.b)}"` : ''}/>`;
};
// картинка OfficeArtBlip: заголовок, 1–2 UID (нечётный inst — два), у растровых — байт tag, дальше данные
function readBlip(d, a) {
  const vi = u16(d, a); const type = u16(d, a + 2); const len = u32(d, a + 4); const inst = vi >> 4;
  const kind = { 0xF01D: 'jpeg', 0xF02A: 'jpeg', 0xF01E: 'png', 0xF01F: 'bmp', 0xF029: 'tiff', 0xF01A: 'emf', 0xF01B: 'wmf', 0xF01C: 'pict' }[type];
  if (!kind) return null;
  if (kind === 'emf' || kind === 'wmf' || kind === 'pict') return { kind, data: null };
  const uids = inst & 1 ? 32 : 16;
  let data = d.slice(a + 8 + uids + 1, a + 8 + len);
  if (kind === 'bmp') { // DIB без заголовка файла — дописать BITMAPFILEHEADER
    const out = new Uint8Array(data.length + 14); out[0] = 0x42; out[1] = 0x4D;
    const dv = new DataView(out.buffer); dv.setUint32(2, out.length, true);
    const hdr = u32(data, 0); const bpp = u16(data, 14); const clr = u32(data, 32) || (bpp <= 8 ? 1 << bpp : 0);
    dv.setUint32(10, 14 + hdr + clr * 4, true); out.set(data, 14); data = out;
  }
  return { kind, data };
}
// OfficeArtFBSE: картинка внутри записи или в потоке WordDocument по foDelay
function bseBlip(d, r, delay) {
  const size = u32(d, r.at + 20); const foDelay = u32(d, r.at + 28); const cbName = d[r.at + 33];
  const inner = r.at + 36 + cbName;
  if (inner + 8 <= r.at + r.len) return readBlip(d, inner);
  if (size && foDelay !== 0xFFFFFFFF && foDelay + 8 <= delay.length) return readBlip(delay, foDelay);
  return null;
}

// ---------- основной разбор ----------
export async function docToDocx(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const cfb = readCfb(bytes);
  const w = cfb.stream('WordDocument');
  if (!w || u16(w, 0) !== 0xA5EC) throw new Error('Это не документ Word 97–2003: нет потока WordDocument');
  if (u16(w, 2) < 0xC1) throw new Error('Документ Word 6.0/95 (или старше) — Перо открывает документы Word 97–2003 и новее');
  const flags = u16(w, 10);
  if (flags & 0x0100) throw new Error('Документ защищён паролем (зашифрован) — Перо такие пока не открывает');
  const T = cfb.stream(flags & 0x0200 ? '1Table' : '0Table');
  if (!T) throw new Error('В документе нет потока таблиц (0Table/1Table)');
  const D = cfb.stream('Data') || new Uint8Array(0);
  const warnings = {};
  const warn = (k) => { warnings[k] = (warnings[k] || 0) + 1; };

  // FIB
  const ccp = { text: u32(w, 76), ftn: u32(w, 80), hdd: u32(w, 84), atn: u32(w, 92), edn: u32(w, 96), txbx: u32(w, 100), hdrTxbx: u32(w, 104) };
  const nFcLcb = u16(w, 152);
  // пара fc/lcb из FIB; часть за пределами потока таблиц — как отсутствующая
  const fcl = (i) => {
    if (i >= nFcLcb) return [0, 0];
    const fc = u32(w, 154 + i * 8); const lcb = u32(w, 158 + i * 8);
    return fc + lcb > T.length ? [0, 0] : [fc, lcb];
  };
  const [fcStshf, lcbStshf] = fcl(1); const [fcFndRef, lcbFndRef] = fcl(2); const [fcFndTxt, lcbFndTxt] = fcl(3);
  const [fcAndRef, lcbAndRef] = fcl(4); const [fcAndTxt, lcbAndTxt] = fcl(5); const [fcSed, lcbSed] = fcl(6);
  const [fcHdd, lcbHdd] = fcl(11); const [fcBteChpx, lcbBteChpx] = fcl(12); const [fcBtePapx, lcbBtePapx] = fcl(13);
  const [fcFfn, lcbFfn] = fcl(15); const [fcBkmk, lcbBkmk] = fcl(21); const [fcBkf, lcbBkf] = fcl(22); const [fcBkl, lcbBkl] = fcl(23);
  const [fcDop, lcbDop] = fcl(31); const [fcClx, lcbClx] = fcl(33); const [fcAtnOwn, lcbAtnOwn] = fcl(36);
  const [fcAtnBkmk, lcbAtnBkmk] = fcl(37); const [fcAtnBkf, lcbAtnBkf] = fcl(42); const [fcAtnBkl, lcbAtnBkl] = fcl(43);
  const [fcEndRef, lcbEndRef] = fcl(46); const [fcRMark, lcbRMark] = fcl(51); const [fcLst, lcbLst] = fcl(73); const [fcLfo, lcbLfo] = fcl(74);
  const [fcAtrdExtra, lcbAtrdExtra] = fcl(112); // FibRgFcLcb2002: даты примечаний (ATRDPost10, 18 байт)
  if (lcbEndRef) warn('концевые сноски');

  // куски текста (Clx: Prc* + Pcdt)
  const prcs = []; let p = fcClx; const pieces = [];
  const bad = (what) => { throw new Error(`Документ повреждён (${what}) — Перо не может его прочитать`); };
  const clxEnd = fcClx + lcbClx;
  while (p < clxEnd) {
    if (T[p] === 0x01) {
      const cb = i16(T, p + 1);
      if (cb < 0 || p + 3 + cb > clxEnd) bad('Clx');
      prcs.push(T.subarray(p + 3, p + 3 + cb)); p += 3 + cb; continue;
    }
    if (T[p] === 0x02) {
      const lcb = u32(T, p + 1); const base = p + 5;
      if (lcb < 4 || base + lcb > clxEnd) bad('таблица кусков');
      const n = Math.floor((lcb - 4) / 12);
      for (let i = 0; i < n; i++) {
        const cp0 = u32(T, base + i * 4); const cp1 = u32(T, base + (i + 1) * 4);
        const pcd = base + (n + 1) * 4 + i * 8;
        const fcRaw = u32(T, pcd + 2); const compressed = !!(fcRaw & 0x40000000);
        const fc = compressed ? (fcRaw & 0x3FFFFFFF) / 2 : fcRaw;
        // куски — по возрастанию CP и целиком внутри потока WordDocument
        if (cp1 < cp0 || (!pieces.length && cp0 !== 0) || fc + (cp1 - cp0) * (compressed ? 1 : 2) > w.length) bad('кусок текста');
        pieces.push({ cp0, cp1, fc, compressed, prm: u16(T, pcd + 6) });
      }
      break;
    }
    break;
  }
  if (!pieces.length) bad('нет текста');
  const dec1252 = new TextDecoder('windows-1252');
  const total = pieces[pieces.length - 1].cp1;
  // длины историй из FIB — в пределах текста (по порядку: основной, сноски, колонтитулы, примечания, концевые, надписи)
  { let acc = 0; for (const k of ['text', 'ftn', 'hdd', 'atn', 'edn', 'txbx', 'hdrTxbx']) { ccp[k] = Math.max(0, Math.min(ccp[k], total - acc)); acc += ccp[k]; } }
  let text = '';
  for (const pc of pieces) {
    const n = pc.cp1 - pc.cp0;
    text += pc.compressed ? dec1252.decode(w.subarray(pc.fc, pc.fc + n)) : utf16(w, pc.fc, n);
  }
  const pieceAt = (cp) => { let lo = 0; let hi = pieces.length - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (pieces[m].cp0 <= cp) lo = m; else hi = m - 1; } return pieces[lo]; };
  const fcOf = (cp) => { const pc = pieceAt(cp); return pc.fc + (cp - pc.cp0) * (pc.compressed ? 1 : 2); };

  // FKP: свойства символов (CHPX) и абзацев (PAPX) по диапазонам FC
  const chpx = []; const papx = [];
  {
    // повторная ссылка на ту же страницу FKP (раздувание из 8 байт в сотню объектов) и страница за файлом — пропуск
    const pages = new Set(); const fresh = (page) => page + 512 <= w.length && !pages.has(page) && pages.add(page);
    const bc = plc(T, fcBteChpx, lcbBteChpx, 4);
    for (const pn of bc.data) {
      const page = (u32(pn, 0) & 0x3FFFFF) * 512; if (!fresh(page)) continue; const crun = w[page + 511];
      for (let i = 0; i < crun; i++) {
        const off = w[page + (crun + 1) * 4 + i] * 2;
        const g = off ? w.subarray(page + off + 1, page + off + 1 + w[page + off]) : new Uint8Array(0);
        chpx.push({ fc0: u32(w, page + i * 4), fc1: u32(w, page + (i + 1) * 4), s: sprms(g) });
      }
    }
    const bp = plc(T, fcBtePapx, lcbBtePapx, 4);
    for (const pn of bp.data) {
      const page = (u32(pn, 0) & 0x3FFFFF) * 512; if (!fresh(page)) continue; const cpara = w[page + 511];
      for (let i = 0; i < cpara; i++) {
        const off = w[page + (cpara + 1) * 4 + i * 13] * 2;
        let istd = 0; let g = new Uint8Array(0);
        if (off) {
          let cb = w[page + off]; let q = page + off + 1;
          if (cb === 0) { cb = w[q] * 2; q++; } else cb = cb * 2 - 1;
          istd = u16(w, q); g = w.subarray(q + 2, q + cb);
        }
        // sprmPHugePapx: свойства абзаца не поместились в FKP — они в потоке Data (PrcData: cbGrpprl u16 + GrpPrl)
        let s = sprms(g);
        if (s[0] && s[0].sprm === 0x6646) {
          const off = u32(s[0].g, s[0].at); const cb = u16(D, off);
          s = off + 2 + cb <= D.length ? sprms(D.subarray(off + 2, off + 2 + cb)) : [];
        }
        papx.push({ fc0: u32(w, page + i * 4), fc1: u32(w, page + (i + 1) * 4), istd, s });
      }
    }
  }
  const find = (list, fc) => { let lo = 0; let hi = list.length - 1; while (lo <= hi) { const m = (lo + hi) >> 1; if (list[m].fc1 <= fc) lo = m + 1; else if (list[m].fc0 > fc) hi = m - 1; else return list[m]; } return null; };
  // Prm куска: один sprm (Prm0) или набор из Prc (Prm1)
  const prmSprms = (prm) => {
    if (!prm) return [];
    if (prm & 1) { const g = prcs[prm >> 1]; return g ? sprms(g) : []; }
    const isprm = (prm >> 1) & 0x7F; const val = prm >> 8;
    const sp = PRM0[isprm]; if (!sp) return [];
    return sprms(Uint8Array.from([sp & 0xFF, sp >> 8, val & 0xFF])); // у всех sprm таблицы операнд — 1 байт
  };

  // шрифты
  const fonts = [];
  if (lcbFfn) {
    let q = fcFfn + 4; const n = u16(T, fcFfn);
    for (let i = 0; i < n; i++) { const cb = T[q]; const nameAt = q + 1 + 39; let e = nameAt; while (e < q + 1 + cb && u16(T, e)) e += 2; fonts.push(utf16(T, nameAt, (e - nameAt) / 2)); q += 1 + cb; }
  }
  const authors = sttb(T, fcRMark, lcbRMark).strs;

  // ---------- свойства → OOXML ----------
  const brc8 = (s, o = s.at) => { const g = s.g; const t = g[o + 5]; if (!t || t === 0xFF) return { val: 'nil' }; return { val: BRC[t] || 'single', sz: g[o + 4], space: g[o + 6] & 0x1F, color: colorref(g, o) }; };
  const brc80 = (g, o) => { const t = g[o + 1]; if (!t || t === 0xFF) return null; return { val: BRC[t] || 'single', sz: g[o], space: g[o + 3] & 0x1F, color: ICO[g[o + 2]] || 'auto' }; };
  const brcXml = (name, b) => (!b ? '' : b.val === 'nil' ? `<w:${name} w:val="nil"/>` : `<w:${name} w:val="${b.val}" w:sz="${b.sz}" w:space="${b.space}" w:color="${b.color}"/>`);
  const shd10 = (g, o) => { const back = colorref(g, o + 4); const fore = colorref(g, o); const pat = u16(g, o + 8); if (pat === 0xFFFF) return null; return `<w:shd w:val="${SHD_PAT[pat] || 'clear'}" w:color="${fore}" w:fill="${back}"/>`; };
  const shd80 = (v) => { const fore = ICO[v & 0x1F] || 'auto'; const back = ICO[(v >> 5) & 0x1F] || 'auto'; const pat = v >> 10; return `<w:shd w:val="${SHD_PAT[pat] || 'clear'}" w:color="${fore}" w:fill="${back}"/>`; };

  // свойства символа: sprm → { элемент: xml }; toggle(name) — значение стиля для 0x81
  function runProps(list, styleToggle = () => false) {
    const r = {};
    const tog = (name, v) => { if (v === 0x80) return; const on = v === 0x81 ? !styleToggle(name) : !!v; r[name] = on ? `<w:${name}/>` : `<w:${name} w:val="0"/>`; };
    for (const s of list) {
      switch (s.sprm) {
        case 0x0835: tog('b', opU8(s)); break;
        case 0x0836: tog('i', opU8(s)); break;
        case 0x0837: tog('strike', opU8(s)); break;
        case 0x2A53: tog('dstrike', opU8(s)); break;
        case 0x0838: tog('outline', opU8(s)); break;
        case 0x0839: tog('shadow', opU8(s)); break;
        case 0x083A: tog('smallCaps', opU8(s)); break;
        case 0x083B: tog('caps', opU8(s)); break;
        case 0x083C: tog('vanish', opU8(s)); break;
        case 0x0858: tog('emboss', opU8(s)); break;
        case 0x0854: tog('imprint', opU8(s)); break;
        case 0x0875: tog('noProof', opU8(s)); break;
        case 0x4A43: r.sz = `<w:sz w:val="${opU16(s)}"/>`; break;
        case 0x4A61: r.szCs = `<w:szCs w:val="${opU16(s)}"/>`; break;
        case 0x2A42: { const c = ICO[opU8(s)]; r.color = `<w:color w:val="${c || 'auto'}"/>`; break; }
        case 0x6870: r.color = `<w:color w:val="${colorref(s.g, s.at)}"/>`; break;
        case 0x2A0C: { const h = HIGHLIGHT[opU8(s)]; if (h) r.highlight = `<w:highlight w:val="${h}"/>`; else delete r.highlight; break; }
        case 0x2A3E: { const k = opU8(s); r.u = k ? `<w:u w:val="${KUL[k] || 'single'}"/>` : '<w:u w:val="none"/>'; break; }
        case 0x6877: r.ucolor = colorref(s.g, s.at); break;
        case 0x2A48: { const v = opU8(s); r.vertAlign = v === 1 ? '<w:vertAlign w:val="superscript"/>' : v === 2 ? '<w:vertAlign w:val="subscript"/>' : '<w:vertAlign w:val="baseline"/>'; break; }
        case 0x4A4F: r.f0 = fonts[opU16(s)]; break;
        case 0x4A50: r.f1 = fonts[opU16(s)]; break;
        case 0x4A51: r.f2 = fonts[opU16(s)]; break;
        case 0x4A5E: r.f3 = fonts[opU16(s)]; break;
        case 0x8840: r.spacing = `<w:spacing w:val="${opI16(s)}"/>`; break;
        case 0x4845: r.position = `<w:position w:val="${opI16(s)}"/>`; break;
        case 0x4A30: r.istd = opU16(s); break;
        case 0x4866: r.shd = shd80(opU16(s)); break;
        case 0xCA71: { const x = shd10(s.g, s.at); if (x) r.shd = x; break; }
        case 0x0801: r.ins = opU8(s); break;
        case 0x0800: r.del = opU8(s); break;
        case 0x4804: r.ibst = opU16(s); break;
        case 0x6805: r.dttm = u32(s.g, s.at); break;
        case 0x4863: r.ibstDel = opU16(s); break;
        case 0x6864: r.dttmDel = u32(s.g, s.at); break;
        case 0x0855: r.spec = opU8(s); break;
        case 0x6A03: r.pic = u32(s.g, s.at); break;
        case 0x080A: r.ole = opU8(s); break;
        case 0x0806: r.data = opU8(s); break;
        case 0x6A09: r.sym = { font: fonts[u16(s.g, s.at)], ch: u16(s.g, s.at + 2) }; break;
        case 0x486D: r.lang = u16(s.g, s.at); break;
        default: break;
      }
    }
    return r;
  }
  const LANG = { 0x0419: 'ru-RU', 0x0409: 'en-US', 0x0809: 'en-GB', 0x0407: 'de-DE', 0x040C: 'fr-FR', 0x0422: 'uk-UA', 0x0423: 'be-BY' };
  function rPrXml(r, extraStyle) {
    const fontsXml = r.f0 || r.f1 || r.f2 || r.f3
      ? `<w:rFonts${r.f0 ? ` w:ascii="${esc(r.f0)}"` : ''}${r.f2 ? ` w:hAnsi="${esc(r.f2)}"` : r.f0 ? ` w:hAnsi="${esc(r.f0)}"` : ''}${r.f1 ? ` w:eastAsia="${esc(r.f1)}"` : ''}${r.f3 ? ` w:cs="${esc(r.f3)}"` : ''}/>` : '';
    const u = r.u && r.ucolor && r.ucolor !== 'auto' ? r.u.replace('/>', ` w:color="${r.ucolor}"/>`) : r.u;
    const style = extraStyle || (r.istd != null && styleIdOf(r.istd, 'character'));
    const parts = [style ? `<w:rStyle w:val="${style}"/>` : '', fontsXml, r.b, r.i, r.caps, r.smallCaps, r.strike, r.dstrike, r.outline, r.shadow, r.emboss, r.imprint, r.noProof, r.vanish,
      r.color, r.spacing, r.position, r.sz, r.szCs, r.highlight, u, r.shd, r.vertAlign, r.lang && LANG[r.lang] ? `<w:lang w:val="${LANG[r.lang]}"/>` : ''];
    const x = parts.filter(Boolean).join('');
    return x ? `<w:rPr>${x}</w:rPr>` : '';
  }

  // свойства абзаца
  function paraProps(list) {
    const p = {};
    for (const s of list) {
      switch (s.sprm) {
        case 0x4600: p.istd = opU16(s); break;
        case 0x2403: case 0x2461: p.jc = ['left', 'center', 'right', 'both', 'distribute'][opU8(s)] || 'left'; break;
        case 0x2405: p.keepLines = opU8(s); break;
        case 0x2406: p.keepNext = opU8(s); break;
        case 0x2407: p.pageBreakBefore = opU8(s); break;
        case 0x2431: p.widowControl = opU8(s); break;
        case 0x240C: p.noLineNumbers = opU8(s); break;
        case 0x242A: p.noHyphens = opU8(s); break;
        case 0x260A: p.ilvl = opU8(s); break;
        case 0x460B: p.ilfo = opU16(s); break;
        case 0x840F: case 0x845E: p.left = opI16(s); break;
        case 0x840E: case 0x845D: p.right = opI16(s); break;
        case 0x8411: case 0x8460: p.first = opI16(s); break;
        case 0xA413: p.before = opU16(s); break;
        case 0xA414: p.after = opU16(s); break;
        case 0x245B: p.beforeAuto = opU8(s); break;
        case 0x245C: p.afterAuto = opU8(s); break;
        case 0x6412: p.line = { dya: i16(s.g, s.at), mult: i16(s.g, s.at + 2) }; break;
        case 0x246D: p.contextual = opU8(s); break;
        case 0x2416: p.inTable = opU8(s); break;
        case 0x2417: p.ttp = opU8(s); break;
        case 0x6649: p.itap = u32(s.g, s.at); break;
        case 0x244B: p.innerCell = opU8(s); break;
        case 0x244C: p.innerTtp = opU8(s); break;
        case 0x2640: p.outline = opU8(s); break;
        case 0xC64E: p.bTop = brc8(s); break;
        case 0xC64F: p.bLeft = brc8(s); break;
        case 0xC650: p.bBottom = brc8(s); break;
        case 0xC651: p.bRight = brc8(s); break;
        case 0xC652: p.bBetween = brc8(s); break;
        case 0x6424: p.bTop ||= brc80(s.g, s.at); break;
        case 0x6425: p.bLeft ||= brc80(s.g, s.at); break;
        case 0x6426: p.bBottom ||= brc80(s.g, s.at); break;
        case 0x6427: p.bRight ||= brc80(s.g, s.at); break;
        case 0x6428: p.bBetween ||= brc80(s.g, s.at); break;
        case 0x442D: p.shd ||= shd80(opU16(s)); break;
        case 0xC64D: { const x = shd10(s.g, s.at); if (x) p.shd = x; break; }
        case 0xC60D: { // sprmPChgTabsPapx: удалённые и добавленные позиции
          const g = s.g; let o = s.at; const nDel = g[o]; o++;
          const del = []; for (let k = 0; k < nDel; k++) del.push(i16(g, o + k * 2)); o += nDel * 2;
          const nAdd = g[o]; o++;
          const add = []; for (let k = 0; k < nAdd; k++) add.push({ pos: i16(g, o + k * 2), tbd: g[o + nAdd * 2 + k] });
          p.tabs = { del, add };
          break;
        }
        case 0xC615: { // sprmPChgTabs: удалённые (с допуском rgdxaClose) и добавленные позиции
          const g = s.g; let o = s.at; const nDel = g[o]; o++;
          const del = []; for (let k = 0; k < nDel; k++) del.push(i16(g, o + k * 2)); o += nDel * 4;
          const nAdd = g[o]; o++;
          const add = []; for (let k = 0; k < nAdd; k++) add.push({ pos: i16(g, o + k * 2), tbd: g[o + nAdd * 2 + k] });
          p.tabs = { del, add };
          break;
        }
        default: break;
      }
    }
    return p;
  }
  const TAB_JC = ['left', 'center', 'right', 'decimal', 'bar'];
  const TAB_TLC = [null, 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];
  function pPrXml(p, { style, sectPr, numPr, rPr = '' } = {}) {
    const ON1 = (v, name) => (v == null ? '' : v ? `<w:${name}/>` : `<w:${name} w:val="0"/>`);
    let num = numPr;
    if (num == null && p.ilfo != null) num = p.ilfo === 0 ? '<w:numPr><w:numId w:val="0"/></w:numPr>' : `<w:numPr>${p.ilvl != null ? `<w:ilvl w:val="${p.ilvl}"/>` : ''}<w:numId w:val="${p.ilfo}"/></w:numPr>`;
    else if (num == null && p.ilvl != null) num = `<w:numPr><w:ilvl w:val="${p.ilvl}"/></w:numPr>`;
    const bdr = [brcXml('top', p.bTop), brcXml('left', p.bLeft), brcXml('bottom', p.bBottom), brcXml('right', p.bRight), brcXml('between', p.bBetween)].join('');
    let tabs = '';
    if (p.tabs && (p.tabs.del.length || p.tabs.add.length)) {
      const list = [...p.tabs.del.filter((x) => !p.tabs.add.some((a) => a.pos === x)).map((x) => ({ pos: x, clear: true })), ...p.tabs.add].sort((a, b) => a.pos - b.pos);
      tabs = `<w:tabs>${list.map((t) => (t.clear ? `<w:tab w:val="clear" w:pos="${t.pos}"/>` : `<w:tab w:val="${TAB_JC[t.tbd & 7] || 'left'}"${TAB_TLC[(t.tbd >> 3) & 7] ? ` w:leader="${TAB_TLC[(t.tbd >> 3) & 7]}"` : ''} w:pos="${t.pos}"/>`)).join('')}</w:tabs>`;
    }
    let spacing = '';
    if (p.before != null || p.after != null || p.line || p.beforeAuto || p.afterAuto) {
      const line = p.line ? (p.line.mult ? ` w:line="${p.line.dya}" w:lineRule="auto"` : p.line.dya < 0 ? ` w:line="${-p.line.dya}" w:lineRule="exact"` : ` w:line="${p.line.dya}" w:lineRule="atLeast"`) : '';
      spacing = `<w:spacing${p.before != null ? ` w:before="${p.before}"` : ''}${p.beforeAuto ? ' w:beforeAutospacing="1"' : ''}${p.after != null ? ` w:after="${p.after}"` : ''}${p.afterAuto ? ' w:afterAutospacing="1"' : ''}${line}/>`;
    }
    let ind = '';
    if (p.left != null || p.right != null || p.first != null) {
      ind = `<w:ind${p.left != null ? ` w:left="${p.left}"` : ''}${p.right != null ? ` w:right="${p.right}"` : ''}${p.first != null ? (p.first < 0 ? ` w:hanging="${-p.first}"` : ` w:firstLine="${p.first}"`) : ''}/>`;
    }
    const sty = style || (p.istd != null && p.istd !== 0 ? styleIdOf(p.istd, 'paragraph') : null);
    const x = [sty ? `<w:pStyle w:val="${sty}"/>` : '', ON1(p.keepNext, 'keepNext'), ON1(p.keepLines, 'keepLines'), ON1(p.pageBreakBefore, 'pageBreakBefore'),
      ON1(p.widowControl, 'widowControl'), num || '', ON1(p.noLineNumbers, 'suppressLineNumbers'), bdr ? `<w:pBdr>${bdr}</w:pBdr>` : '', p.shd || '', tabs,
      ON1(p.noHyphens, 'suppressAutoHyphens'), spacing, ind, ON1(p.contextual, 'contextualSpacing'), p.jc ? `<w:jc w:val="${p.jc}"/>` : '',
      p.outline != null && p.outline < 9 ? `<w:outlineLvl w:val="${p.outline}"/>` : '', rPr, sectPr || ''].join('');
    return x ? `<w:pPr>${x}</w:pPr>` : '';
  }

  // ---------- стили (STSH) ----------
  const styles = []; // { istd, sti, stk, name, base, next, papx, chpx }
  let ftcDefault = [0, 0, 0, 0];
  if (lcbStshf) {
    const cbStshi = u16(T, fcStshf); const sh = fcStshf + 2;
    const cstd = u16(T, sh); const cbBase = u16(T, sh + 2);
    // STSHI: cstd, cbSTDBaseInFile, флаги, stiMax, istdMaxFixed, nVerBuiltIn, rgftcStandardChpStsh (ascii, fe, прочие), ftcBi
    if (cbStshi >= 20) ftcDefault = [u16(T, sh + 12), u16(T, sh + 14), u16(T, sh + 16), u16(T, sh + 18)];
    let q = sh + cbStshi;
    for (let istd = 0; istd < cstd && q + 2 <= fcStshf + lcbStshf; istd++) {
      const cb = u16(T, q); q += 2;
      if (!cb) { styles.push(null); continue; }
      const st = q; const end = q + cb; q = end;
      const sti = u16(T, st) & 0x0FFF; const stk = u16(T, st + 2) & 0x0F; const base = u16(T, st + 2) >> 4; const next = u16(T, st + 4) >> 4;
      const cupx = u16(T, st + 4) & 0x0F;
      let o = st + cbBase;
      const cch = u16(T, o); const name = utf16(T, o + 2, cch); o += 2 + (cch + 1) * 2;
      const upx = [];
      for (let k = 0; k < cupx && o + 2 <= end; k++) { const n = u16(T, o); upx.push(T.subarray(o + 2, o + 2 + n)); o += 2 + n + (n & 1); }
      let pg = null; let cg = null;
      if (stk === 1) { pg = upx[0] ? upx[0].subarray(2) : null; cg = upx[1] || null; } else if (stk === 2) cg = upx[0] || null;
      else if (stk === 3) { pg = upx[1] ? upx[1].subarray(2) : null; cg = upx[2] || null; } else if (stk === 4) pg = upx[0] ? upx[0].subarray(2) : null;
      styles.push({ istd, sti, stk, name, base: base === 0x0FFF ? null : base, next: next === 0x0FFF ? null : next, papx: pg ? sprms(pg) : [], chpx: cg ? sprms(cg) : [] });
    }
  }
  const styleIds = new Map(); const usedIds = new Set(); const idCount = new Map();
  for (const s of styles) {
    if (!s) continue;
    const nm = STI[s.sti] && s.sti !== 4094 ? STI[s.sti] : s.name;
    s.oname = s.sti < 4094 && STI[s.sti] ? STI[s.sti] : s.name;
    let id = nm.replace(/[^A-Za-z0-9Ѐ-ӿ]/g, '') || `s${s.istd}`;
    if (/^\d/.test(id)) id = `s${id}`;
    // повтор имени — номер (счётчик на имя: тысячи одинаковых стилей не дают квадратичного перебора)
    if (usedIds.has(id)) { let k = (idCount.get(id) || 1) + 1; while (usedIds.has(`${id}_${k}`)) k++; idCount.set(id, k); id = `${id}_${k}`; } // «_»: не спутать с «Heading2»
    usedIds.add(id); styleIds.set(s.istd, id);
  }
  function styleIdOf(istd) { return styleIds.get(istd) || null; }
  // значение «переключателя» в стиле (для 0x81) — по цепочке basedOn
  // кеш: (стиль, свойство) → значение и свойства стиля → runProps один раз (цепочки до 4095 и тысячи стилей — не квадратично)
  const toggleCache = new Map(); const ownRun = new Map();
  const styleToggle = (istd) => (name) => {
    const key = `${istd}|${name}`;
    if (toggleCache.has(key)) return toggleCache.get(key);
    let val = false; const seen = new Set();
    for (let s = styles[istd]; s && !seen.has(s.istd); s = s.base != null ? styles[s.base] : null) {
      seen.add(s.istd);
      const pk = `${s.istd}|${name}`;
      if (toggleCache.has(pk)) { val = toggleCache.get(pk); break; }
      if (!ownRun.has(s.istd)) ownRun.set(s.istd, runProps(s.chpx));
      const r = ownRun.get(s.istd);
      if (r[name] != null) { val = !/w:val="0"/.test(r[name]); break; }
    }
    toggleCache.set(key, val);
    return val;
  };
  const STK = { 1: 'paragraph', 2: 'character', 3: 'table', 4: 'numbering' };
  let stylesXml = '';
  for (const s of styles) {
    if (!s || !STK[s.stk]) continue;
    const id = styleIds.get(s.istd);
    const pp = s.stk === 1 || s.stk === 3 ? paraProps(s.papx) : null;
    if (pp) delete pp.istd;
    const rr = runProps(s.chpx, s.base != null ? styleToggle(s.base) : () => false);
    delete rr.istd;
    const isDefault = (s.stk === 1 && s.istd === 0) || s.sti === 65 || s.sti === 105 || s.sti === 107;
    stylesXml += `<w:style w:type="${STK[s.stk]}"${isDefault ? ' w:default="1"' : ''}${s.sti >= 4094 || !STI[s.sti] ? ' w:customStyle="1"' : ''} w:styleId="${id}">`
      + `<w:name w:val="${esc(s.oname)}"/>${s.base != null && styleIds.has(s.base) ? `<w:basedOn w:val="${styleIds.get(s.base)}"/>` : ''}`
      + `${s.next != null && styleIds.has(s.next) && s.next !== s.istd ? `<w:next w:val="${styleIds.get(s.next)}"/>` : ''}<w:qFormat/>`
      + `${pp ? pPrXml(pp, {}) : ''}${rPrXml(rr)}</w:style>`;
  }
  const defFont = fonts[ftcDefault[0]] || 'Times New Roman';
  const stylesPart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
    + `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${esc(defFont)}" w:eastAsia="${esc(fonts[ftcDefault[1]] || defFont)}" w:hAnsi="${esc(fonts[ftcDefault[2]] || defFont)}" w:cs="${esc(fonts[ftcDefault[3]] || defFont)}"/><w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>`
    + `${stylesXml}</w:styles>`;

  // ---------- списки (PlfLst / PlfLfo) ----------
  let numberingXml = '';
  if (lcbLst) {
    // lcbLst покрывает только массив LSTF; уровни (LVL) лежат следом за ним — их граница — конец потока
    const lstEnd = T.length;
    const n = Math.min(u16(T, fcLst), Math.floor((lcbLst - 2) / 28)); const lsts = []; let q = fcLst + 2;
    for (let i = 0; i < n; i++) { lsts.push({ lsid: i32(T, q), styles: Array.from({ length: 9 }, (_, k) => u16(T, q + 8 + k * 2)), simple: !!(T[q + 26] & 1) }); q += 28; }
    // уровни: LVLF (28 байт) + grpprlPapx + grpprlChpx + xst
    for (const l of lsts) {
      l.lvls = [];
      for (let k = 0; k < (l.simple ? 1 : 9) && q + 28 <= lstEnd; k++) {
        const start = i32(T, q); const nfc = T[q + 4]; const fl = T[q + 5]; const follow = T[q + 15];
        const cbC = T[q + 24]; const cbP = T[q + 25];
        const legal = !!(fl & 0x04); const noRestart = !!(fl & 0x08);
        q += 28;
        const pg = T.subarray(q, q + cbP); q += cbP;
        const cg = T.subarray(q, q + cbC); q += cbC;
        const cch = u16(T, q); const xst = utf16(T, q + 2, cch); q += 2 + cch * 2;
        l.lvls.push({ start, nfc, jc: ['left', 'center', 'right'][fl & 3] || 'left', follow: ['tab', 'space', 'nothing'][follow] || 'tab', legal, noRestart, p: paraProps(sprms(pg)), r: runProps(sprms(cg)), xst });
      }
    }
    let abs = '';
    lsts.forEach((l, ai) => {
      abs += `<w:abstractNum w:abstractNumId="${ai}"><w:nsid w:val="${(l.lsid >>> 0).toString(16).padStart(8, '0').toUpperCase()}"/><w:multiLevelType w:val="${l.simple ? 'singleLevel' : 'multilevel'}"/>`;
      l.lvls.forEach((v, k) => {
        const txt = v.nfc === 23 ? v.xst : [...v.xst].map((ch) => (ch.charCodeAt(0) < 9 ? `%${ch.charCodeAt(0) + 1}` : ch)).join('');
        const pst = l.styles[k] !== 0x0FFF && styleIds.has(l.styles[k]) ? `<w:pStyle w:val="${styleIds.get(l.styles[k])}"/>` : '';
        const pp = { ...v.p }; delete pp.istd;
        // порядок CT_Lvl: start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlJc, pPr, rPr
        abs += `<w:lvl w:ilvl="${k}"><w:start w:val="${v.start}"/><w:numFmt w:val="${NFC[v.nfc] || 'decimal'}"/>${v.noRestart ? '<w:lvlRestart w:val="0"/>' : ''}${pst}${v.legal ? '<w:isLgl/>' : ''}<w:suff w:val="${v.follow}"/><w:lvlText w:val="${esc(txt)}"/><w:lvlJc w:val="${v.jc}"/>${pPrXml(pp)}${rPrXml(v.r)}</w:lvl>`;
      });
      abs += '</w:abstractNum>';
    });
    let nums = '';
    if (lcbLfo) {
      const m = Math.min(u32(T, fcLfo), Math.floor((lcbLfo - 4) / 16)); let o = fcLfo + 4; const lfos = [];
      for (let i = 0; i < m; i++) { lfos.push({ lsid: i32(T, o), clfolvl: T[o + 12] }); o += 16; }
      for (const f of lfos) {
        f.over = [];
        o += 4; // cp
        for (let k = 0; k < f.clfolvl; k++) {
          const startAt = i32(T, o); const b = T[o + 4]; o += 8;
          const ilvl = b & 0x0F; const fStart = !!(b & 0x10); const fFmt = !!(b & 0x20);
          if (fFmt) { const cbC = T[o + 24]; const cbP = T[o + 25]; o += 28 + cbP + cbC; const cch = u16(T, o); o += 2 + cch * 2; }
          f.over.push({ ilvl, startAt: fStart ? startAt : null });
        }
      }
      lfos.forEach((f, i) => {
        const ai = lsts.findIndex((l) => l.lsid === f.lsid);
        if (ai < 0) return;
        nums += `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${ai}"/>${f.over.filter((x) => x.startAt != null).map((x) => `<w:lvlOverride w:ilvl="${x.ilvl}"><w:startOverride w:val="${x.startAt}"/></w:lvlOverride>`).join('')}</w:num>`;
      });
    }
    numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abs}${nums}</w:numbering>`;
  }

  // ---------- разделы ----------
  const sed = plc(T, fcSed, lcbSed, 12);
  const sections = sed.data.map((d, i) => {
    const fc = u32(d, 2); const s = { end: sed.cps[i + 1], list: [] };
    if (fc !== 0xFFFFFFFF && fc < w.length) { const cb = i16(w, fc); s.list = sprms(w, fc + 2, fc + 2 + cb); }
    return s;
  });
  const sectEnds = new Set(sections.map((s) => s.end));

  // ---------- закладки, примечания, исправления ----------
  const marks = new Map(); // cp → [xml перед символом]
  const addMark = (cp, xml) => { if (!marks.has(cp)) marks.set(cp, []); marks.get(cp).push(xml); };
  {
    const names = sttb(T, fcBkmk, lcbBkmk).strs; const bkf = plc(T, fcBkf, lcbBkf, 4); const bkl = plc(T, fcBkl, lcbBkl, 0);
    bkf.data.forEach((d, i) => {
      const ibkl = u16(d, 0); const end = bkl.cps[ibkl]; const name = names[i];
      if (name == null || end == null) return;
      addMark(bkf.cps[i], `<w:bookmarkStart w:id="${i}" w:name="${esc(name)}"/>`);
      addMark(end, `<w:bookmarkEnd w:id="${i}"/>`);
    });
  }
  // примечания: ссылки (ATRDPre10), авторы, диапазоны по закладкам примечаний
  const atn = plc(T, fcAndRef, lcbAndRef, 30);
  const atnOwners = [];
  { let o = fcAtnOwn; while (o + 2 <= fcAtnOwn + lcbAtnOwn) { const cch = u16(T, o); atnOwners.push(utf16(T, o + 2, cch)); o += 2 + cch * 2; } }
  const atnRanges = new Map();
  {
    const bk = sttb(T, fcAtnBkmk, lcbAtnBkmk); const f = plc(T, fcAtnBkf, lcbAtnBkf, 4); const l = plc(T, fcAtnBkl, lcbAtnBkl, 0);
    f.data.forEach((d, i) => { const tag = bk.extra[i] && bk.extra[i].length >= 6 ? i32(bk.extra[i], 2) : null; const end = l.cps[u16(d, 0)]; if (tag != null && end != null) atnRanges.set(tag, [f.cps[i], end]); });
  }
  const comments = atn.data.map((d, i) => {
    const cchI = u16(d, 0); const initials = utf16(d, 2, Math.min(cchI, 9)); const author = atnOwners[i16(d, 20)] || '';
    const tag = i32(d, 26); const range = atnRanges.get(tag);
    if (range) { addMark(range[0], `<w:commentRangeStart w:id="${i}"/>`); addMark(range[1], `<w:commentRangeEnd w:id="${i}"/>`); }
    const date = lcbAtrdExtra >= (i + 1) * 18 ? dttm(u32(T, fcAtrdExtra + i * 18)) : null;
    return { id: i, cp: atn.cps[i], author, initials, date };
  });
  const atnByCp = new Map(comments.map((c) => [c.cp, c]));
  const fnd = plc(T, fcFndRef, lcbFndRef, 2);
  const fndByCp = new Map(fnd.cps.slice(0, -1).map((cp, i) => [cp, i + 1]));

  // ---------- рисунки ----------
  const media = []; // { name, data, ext }
  // связи — у каждой части пакета свои (рисунок в колонтитуле ссылается через связи колонтитула, не документа)
  const relsByPart = new Map(); // часть (имя в word/) → [{ id, type, target }]
  let curPart = 'document.xml';
  let relN = 0;
  const rel = (type, target) => {
    const id = `rIdDoc${++relN}`;
    if (!relsByPart.has(curPart)) relsByPart.set(curPart, []);
    relsByPart.get(curPart).push({ id, type, target });
    return id;
  };
  const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  // картинка (растровая) → часть media и связь; { rid, name, id } | null
  function addImage(blip) {
    if (!blip || !blip.data) { warn(blip ? 'рисунки EMF/WMF/PICT (не показываются)' : 'рисунки, которые Перо не прочитало'); return null; }
    const ext = blip.kind;
    const name = `doc_image${media.length + 1}.${ext}`;
    media.push({ name, data: blip.data, ext });
    return { rid: rel(`${RT}/image`, `media/${name}`), name, id: media.length };
  }
  const picXml = (img, cx, cy, xfrm = '', crop = '') => '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${img.id}" name="${img.name}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${img.rid}"/>${crop}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
    + '</a:graphicData></a:graphic>';
  // рисунок в строке: PICF + OfficeArtInlineSpContainer (SpContainer и OfficeArtFBSE с картинкой) в потоке Data
  function picture(fcPic) {
    if (fcPic + 68 > D.length) return null;
    const lcb = u32(D, fcPic); const cbHeader = u16(D, fcPic + 4); const mm = u16(D, fcPic + 6);
    const dxaGoal = i16(D, fcPic + 28); const dyaGoal = i16(D, fcPic + 30); const mx = u16(D, fcPic + 32); const my = u16(D, fcPic + 34);
    let o = fcPic + cbHeader; const end = Math.min(fcPic + lcb, D.length);
    if (mm === 0x66) o += 1 + D[o];
    const recs = artRecords(D, o, end);
    const sp = recs.find((r) => r.type === 0xF004);
    const opt = sp ? shapeOpt(D, sp) : new Map();
    let blip = null;
    // формула Word 2007+ в .doc — картинка, а сама формула (OMML внутри XML Word 2003) — в свойствах фигуры
    for (const e of opt.values()) {
      if (!e.raw || e.raw.length < 20 || e.raw[0] !== 0x3C || e.raw[1] !== 0x3F) continue;
      const xml = new TextDecoder('utf-8').decode(e.raw);
      const clean = (s) => safeOmml(s.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '').replace(/<\/?aml:[^>]*>/g, ''));
      const para = /<m:oMathPara\b[\s\S]*?<\/m:oMathPara>/.exec(xml); const one = /<m:oMath\b[\s\S]*?<\/m:oMath>/.exec(xml);
      // в XML формула всегда в oMathPara; в строке текста Word ставит один oMath
      const math = one ? clean(one[0]) : null; // не прошла проверку — показываем картинку, как без формулы
      if (math) return { math, para: para ? clean(para[0]) : null };
    }
    for (const r of recs) if (r.type === 0xF007) { blip = bseBlip(D, r, w); if (blip) break; }
    const img = addImage(blip);
    if (!img) return null;
    // размер рамки: (исходный размер PICF − поля обрезки PICF, твипы) × масштаб (‰); доли обрезки — в OPT
    const cl = i16(D, fcPic + 36); const ct = i16(D, fcPic + 38); const cr = i16(D, fcPic + 40); const cb = i16(D, fcPic + 42);
    const ex = Math.round(((dxaGoal - cl - cr) * mx / 1000) * 635); const ey = Math.round(((dyaGoal - ct - cb) * my / 1000) * 635);
    const rot = opt.has(0x0004) ? (opt.get(0x0004).val | 0) / 65536 : 0;
    const xfrm = rot ? ` rot="${Math.round(rot * 60000)}"` : '';
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${ex}" cy="${ey}"/><wp:docPr id="${img.id}" name="Рисунок ${img.id}"/>${picXml(img, ex, ey, xfrm, cropXml(opt))}</wp:inline></w:drawing>`;
  }

  // ---------- плавающие объекты (OfficeArt: DggInfo + FSPA) ----------
  const [fcDgg, lcbDgg] = fcl(50); const [fcSpaMom, lcbSpaMom] = fcl(40); const [fcSpaHdr, lcbSpaHdr] = fcl(41);
  const [fcTxbx, lcbTxbx] = fcl(56); const [fcHdrTxbx, lcbHdrTxbx] = fcl(58);
  const art = { bse: [], shapes: new Map(), background: null };
  if (lcbDgg) {
    for (const r of artRecords(T, fcDgg, fcDgg + lcbDgg)) {
      if (r.type === 0xF007) art.bse.push(r);
      if (r.type === 0xF004) {
        const fsp = r.kids.find((k) => k.type === 0xF00A);
        if (!fsp) continue;
        const spid = u32(T, fsp.at); const fl = u32(T, fsp.at + 4);
        const sh = { spid, type: fsp.inst, flags: fl, opt: shapeOpt(T, r) };
        art.shapes.set(spid, sh);
        if (fl & 0x400) art.background = sh;
      }
    }
  }
  const spaMom = plc(T, fcSpaMom, lcbSpaMom, 26); const spaHdr = plc(T, fcSpaHdr, lcbSpaHdr, 26);
  const spaAt = new Map(); spaMom.cps.slice(0, -1).forEach((cp, i) => spaAt.set(cp, spaMom.data[i]));
  const spaHdrAt = new Map(); spaHdr.cps.slice(0, -1).forEach((cp, i) => spaHdrAt.set(ccp.text + ccp.ftn + cp, spaHdr.data[i]));
  // тексты надписей: FTXBXS (22 байта), lid — номер фигуры; повторно используемые (fReusable) — пропуск
  const txbxStart = ccp.text + ccp.ftn + ccp.hdd + ccp.atn + ccp.edn;
  const txbxByShape = new Map();
  const readTxbx = (fc, lcb, base) => { const x = plc(T, fc, lcb, 22); x.data.forEach((d, i) => { if (!u16(d, 8)) txbxByShape.set(i32(d, 14), [base + x.cps[i], base + x.cps[i + 1]]); }); };
  readTxbx(fcTxbx, lcbTxbx, txbxStart); readTxbx(fcHdrTxbx, lcbHdrTxbx, txbxStart + ccp.txbx);
  let zOrder = 0;
  // msospt → предустановленная геометрия DrawingML
  const PRST = { 1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon', 10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 15: 'homePlate', 16: 'cube', 20: 'line', 21: 'plaque', 22: 'can', 23: 'donut', 32: 'straightConnector1', 55: 'chevron', 56: 'pentagon', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow', 74: 'heart', 96: 'smileyFace', 183: 'sun', 184: 'moon', 202: 'rect' };
  // плавающий объект у символа 0x08 → w:drawing (wp:anchor) | VML подложки | null
  function floating(cp, kind) {
    const fspa = (kind === 'hdr' ? spaHdrAt : spaAt).get(cp);
    if (!fspa) return null;
    const spid = i32(fspa, 0); const sh = art.shapes.get(spid);
    if (!sh) { warn('плавающие рисунки и фигуры'); return null; }
    const xa = i32(fspa, 4); const ya = i32(fspa, 8); const xb = i32(fspa, 12); const yb = i32(fspa, 16); const fl = u16(fspa, 20);
    const o = sh.opt;
    const num = (pid, def) => (o.has(pid) ? o.get(pid).val : def);
    // градусы (16.16); при одном отражении угол в DrawingML — с другим знаком (как Word переводит .doc)
    let rot = (num(0x0004, 0) | 0) / 65536;
    if (rot && (!(sh.flags & 0x40) !== !(sh.flags & 0x80))) rot = 360 - rot;
    let cx = Math.round((xb - xa) * 635); let cy = Math.round((yb - ya) * 635); let x0 = xa; let y0 = ya;
    const r360 = ((rot % 360) + 360) % 360;
    if ((r360 >= 45 && r360 < 135) || (r360 >= 225 && r360 < 315)) { // в FSPA — рамка повёрнутой фигуры: стороны наоборот вокруг центра
      const mx = (xa + xb) / 2; const my = (ya + yb) / 2; const hw = (yb - ya) / 2; const hh = (xb - xa) / 2;
      x0 = Math.round(mx - hw); y0 = Math.round(my - hh); [cx, cy] = [cy, cx];
    }
    // подложка (WordArt «PowerPlusWaterMarkObject…») — VML, как её читает «Конструктор»
    const name = o.has(0x0380) ? o.get(0x0380).str : '';
    if (sh.type === 136 && /PowerPlusWaterMarkObject/.test(name)) {
      const txt = o.has(0x00C0) ? o.get(0x00C0).str : '';
      const font = o.has(0x00C5) ? o.get(0x00C5).str : 'Calibri';
      const fc = artColor(num(0x0181, 0xC0C0C0)) || 'C0C0C0';
      const semi = num(0x0182, 65536) < 65536;
      const wm = { text: txt, font, color: `#${fc.toLowerCase()}`, semi, layout: Math.abs(r360 - 315) < 1 ? 'diagonal' : 'horizontal', size: null };
      return watermarkRunXml(wm, { w: Math.round(cx / 1270) / 10, h: Math.round(cy / 1270) / 10 }).replace(/^<w:r>(?:<w:rPr><w:noProof\/><\/w:rPr>)?/, '').replace(/<\/w:r>$/, '');
    }
    // положение: FSPA (bx/by) или свойства posh/posrelh/posv/posrelv (Word 2000+)
    const bx = (fl >> 1) & 3; const by = (fl >> 3) & 3; const wr = (fl >> 5) & 0xF; const wrk = (fl >> 9) & 0xF; const below = !!(fl & 0x4000);
    const relH = o.has(0x0390) ? ['margin', 'page', 'column', 'character'][num(0x0390, 2)] || 'column' : ['margin', 'page', 'column'][bx] || 'column';
    const relV = o.has(0x0392) ? ['margin', 'page', 'paragraph', 'line'][num(0x0392, 2)] || 'paragraph' : ['margin', 'page', 'paragraph'][by] || 'paragraph';
    const posh = num(0x038F, 0); const posv = num(0x0391, 0);
    // выравнивание из файла — только известные значения, иначе смещение (wp:align с мусором нарушает схему)
    const H = posh >= 1 && posh <= 5 ? `<wp:align>${['', 'left', 'center', 'right', 'inside', 'outside'][posh]}</wp:align>` : `<wp:posOffset>${Math.round(x0 * 635)}</wp:posOffset>`;
    const V = posv >= 1 && posv <= 5 ? `<wp:align>${['', 'top', 'center', 'bottom', 'inside', 'outside'][posv]}</wp:align>` : `<wp:posOffset>${Math.round(y0 * 635)}</wp:posOffset>`;
    const side = ['bothSides', 'left', 'right', 'largest'][wrk] || 'bothSides';
    const poly = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
    const wrap = wr === 1 ? '<wp:wrapTopAndBottom/>' : wr === 3 ? '<wp:wrapNone/>' : wr === 4 ? `<wp:wrapTight wrapText="${side}">${poly}</wp:wrapTight>` : wr === 5 ? `<wp:wrapThrough wrapText="${side}">${poly}</wp:wrapThrough>` : `<wp:wrapSquare wrapText="${side}"/>`;
    const dl = num(0x0384, 114300); const dt = num(0x0385, 0); const dr = num(0x0386, 114300); const db = num(0x0387, 0);
    const xfrm = `${rot ? ` rot="${Math.round(rot * 60000)}"` : ''}${sh.flags & 0x40 ? ' flipH="1"' : ''}${sh.flags & 0x80 ? ' flipV="1"' : ''}`;
    zOrder++;
    const head = `<w:drawing><wp:anchor distT="${dt}" distB="${db}" distL="${dl}" distR="${dr}" simplePos="0" relativeHeight="${251658240 + zOrder * 1024}" behindDoc="${below ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
      + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${relH}">${H}</wp:positionH><wp:positionV relativeFrom="${relV}">${V}</wp:positionV>`
      + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}<wp:docPr id="${1000 + zOrder}" name="${esc(name || `Фигура ${zOrder}`)}"/><wp:cNvGraphicFramePr/>`;
    const tail = '</wp:anchor></w:drawing>';
    // рисунок (рамка картинки или фигура с картинкой без своей геометрии)
    const pib = num(0x0104, 0);
    if (pib && (sh.type === 75 || !PRST[sh.type])) {
      const img = addImage(art.bse[pib - 1] ? bseBlip(T, art.bse[pib - 1], w) : null);
      if (!img) return null;
      return `${head}${picXml(img, cx, cy, xfrm, cropXml(o))}${tail}`;
    }
    // фигура или надпись: wps:wsp (геометрия, заливка, линия, текст)
    const prst = PRST[sh.type];
    if (!prst) { warn('фигуры, которые Перо не прочитало'); return null; }
    // логические свойства: биты 0–15 — значения, 16–31 — «задано» (fUse…)
    const bool = (pid, bit, def) => { if (!o.has(pid)) return def; const v = o.get(pid).val; return v & (1 << (bit + 16)) ? !!(v & (1 << bit)) : def; };
    const filled = bool(0x01BF, 4, sh.type !== 20 && sh.type !== 32); const lined = bool(0x01FF, 3, true);
    const fill = filled ? `<a:solidFill><a:srgbClr val="${artColor(num(0x0181, 0xFFFFFF)) || 'FFFFFF'}"/></a:solidFill>` : '<a:noFill/>';
    const ARROW = [null, 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];
    const ends = `${ARROW[num(0x01D0, 0)] ? `<a:headEnd type="${ARROW[num(0x01D0, 0)]}"/>` : ''}${ARROW[num(0x01D1, 0)] ? `<a:tailEnd type="${ARROW[num(0x01D1, 0)]}"/>` : ''}`;
    const ln = lined ? `<a:ln w="${num(0x01CB, 9525)}"><a:solidFill><a:srgbClr val="${artColor(num(0x01C0, 0)) || '000000'}"/></a:solidFill>${ends}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
    const tx = txbxByShape.get(spid);
    let txt = '';
    if (tx) {
      const content = story(tx[0], tx[1] - (text.charCodeAt(tx[1] - 1) === 0x0D && tx[1] - tx[0] > 1 && text.charCodeAt(tx[1] - 2) === 0x0D ? 1 : 0), 'txbx');
      txt = `<wps:txbx><w:txbxContent>${content || '<w:p/>'}</w:txbxContent></wps:txbx>`;
    }
    const anchorT = ['t', 'ctr', 'b', 't', 'ctr', 'b'][num(0x0087, 0)] || 't';
    const body = `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${num(0x0081, 91440)}" tIns="${num(0x0082, 45720)}" rIns="${num(0x0083, 91440)}" bIns="${num(0x0084, 45720)}" anchor="${anchorT}" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
    return `${head}<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`
      + `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr${tx ? ' txBox="1"' : ''}/><wps:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr>${txt}${body}</wps:wsp></a:graphicData></a:graphic>${tail}`;
  }

  // ---------- текст истории → абзацы и таблицы ----------
  // runsOf: диапазон CP → куски с одинаковыми свойствами символов
  function runsOf(a, b) {
    const out = [];
    let cp = a;
    while (cp < b) {
      const pc = pieceAt(cp); const lim = Math.min(b, pc.cp1);
      const fc = fcOf(cp); const run = find(chpx, fc);
      const per = pc.compressed ? 1 : 2;
      const runEnd = run ? Math.min(lim, cp + Math.ceil((run.fc1 - fc) / per)) : lim;
      const extra = prmSprms(pc.prm);
      out.push({ a: cp, b: Math.max(runEnd, cp + 1), s: extra.length ? [...(run ? run.s : []), ...extra] : run ? run.s : [] });
      cp = Math.max(runEnd, cp + 1);
    }
    return out;
  }
  const papAt = (cp) => { const pc = pieceAt(cp); const r = find(papx, fcOf(cp)); const extra = prmSprms(pc.prm); return r ? { istd: r.istd, s: extra.length ? [...r.s, ...extra] : r.s } : { istd: 0, s: extra }; };

  const fieldStack = [];
  // одна история (основной текст, сноска, колонтитул…) → OOXML блоков; story: 'main' | 'ftn' | 'hdr' | 'atn' | 'txbx';
  // part — часть пакета, куда пойдёт текст (связи рисунков и ссылок — в её .rels); у надписи — часть, где стоит фигура
  // диапазон каждой истории — границы для CP из PLC файла; каждая побочная история выводится не больше одного раза
  const atnBase = ccp.text + ccp.ftn + ccp.hdd;
  const RANGE = {
    main: [0, ccp.text], ftn: [ccp.text, ccp.text + ccp.ftn], hdr: [ccp.text + ccp.ftn, atnBase], atn: [atnBase, atnBase + ccp.atn],
    txbx: [atnBase + ccp.atn + ccp.edn, atnBase + ccp.atn + ccp.edn + ccp.txbx + ccp.hdrTxbx],
  };
  const storyDone = new Set();
  function story(a0, b0, kind, sectBreaks = null, part = null) {
    const lim = RANGE[kind] || RANGE.main;
    const a = Math.max(a0, lim[0]); const b = Math.min(b0, lim[1], text.length);
    if (!(b > a)) return '';
    if (kind !== 'main') { const key = `${kind}:${a}:${b}`; if (storyDone.has(key)) return ''; storyDone.add(key); }
    const prevPart = curPart; const prevStack = fieldStack.splice(0); // незакрытое поле не «протекает» между историями
    if (part) curPart = part;
    try { return storyBody(a, b, kind, sectBreaks); } finally { curPart = prevPart; fieldStack.length = 0; fieldStack.push(...prevStack); }
  }
  function storyBody(a, b, kind, sectBreaks) {
    let xml = '';
    let cpStart = a;
    let table = null; // { rows: [{ cells: [xml], tap }], cell: xml }
    let cellXml = '';
    const flushTable = () => { if (table) { xml += tableXml(table); table = null; } };
    for (let cp = a; cp < b; cp++) {
      const ch = text.charCodeAt(cp);
      const isSect = ch === 0x0C && sectBreaks && sectBreaks.has(cp + 1);
      if (ch !== 0x0D && ch !== 0x07 && !isSect) continue;
      // абзац [cpStart, cp] — его свойства по метке конца
      const pap = papAt(cp);
      const pp = paraProps(pap.s); pp.istd = pp.istd ?? pap.istd;
      const inTable = pp.inTable && kind !== 'nested';
      const itap = pp.itap || (inTable ? 1 : 0);
      if (itap > 1) { warn('вложенные таблицы (показаны текстом)'); }
      if (ch === 0x07 && pp.ttp) { // конец строки таблицы: TAPX этой метки
        if (!table) table = { rows: [] };
        // закладки и диапазоны примечаний на метке строки — в конец последней ячейки
        const pend = table.pending || [];
        if (marks.has(cp) && pend.length) { const last = pend.length - 1; const i = pend[last].lastIndexOf('</w:p>'); if (i >= 0) pend[last] = `${pend[last].slice(0, i)}${marks.get(cp).join('')}${pend[last].slice(i)}`; }
        table.rows.push({ cells: pend, tap: pap.s });
        table.pending = [];
        cpStart = cp + 1;
        continue;
      }
      let sectPr = null;
      if (isSect) { const si = sections.findIndex((s) => s.end === cp + 1); sectPr = sectPrXml(si); }
      const para = paragraphXml(cpStart, cp, pp, kind, sectPr);
      if (inTable && itap <= 1) {
        if (!table) table = { rows: [], pending: [] };
        cellXml += para;
        if (ch === 0x07) { table.pending.push(cellXml); cellXml = ''; }
      } else if (inTable) {
        cellXml += para; // вложенная таблица — абзацами в ячейке
        if (ch === 0x07 && pp.innerTtp) cellXml = cellXml.replace(/<w:p>(?:(?!<w:p>).)*$/, '');
      } else {
        flushTable();
        xml += para;
      }
      cpStart = cp + 1;
    }
    if (cellXml && table) { table.pending.push(cellXml); }
    flushTable();
    return xml;
  }

  // абзац [a, endCp): текст до метки конца; метка — свойства абзаца
  function paragraphXml(a, endCp, pp, kind, sectPr) {
    let body = '';
    const markRuns = runsOf(endCp, endCp + 1);
    const markR = runProps(markRuns[0] ? markRuns[0].s : [], styleToggle(pp.istd || 0));
    let skipTo = -1; // поле, заменённое целиком (формула в QUOTE), — пропустить до его конца
    for (const run of a < endCp ? runsOf(a, endCp) : []) {
      if (run.b <= skipTo) { for (let k = run.a; k < run.b; k++) if (marks.has(k)) body += marks.get(k).join(''); continue; } // метки внутри пропущенного поля — не терять
      const r = runProps(run.s, styleToggle(pp.istd || 0));
      let buf = ''; let instr = '';
      const rpr = rPrXml(r);
      const flush = () => { if (buf) { body += wrapRev(r, `<w:r>${rpr}<w:${r.del ? 'delText' : 't'} xml:space="preserve">${esc(buf)}</w:${r.del ? 'delText' : 't'}></w:r>`); buf = ''; } };
      const emit = (inner) => { flush(); body += wrapRev(r, `<w:r>${rpr}${inner}</w:r>`); };
      for (let cp = run.a; cp < run.b; cp++) {
        if (marks.has(cp)) { flush(); body += marks.get(cp).join(''); }
        const c = text.charCodeAt(cp);
        const inInstr = fieldStack.length && fieldStack[fieldStack.length - 1] === 'instr';
        if (cp < skipTo) continue;
        if (c === 0x13) { // формула в строке: Word 2007+ пишет её в .doc полем QUOTE с картинкой-формулой
          const q = quoteMath(cp, endCp);
          if (q) { flush(); body += q.math; skipTo = q.end + 1; continue; }
        }
        if (c === 0x13) { flush(); fieldStack.push('instr'); body += `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r>`; continue; }
        if (c === 0x14 && fieldStack.length) { flush(); fieldStack[fieldStack.length - 1] = 'result'; body += `<w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r>`; continue; }
        if (c === 0x15 && fieldStack.length) { flush(); fieldStack.pop(); body += `<w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>`; continue; }
        if (inInstr) { // код поля — одним прогоном до разделителя
          if (c === 0x01 || c === 0x08) continue;
          flush(); instr += text[cp];
          if (cp + 1 >= run.b || [0x13, 0x14, 0x15].includes(text.charCodeAt(cp + 1)) || marks.has(cp + 1)) { body += `<w:r>${rpr}<w:instrText xml:space="preserve">${esc(instr)}</w:instrText></w:r>`; instr = ''; }
          continue;
        }
        if (r.sym && (c === 0x28 || r.spec)) { emit(`<w:sym w:font="${esc(r.sym.font || 'Symbol')}" w:char="${r.sym.ch.toString(16).toUpperCase().padStart(4, '0')}"/>`); continue; }
        if (r.spec) {
          if (c === 0x01 && r.pic != null && !r.ole) {
            const x = picture(r.pic);
            // формула одна в абзаце — отдельной строкой (oMathPara), иначе — в строке
            if (x && x.math) { flush(); body += x.para && /^[\x01\s]*$/.test(text.slice(a, endCp)) ? x.para : x.math; } else if (x) emit(x);
            continue;
          }
          if (c === 0x01) { warn('объекты OLE'); continue; }
          if (c === 0x02) { // знак сноски
            if (kind === 'main' && fndByCp.has(cp)) emit(`<w:footnoteReference w:id="${fndByCp.get(cp)}"/>`);
            else if (kind === 'ftn') emit('<w:footnoteRef/>');
            continue;
          }
          if (c === 0x05) { // знак примечания
            if (kind === 'main' && atnByCp.has(cp)) emit(`<w:commentReference w:id="${atnByCp.get(cp).id}"/>`);
            else if (kind === 'atn') emit('<w:annotationRef/>');
            continue;
          }
          if (c === 0x08) { if (kind !== 'txbx') { const x = floating(cp, kind); if (x) emit(x); } continue; } // в надписи — без фигур (иначе петля)
          if (c === 0x28 || c === 0x29 || c < 0x20) continue;
        }
        if (c === 0x09) { emit('<w:tab/>'); continue; }
        if (c === 0x0B) { emit('<w:br/>'); continue; }
        if (c === 0x0C) { emit('<w:br w:type="page"/>'); continue; }
        if (c === 0x0E) { emit('<w:br w:type="column"/>'); continue; }
        if (c === 0x1E) { emit('<w:noBreakHyphen/>'); continue; }
        if (c === 0x1F) { emit('<w:softHyphen/>'); continue; }
        if (c < 0x20) continue;
        buf += text[cp];
      }
      flush();
    }
    if (marks.has(endCp)) body += marks.get(endCp).join('');
    // свойства знака абзаца (размер и т. п.) — как Word пишет в pPr/rPr
    const markRpr = rPrXml(markR);
    // вставленный / удалённый знак абзаца (CT_ParaRPr: ins, del — до rStyle и прочего)
    const insM = markR.ins ? `<w:ins w:id="${revN++}" w:author="${esc(authors[markR.ibst || 0] || '')}" w:date="${dttm(markR.dttm)}"/>` : '';
    const delM = markR.del ? `<w:del w:id="${revN++}" w:author="${esc(authors[markR.ibstDel ?? markR.ibst ?? 0] || '')}" w:date="${dttm(markR.dttmDel || markR.dttm)}"/>` : '';
    const mark = insM || delM ? `<w:rPr>${insM}${delM}${markRpr.slice(7, -8)}</w:rPr>` : markRpr;
    return `<w:p>${pPrXml(pp, { sectPr, rPr: mark })}${hyperlinks(body)}</w:p>`;
  }
  let revN = 1;
  // поле QUOTE, в коде которого — картинка-формула: { math, end } (end — CP метки конца поля)
  // разделитель и конец каждого поля — одним проходом по всему тексту (стек), а не поиском от каждого 0x13
  let fieldPairs = null;
  function fieldOf(cp) {
    if (!fieldPairs) {
      fieldPairs = new Map(); const st = [];
      for (let k = 0; k < text.length; k++) {
        const c = text.charCodeAt(k);
        if (c === 0x13) st.push({ start: k, sep: -1 });
        else if (c === 0x14 && st.length && st[st.length - 1].sep < 0) st[st.length - 1].sep = k;
        else if (c === 0x15 && st.length) { const f = st.pop(); fieldPairs.set(f.start, { sep: f.sep, end: k }); }
      }
    }
    return fieldPairs.get(cp) || null;
  }
  function quoteMath(cp, lim) {
    const f = fieldOf(cp);
    if (!f || f.end >= lim) return null;
    const { sep, end } = f;
    const instr = text.slice(cp + 1, sep > 0 ? sep : end);
    if (!/^\s*QUOTE\b/.test(instr)) return null;
    const at = cp + 1 + instr.indexOf('\u0001');
    if (at <= cp) return null;
    const r = runProps(runsOf(at, at + 1)[0].s);
    const x = r.pic != null ? picture(r.pic) : null;
    return x && x.math ? { math: x.math, end } : null;
  }
  // поле HYPERLINK с простым результатом — элемент w:hyperlink (так Word переводит .doc в .docx)
  function hyperlinks(body) {
    const R = (inner) => `<w:r>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?${inner}</w:r>`;
    const re = new RegExp(`${R('<w:fldChar w:fldCharType="begin"/>')}${R('<w:instrText xml:space="preserve">\\s*HYPERLINK\\s+([^<]*)</w:instrText>')}${R('<w:fldChar w:fldCharType="separate"/>')}((?:(?!<w:fldChar)(?!<w:instrText).)*?)${R('<w:fldChar w:fldCharType="end"/>')}`, 'g');
    return body.replace(re, (m, instr, result) => {
      const a = unesc(instr); // код поля уже экранирован в instrText — вернуть как было (иначе &amp;lt; при повторном esc)
      const anchor = /\\l\s+"([^"]*)"/.exec(a); // адрес — в кавычках или первым словом без кавычек (не ключ вида \l, \o)
      let url = /^\s*"([^"]*)"/.exec(a) || /^\s*(?!\\[a-z](\s|$))(\S+)/i.exec(a);
      if (url && url.length > 2 && url[2] != null) url = [url[0], url[2]];
      const tip = /\\o\s+"([^"]*)"/.exec(a);
      // ссылка на файл, сетевую папку и т. п. — остаётся полем, как в исходном .doc (не теряется при сохранении; в Пере
      // поле по щелчку не открывается, Word при переходе по такой ссылке сам предупреждает)
      if (url && !safeUrl(url[1])) return m;
      if (!url && !anchor) return m;
      const attrs = `${url ? ` r:id="${rel(`${RT}/hyperlink`, url[1])}"` : ''}${anchor ? ` w:anchor="${esc(anchor[1])}"` : ''}${tip ? ` w:tooltip="${esc(tip[1])}"` : ''} w:history="1"`;
      return `<w:hyperlink${attrs}>${result}</w:hyperlink>`;
    });
  }
  // исправления: вставка, удаление или вставка, удалённая другим рецензентом (w:del внутри w:ins)
  function wrapRev(r, x) {
    let out = x;
    if (r.del) out = `<w:del w:id="${revN++}" w:author="${esc(authors[r.ibstDel ?? r.ibst ?? 0] || '')}" w:date="${dttm(r.dttmDel || r.dttm)}">${out}</w:del>`;
    if (r.ins) out = `<w:ins w:id="${revN++}" w:author="${esc(authors[r.ibst || 0] || '')}" w:date="${dttm(r.dttm)}">${out}</w:ins>`;
    return out;
  }

  // таблица: строки с TAPX (sprmTDefTable — границы столбцов и TC80 ячеек)
  function tableXml(t) {
    const rows = t.rows.map((row) => {
      const tap = { centers: [], tcs: [], shd: [], brc: null, jc: null, height: null, header: false, cantSplit: false, left: null, vAlign: [], setBrc: [], width: null, style: null, colors: [[], [], [], []], shdNew: [] };
      for (const s of row.tap) {
        const g = s.g; const o = s.at;
        switch (s.sprm) {
          case 0xD608: {
            const n = g[o]; for (let k = 0; k <= n; k++) tap.centers.push(i16(g, o + 1 + k * 2));
            const tcAt = o + 1 + (n + 1) * 2; const left = s.at + s.len - tcAt;
            for (let k = 0; k < n; k++) {
              if ((k + 1) * 20 > left) { tap.tcs.push({ grf: 0, brc: [] }); continue; }
              const b = tcAt + k * 20;
              tap.tcs.push({ grf: u16(g, b), brc: [brc80(g, b + 4), brc80(g, b + 8), brc80(g, b + 12), brc80(g, b + 16)] });
            }
            break;
          }
          case 0xD613: tap.brc = [0, 1, 2, 3, 4, 5].map((k) => brc8(s, o + k * 8)); break;
          case 0xD605: tap.brc ||= [0, 1, 2, 3, 4, 5].map((k) => { const b = brc80(g, o + k * 4); return b || { val: 'nil' }; }); break;
          case 0xD612: for (let k = 0; k * 10 < s.len; k++) tap.shd[k] = shd10(g, o + k * 10); break;
          case 0xD616: for (let k = 0; k * 10 < s.len; k++) tap.shd[22 + k] = shd10(g, o + k * 10); break;
          case 0xD60C: for (let k = 0; k * 10 < s.len; k++) tap.shd[44 + k] = shd10(g, o + k * 10); break;
          case 0xD609: for (let k = 0; k * 2 < s.len; k++) tap.shd[k] ||= shd80(u16(g, o + k * 2)); break;
          case 0x5400: case 0x548A: tap.jc = ['left', 'center', 'right'][u16(g, o)] || null; break;
          case 0x9407: tap.height = i16(g, o); break;
          case 0x3404: tap.header = !!g[o]; break;
          case 0x3466: tap.cantSplit = !!g[o]; break; // sprmTFCantSplit90; старый 0x3403 Word не учитывает (проверено)
          case 0x9601: tap.left = i16(g, o); break;
          case 0xD62C: tap.vAlign.push({ a: g[o], b: g[o + 1], v: g[o + 2] }); break;
          case 0xD62F: tap.setBrc.push({ a: g[o], b: g[o + 1], sides: g[o + 2], brc: brc8(s, o + 3) }); break;
          case 0xF614: tap.width = { fts: g[o], w: i16(g, o + 1) }; break;
          // точные цвета сторон ячеек (COLORREF по ячейкам): верх, лево, низ, право
          case 0xD61A: case 0xD61B: case 0xD61C: case 0xD61D: for (let k = 0; k * 4 < s.len; k++) tap.colors[s.sprm - 0xD61A][k] = colorref(g, o + k * 4); break;
          case 0xD670: for (let k = 0; k * 10 < s.len; k++) tap.shdNew[k] = shd10(g, o + k * 10); break;
          case 0x563A: tap.style = u16(g, o); break;
          default: break;
        }
      }
      return { ...row, tap };
    });
    // сетка — все границы столбцов всех строк
    const xs = [...new Set(rows.flatMap((r) => r.tap.centers))].sort((a, b) => a - b);
    const first = rows[0] ? rows[0].tap : { brc: null };
    const bn = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
    const tblBorders = first.brc ? `<w:tblBorders>${first.brc.map((b, k) => brcXml(bn[k], b)).join('')}</w:tblBorders>` : '';
    const style = first.style != null && styleIds.has(first.style) ? `<w:tblStyle w:val="${styleIds.get(first.style)}"/>` : '';
    let xml = `<w:tbl><w:tblPr>${style}<w:tblW w:w="${first.width && first.width.fts === 3 ? first.width.w : 0}" w:type="${first.width && first.width.fts === 3 ? 'dxa' : 'auto'}"/>${first.jc && first.jc !== 'left' ? `<w:jc w:val="${first.jc}"/>` : ''}${first.left ? `<w:tblInd w:w="${first.left}" w:type="dxa"/>` : ''}${tblBorders}<w:tblLayout w:type="fixed"/></w:tblPr>`;
    xml += `<w:tblGrid>${xs.slice(0, -1).map((x, k) => `<w:gridCol w:w="${xs[k + 1] - x}"/>`).join('')}</w:tblGrid>`;
    for (const row of rows) {
      const tp = row.tap;
      const trPr = [tp.cantSplit ? '<w:cantSplit/>' : '', tp.height ? `<w:trHeight w:val="${Math.abs(tp.height)}"${tp.height < 0 ? ' w:hRule="exact"' : ''}/>` : '', tp.header ? '<w:tblHeader/>' : ''].join('');
      xml += `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}`;
      // ячеек больше, чем описано в TDefTable (повреждённая или урезанная строка), — без свойств, но с содержимым
      // ячеек больше, чем описано в TDefTable: содержимое лишних — в последнюю описанную (у лишних нет столбца сетки)
      if (row.cells.length > tp.tcs.length) {
        if (!tp.tcs.length) tp.tcs.push({ grf: 0, brc: [], extra: true });
        const keep = tp.tcs.length;
        row.cells = [...row.cells.slice(0, keep - 1), row.cells.slice(keep - 1).join('')];
      }
      for (let k = 0; k < tp.tcs.length; k++) {
        const tc = tp.tcs[k]; const grf = tc.grf;
        if (tc.extra) { xml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${row.cells[k] || '<w:p/>'}</w:tc>`; continue; } // строка без TDefTable
        if (grf & 0x02) continue; // объединена с предыдущей (старый способ)
        let span = 1; while (k + span < tp.tcs.length && (tp.tcs[k + span].grf & 0x02)) span++;
        const x0 = tp.centers[k]; const x1 = tp.centers[k + span];
        const gs = xs.indexOf(x1) - xs.indexOf(x0);
        const vm = grf & 0x20 ? (grf & 0x40 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '';
        let va = (grf >> 7) & 3; for (const v of tp.vAlign) if (k >= v.a && k < v.b) va = v.v;
        const brc = tc.brc.map((b, q) => (b && b.val !== 'nil' && tp.colors[q][k] ? { ...b, color: tp.colors[q][k] } : b));
        for (const sb of tp.setBrc) if (k >= sb.a && k < sb.b) for (let q = 0; q < 4; q++) if (sb.sides & (1 << q)) brc[q] = sb.brc;
        const bxml = [brcXml('top', brc[0]), brcXml('left', brc[1]), brcXml('bottom', brc[2]), brcXml('right', brc[3])].join('');
        const tcPr = `<w:tcPr><w:tcW w:w="${x1 - x0}" w:type="dxa"/>${gs > 1 ? `<w:gridSpan w:val="${gs}"/>` : ''}${vm}${bxml ? `<w:tcBorders>${bxml}</w:tcBorders>` : ''}${tp.shdNew[k] || tp.shd[k] || ''}${va ? `<w:vAlign w:val="${['top', 'center', 'bottom'][va]}"/>` : ''}</w:tcPr>`;
        let content = row.cells[k] || '<w:p/>';
        for (let q = 1; q < span; q++) if (row.cells[k + q] && /<w:t\b/.test(row.cells[k + q])) content += row.cells[k + q];
        xml += `<w:tc>${tcPr}${content}</w:tc>`;
      }
      xml += '</w:tr>';
    }
    return `${xml}</w:tbl>`;
  }

  // ---------- колонтитулы ----------
  const hdd = plc(T, fcHdd, lcbHdd, 0).cps;
  const hddBase = ccp.text + ccp.ftn;
  const hfParts = []; // { name, xml, type: header|footer }
  const hfStory = (k) => { if (hdd.length < k + 2) return null; const a = hdd[k]; const b = hdd[k + 1]; if (b <= a) return null; return [hddBase + a, hddBase + b]; };
  function hfRef(si, k) {
    const st = hfStory(6 + si * 6 + k);
    if (!st) return null;
    const isHeader = k === 0 || k === 1 || k === 4;
    const name = `${isHeader ? 'header' : 'footer'}_doc${hfParts.length + 1}.xml`;
    let body = story(st[0], st[1] - (text.charCodeAt(st[1] - 1) === 0x0D && st[1] - st[0] > 1 && text.charCodeAt(st[1] - 2) === 0x0D ? 1 : 0), 'hdr', null, name);
    if (!body) body = '<w:p/>';
    hfParts.push({ name, type: isHeader ? 'header' : 'footer', xml: body });
    return { id: rel(`${RT}/${isHeader ? 'header' : 'footer'}`, name), isHeader, type: k === 0 || k === 2 ? 'even' : k === 4 || k === 5 ? 'first' : 'default' };
  }

  // колонки: равные — число и промежуток; разные — w:col каждой (ширина, промежуток после)
  function colsXml(v) {
    const head = `<w:cols${v.cols > 1 ? ` w:num="${v.cols}"` : ''}${v.sep ? ' w:sep="1"' : ''} w:space="${v.colSpace}"`;
    if (v.cols < 2 || v.even !== false || !v.colW) return `${head}/>`;
    let x = `${head} w:equalWidth="0">`;
    for (let k = 0; k < v.cols; k++) x += `<w:col w:w="${v.colW[k] || 0}" w:space="${(v.colSp && v.colSp[k]) ?? v.colSpace}"/>`;
    return `${x}</w:cols>`;
  }
  // sectPr раздела
  function sectPrXml(si) {
    const sec = sections[si];
    const v = { w: 12240, h: 15840, l: 1800, r: 1800, t: 1440, b: 1440, hdr: 720, ftr: 720, gutter: 0, orient: null, bkc: 2, title: false, cols: 1, colSpace: 720, sep: false, lnn: null, pgnStart: null, pgnRestart: false, vjc: 0 };
    for (const s of sec.list) {
      switch (s.sprm) {
        case 0xB01F: v.w = opU16(s); break;
        case 0xB020: v.h = opU16(s); break;
        case 0xB021: v.l = opU16(s); break;
        case 0xB022: v.r = opU16(s); break;
        case 0x9023: v.t = opI16(s); break;
        case 0x9024: v.b = opI16(s); break;
        case 0xB025: v.gutter = opU16(s); break;
        case 0xB017: v.hdr = opU16(s); break;
        case 0xB018: v.ftr = opU16(s); break;
        case 0x301D: v.orient = opU8(s) === 2 ? 'landscape' : null; break;
        case 0x3009: v.bkc = opU8(s); break;
        case 0x300A: v.title = !!opU8(s); break;
        case 0x500B: v.cols = opU16(s) + 1; break;
        case 0x900C: v.colSpace = opU16(s); break;
        case 0x3019: v.sep = !!opU8(s); break;
        case 0x5015: v.lnn = { ...(v.lnn || {}), countBy: opU16(s) }; break;
        case 0x9016: v.lnn = { ...(v.lnn || {}), distance: opI16(s) }; break;
        case 0x3013: v.lnn = { ...(v.lnn || {}), restart: ['newPage', 'newSection', 'continuous'][opU8(s)] }; break;
        case 0x501B: v.lnn = { ...(v.lnn || {}), start: opU16(s) }; break;
        case 0x501C: case 0x7044: v.pgnStart = s.len === 4 ? u32(s.g, s.at) : opU16(s); break;
        case 0x3011: v.pgnRestart = !!opU8(s); break;
        case 0x301A: v.vjc = opU8(s); break;
        case 0x9031: v.pitch = opI16(s); break;
        case 0x3005: v.even = !!opU8(s); break;
        case 0xF203: (v.colW ||= [])[s.g[s.at]] = u16(s.g, s.at + 1); break;
        case 0xF204: (v.colSp ||= [])[s.g[s.at]] = u16(s.g, s.at + 1); break;
        case 0x5032: v.clm = opU16(s); break;
        case 0xD234: case 0xD235: case 0xD236: case 0xD237: (v.pgb ||= {})[['top', 'left', 'bottom', 'right'][s.sprm - 0xD234]] = brc8(s); break;
        case 0x522F: v.pgbProp = opU16(s); break;
        default: break;
      }
    }
    const refs = [];
    for (const k of [0, 1, 2, 3, 4, 5]) { const x = hfRef(si, k); if (x) refs.push(x); } // порядок, как пишет Word
    const BKC = ['continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'];
    return `<w:sectPr>${refs.map((x) => `<w:${x.isHeader ? 'header' : 'footer'}Reference w:type="${x.type}" r:id="${x.id}"/>`).join('')}`
      + `${v.bkc !== 2 ? `<w:type w:val="${BKC[v.bkc] || 'nextPage'}"/>` : ''}<w:pgSz w:w="${v.w}" w:h="${v.h}"${v.orient ? ` w:orient="${v.orient}"` : ''}/>`
      + `<w:pgMar w:top="${v.t}" w:right="${v.r}" w:bottom="${v.b}" w:left="${v.l}" w:header="${v.hdr}" w:footer="${v.ftr}" w:gutter="${v.gutter}"/>`
      + `${v.pgb ? `<w:pgBorders${(v.pgbProp >> 5) & 7 ? ' w:offsetFrom="page"' : ''}${(v.pgbProp >> 3) & 3 ? ' w:zOrder="back"' : ''}${v.pgbProp & 7 ? ` w:display="${['allPages', 'firstPage', 'notFirstPage'][v.pgbProp & 7]}"` : ''}>${['top', 'left', 'bottom', 'right'].map((k) => brcXml(k, v.pgb[k])).join('')}</w:pgBorders>` : ''}`
      + `${v.lnn && v.lnn.countBy ? `<w:lnNumType w:countBy="${v.lnn.countBy}"${v.lnn.start ? ` w:start="${v.lnn.start}"` : ''}${v.lnn.distance ? ` w:distance="${v.lnn.distance}"` : ''}${v.lnn.restart && v.lnn.restart !== 'newPage' ? ` w:restart="${v.lnn.restart}"` : ''}/>` : ''}`
      + `${v.pgnRestart && v.pgnStart != null ? `<w:pgNumType w:start="${v.pgnStart}"/>` : ''}${colsXml(v)}`
      + `${v.vjc ? `<w:vAlign w:val="${['top', 'center', 'both', 'bottom'][v.vjc]}"/>` : ''}${v.title ? '<w:titlePg/>' : ''}`
      + `<w:docGrid${v.clm ? ` w:type="${['default', 'linesAndChars', 'lines', 'snapToChars'][v.clm] || 'lines'}"` : ''} w:linePitch="${v.pitch ?? 360}"/></w:sectPr>`;
  }

  // ---------- сборка ----------
  const bodyXml = story(0, ccp.text, 'main', sectEnds);
  const lastSect = sectPrXml(sections.length - 1);
  // сноски
  let footnotesXml = '';
  if (lcbFndTxt && fnd.cps.length > 1) {
    const ft = plc(T, fcFndTxt, lcbFndTxt, 0).cps;
    for (let i = 0; i + 1 < ft.length && i < fnd.cps.length - 1; i++) {
      const a = ccp.text + ft[i]; let b = ccp.text + ft[i + 1];
      if (b > a && text.charCodeAt(b - 1) === 0x0D && b - a > 1 && i + 2 === ft.length - 0) b--;
      footnotesXml += `<w:footnote w:id="${i + 1}">${story(a, b, 'ftn', null, 'footnotes.xml') || '<w:p/>'}</w:footnote>`;
    }
  }
  // примечания
  let commentsXml = '';
  if (comments.length) {
    const at = plc(T, fcAndTxt, lcbAndTxt, 0).cps; const base = ccp.text + ccp.ftn + ccp.hdd;
    comments.forEach((c, i) => {
      const a = base + at[i]; const b = base + at[i + 1];
      commentsXml += `<w:comment w:id="${c.id}" w:author="${esc(c.author)}" w:date="${c.date || '2000-01-01T00:00:00Z'}" w:initials="${esc(c.initials)}">${b > a ? story(a, b, 'atn', null, 'comments.xml') : '<w:p/>'}</w:comment>`;
    });
  }
  // параметры документа (DopBase): байт 0 бит 0 — разные колонтитулы чётных/нечётных; байт 5: бит 3 — переносить
  // прописные, бит 4 — автоперенос; байт 6 бит 5 — зеркальные поля; байты 10–11 — шаг табуляции
  const dop = (k) => (lcbDop > k ? T[fcDop + k] : 0);
  const dopSettings = (bgShape) => [bgShape ? '<w:displayBackgroundShape/>' : '', dop(6) & 0x20 ? '<w:mirrorMargins/>' : '',
    `<w:defaultTabStop w:val="${(lcbDop >= 12 && u16(T, fcDop + 10)) || 720}"/>`, dop(5) & 0x10 ? '<w:autoHyphenation/>' : '',
    dop(5) & 0x10 && !(dop(5) & 0x08) ? '<w:doNotHyphenateCaps/>' : '', dop(0) & 0x01 ? '<w:evenAndOddHeaders/>' : ''].join('');
  const NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const zip = new JSZip();
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  // цвет страницы — фоновая фигура (fBackground) основного рисунка
  const bg = art.background && art.background.opt.has(0x0181) ? artColor(art.background.opt.get(0x0181).val) : null;
  zip.file('word/document.xml', `${head}<w:document ${NS}>${bg ? `<w:background w:color="${bg}"/>` : ''}<w:body>${bodyXml}${lastSect}</w:body></w:document>`);
  zip.file('word/styles.xml', stylesPart);
  const docRels = [{ id: 'rIdStyles', type: `${RT}/styles`, target: 'styles.xml' }, { id: 'rIdSettings', type: `${RT}/settings`, target: 'settings.xml' }];
  zip.file('word/settings.xml', `${head}<w:settings ${NS}>${dopSettings(!!bg)}</w:settings>`);
  const overrides = [['/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'], ['/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'], ['/word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml']];
  if (numberingXml) { zip.file('word/numbering.xml', numberingXml); docRels.push({ id: 'rIdNumbering', type: `${RT}/numbering`, target: 'numbering.xml' }); overrides.push(['/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml']); }
  if (footnotesXml) {
    zip.file('word/footnotes.xml', `${head}<w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${footnotesXml}</w:footnotes>`);
    docRels.push({ id: 'rIdFootnotes', type: `${RT}/footnotes`, target: 'footnotes.xml' }); overrides.push(['/word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml']);
  }
  if (commentsXml) {
    zip.file('word/comments.xml', `${head}<w:comments ${NS}>${commentsXml}</w:comments>`);
    docRels.push({ id: 'rIdComments', type: `${RT}/comments`, target: 'comments.xml' }); overrides.push(['/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml']);
  }
  for (const h of hfParts) {
    zip.file(`word/${h.name}`, `${head}<w:${h.type === 'header' ? 'hdr' : 'ftr'} ${NS}>${h.xml}</w:${h.type === 'header' ? 'hdr' : 'ftr'}>`);
    overrides.push([`/word/${h.name}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${h.type}+xml`]);
  }
  for (const m of media) zip.file(`word/media/${m.name}`, m.data);
  // свойства документа: SummaryInformation / DocumentSummaryInformation → docProps/core.xml
  const props = { ...propertySet(cfb.stream('\u0005SummaryInformation') || cfb.stream('SummaryInformation')) };
  const pkgRels = [`<Relationship Id="rId1" Type="${RT}/officeDocument" Target="word/document.xml"/>`];
  if (Object.keys(props).length) {
    const tag = (n, pid) => (props[pid] ? `<${n}>${esc(props[pid])}</${n}>` : '');
    const dt = (n, pid) => (props[pid] instanceof Date ? `<${n} xsi:type="dcterms:W3CDTF">${props[pid].toISOString().replace(/\.\d+Z$/, 'Z')}</${n}>` : '');
    zip.file('docProps/core.xml', `${head}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
      + `${tag('dc:title', 2)}${tag('dc:subject', 3)}${tag('dc:creator', 4)}${tag('cp:keywords', 5)}${tag('dc:description', 6)}${tag('cp:lastModifiedBy', 8)}${tag('cp:revision', 9)}${dt('dcterms:created', 12)}${dt('dcterms:modified', 13)}</cp:coreProperties>`);
    pkgRels.push('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>');
    overrides.push(['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml']);
  }
  const relsXml = (list) => `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${/\/hyperlink$/.test(r.type) ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
  zip.file('word/_rels/document.xml.rels', relsXml([...docRels, ...(relsByPart.get('document.xml') || [])]));
  // связи колонтитулов, сносок, примечаний — у каждой части свои
  for (const [part, list] of relsByPart) if (part !== 'document.xml' && list.length) zip.file(`word/_rels/${part}.rels`, relsXml(list));
  zip.file('_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${pkgRels.join('')}</Relationships>`);
  const exts = [...new Set(media.map((m) => m.ext))];
  const CT = { jpeg: 'image/jpeg', png: 'image/png', bmp: 'image/bmp', tiff: 'image/tiff' };
  zip.file('[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${exts.map((e) => `<Default Extension="${e}" ContentType="${CT[e]}"/>`).join('')}${overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('')}</Types>`);
  const docx = await zip.generateAsync({ type: 'uint8array' });
  if (cfb.entries.some((e) => /^Macros|_VBA_PROJECT/i.test(e.path))) warn('макросы VBA');
  return { docx, warnings };
}
