// Быстрый разбор листа .xlsx. Почти весь объём листа — <sheetData> (строки и ячейки); xmldom 0.9 на каждом закрывающем
// теге собирает новый RegExp, и книга на 100 000 строк × 10 столбцов разбиралась ~40 с. Здесь <sheetData> разбирается
// сканером строк (регулярные выражения по тексту), а остальное (столбцы, объединения, закрепление, УФ, проверка…) —
// как раньше DOM-ом, но уже без <sheetData>. Строки и ячейки сканера отвечают на getAttribute, как элементы DOM.
import { DOMParser } from '@xmldom/xmldom';

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
export const unescapeXml = (s) => (s.indexOf('&') < 0 ? s : s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (all, e) => (
  e[0] === '#' ? String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e])));

// Атрибуты открывающего тега → объект; префиксы пространств имён у атрибутов не нужны (r, s, t, ht…)
function attrsOf(s) {
  const a = {};
  if (!s) return a;
  const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(s))) a[m[1].includes(':') && !m[1].startsWith('xml') ? m[1].split(':').pop() : m[1]] = unescapeXml(m[2] ?? m[3]);
  return a;
}
// Элемент с атрибутами — как DOM для readLineInfo/isOn и т. п.
class El {
  constructor(attrs) { this.attrs = attrs; }
  getAttribute(n) { return Object.prototype.hasOwnProperty.call(this.attrs, n) ? this.attrs[n] : null; }
}

const P = '(?:[\\w-]+:)?'; // необязательный префикс (x:row — некоторые программы пишут SpreadsheetML с префиксом)
const ROW = new RegExp(`<${P}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${P}row>)`, 'g');
const CELL = new RegExp(`<${P}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${P}c>)`, 'g');
const V = new RegExp(`<${P}v\\b[^>]*?(?:/>|>([\\s\\S]*?)</${P}v>)`);
const F = new RegExp(`<${P}f\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${P}f>)`);
const IS = new RegExp(`<${P}is\\b[^>]*?(?:/>|>([\\s\\S]*?)</${P}is>)`);
const T = new RegExp(`<${P}t\\b[^>]*?(?:/>|>([\\s\\S]*?)</${P}t>)`, 'g');
const RUN = new RegExp(`<${P}r\\b[\\s\\S]*?</${P}r>`, 'g');
const textOf = (s) => unescapeXml(s.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (a, x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;')).replace(/<[^>]*>/g, ''));

// Содержимое <sheetData> → [{ el, cells: [{ el, v, f: { el, text } | null, is }] }]
export function scanSheetData(inner) {
  const rows = [];
  if (!inner) return rows;
  ROW.lastIndex = 0;
  let rm;
  while ((rm = ROW.exec(inner))) {
    const cells = [];
    const body = rm[2];
    if (body) {
      CELL.lastIndex = 0;
      let cm;
      while ((cm = CELL.exec(body))) {
        const content = cm[2] || '';
        let v = null; let f = null; let is = null;
        if (content) {
          const vm = V.exec(content);
          if (vm) v = vm[1] != null ? textOf(vm[1]) : '';
          const fm = F.exec(content);
          if (fm) f = { el: new El(attrsOf(fm[1])), text: fm[2] != null ? textOf(fm[2]) : '' };
          const im = IS.exec(content);
          if (im) {
            const isBody = im[1] || '';
            // как DOM: прямые <t> внутри <is>; иначе — весь текст (форматированный текст: <r><t>…)
            const direct = isBody.replace(RUN, '').replace(new RegExp(`<${P}rPh\\b[\\s\\S]*?</${P}rPh>`, 'g'), '');
            const ts = [...direct.matchAll(T)].map((x) => (x[1] != null ? textOf(x[1]) : ''));
            is = { direct: ts.join(''), hasDirect: ts.length > 0, all: textOf(isBody) };
          }
        }
        cells.push({ el: new El(attrsOf(cm[1])), v, f, is });
      }
    }
    rows.push({ el: new El(attrsOf(rm[1])), cells });
  }
  return rows;
}

// Найти <sheetData> в тексте листа: { start, innerStart, innerEnd, end } или null
export function findSheetData(xml) {
  const open = new RegExp(`<${P}sheetData\\b[^>]*?(/?)>`).exec(xml);
  if (!open) return null;
  const start = open.index;
  const innerStart = start + open[0].length;
  if (open[1] === '/') return { start, innerStart, innerEnd: innerStart, end: innerStart };
  const close = new RegExp(`</${P}sheetData>`, 'g');
  close.lastIndex = innerStart;
  const cm = close.exec(xml);
  if (!cm) return null;
  return { start, innerStart, innerEnd: cm.index, end: cm.index + cm[0].length };
}

// Лист: DOM без содержимого <sheetData> (root.__sheetData — строки сканера) или null, если это не лист
export function parseSheet(xml, { data = true } = {}) {
  if (xml == null) return null;
  const sd = findSheetData(xml);
  const rest = sd ? xml.slice(0, sd.innerStart) + xml.slice(sd.innerEnd) : xml;
  const doc = new DOMParser().parseFromString(rest, 'text/xml');
  if (doc && doc.documentElement) doc.documentElement.__sheetData = sd && data ? scanSheetData(xml.slice(sd.innerStart, sd.innerEnd)) : [];
  return doc;
}
