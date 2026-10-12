// Плавающий рисунок (wp:anchor): обтекание и положение на странице. Общий модуль для импорта, экспорта и показа.
//
// anchor = {
//   wrap: 'square' | 'tight' | 'through' | 'topBottom' | 'none',   // wp:wrapSquare … wp:wrapNone
//   side: 'bothSides' | 'left' | 'right' | 'largest',               // wrapText (с какой стороны обтекает текст)
//   behind: bool,                                                     // behindDoc: «за текстом» (при wrap none)
//   h: { rel, align, off }, v: { rel, align, off },                   // positionH/V: относительно чего, align или смещение (px)
//   dist: { t, b, l, r },                                             // отступ от текста, px (distT/B/L/R)
// }
// h.rel: column | margin | page | character | leftMargin | rightMargin | insideMargin | outsideMargin
// v.rel: paragraph | line | margin | page | topMargin | bottomMargin | insideMargin | outsideMargin
// 1 px = 9525 EMU (как размер рисунка).
import { EMU_PER_PX } from './image.js';

const px = (emu) => Math.round((+emu || 0) / EMU_PER_PX * 100) / 100;
const emu = (v) => Math.round((+v || 0) * EMU_PER_PX);
const WRAP_TAG = { none: 'wrapNone', square: 'wrapSquare', tight: 'wrapTight', through: 'wrapThrough', topBottom: 'wrapTopAndBottom' };
const TAG_WRAP = Object.fromEntries(Object.entries(WRAP_TAG).map(([k, v]) => [v, k]));

// Обтекание «В тексте» и пресеты меню (Word: «Обтекание текстом»)
export const WRAP_KINDS = [
  ['inline', 'В тексте'], ['square', 'Вокруг рамки'], ['tight', 'По контуру'], ['through', 'Сквозное'],
  ['topBottom', 'Сверху и снизу'], ['behind', 'За текстом'], ['front', 'Перед текстом'],
];
export const wrapKind = (a) => (!a ? 'inline' : a.wrap === 'none' ? (a.behind ? 'behind' : 'front') : a.wrap);

const attr = (tag, name) => { const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag || ''); return m ? m[1] : null; };

// Из XML прогона/рисунка (<w:drawing>…<wp:anchor>) — модель; не плавающий — null
export function readAnchor(xml) {
  const open = /<(\w+):anchor\b[^>]*>/.exec(xml || '');
  if (!open) return null;
  const p = open[1];
  const tag = open[0];
  const pos = (axis) => {
    const m = new RegExp(`<${p}:position${axis}\\b[^>]*relativeFrom="([^"]+)"[^>]*>([\\s\\S]*?)</${p}:position${axis}>`).exec(xml);
    if (!m) return { rel: axis === 'H' ? 'column' : 'paragraph', align: null, off: 0 };
    const al = new RegExp(`<${p}:align>([^<]*)</${p}:align>`).exec(m[2]);
    const of = new RegExp(`<${p}:posOffset>(-?\\d+)</${p}:posOffset>`).exec(m[2]);
    return { rel: m[1], align: al ? al[1] : null, off: of ? px(of[1]) : 0 };
  };
  const w = new RegExp(`<${p}:(wrapNone|wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom)\\b([^>]*)`).exec(xml);
  return {
    wrap: w ? TAG_WRAP[w[1]] : 'none',
    side: (w && attr(w[2], 'wrapText')) || 'bothSides',
    behind: attr(tag, 'behindDoc') === '1',
    h: pos('H'),
    v: pos('V'),
    dist: { t: px(attr(tag, 'distT')), b: px(attr(tag, 'distB')), l: px(attr(tag, 'distL')), r: px(attr(tag, 'distR')) },
  };
}

// Разметка обтекания: многоугольник «по контуру» — прямоугольник рисунка (как Word по умолчанию у прямоугольных картинок)
function wrapXml(p, a, old) {
  const tag = WRAP_TAG[a.wrap] || 'wrapNone';
  if (tag === 'wrapNone' || tag === 'wrapTopAndBottom') return `<${p}:${tag}/>`;
  const side = ` wrapText="${a.side || 'bothSides'}"`;
  if (tag === 'wrapSquare') return `<${p}:${tag}${side}/>`;
  // по контуру / сквозное: многоугольник — прежний, если был, иначе прямоугольник 21600×21600
  const poly = (old && new RegExp(`<${p}:wrapPolygon\\b[\\s\\S]*?</${p}:wrapPolygon>`).exec(old)) ||
    [`<${p}:wrapPolygon edited="0"><${p}:start x="0" y="0"/><${p}:lineTo x="0" y="21600"/><${p}:lineTo x="21600" y="21600"/><${p}:lineTo x="21600" y="0"/><${p}:lineTo x="0" y="0"/></${p}:wrapPolygon>`];
  return `<${p}:${tag}${side}>${poly[0]}</${p}:${tag}>`;
}
const posXml = (p, axis, q) => `<${p}:position${axis} relativeFrom="${q.rel}">${q.align ? `<${p}:align>${q.align}</${p}:align>` : `<${p}:posOffset>${emu(q.off)}</${p}:posOffset>`}</${p}:position${axis}>`;

// Части рисунка, которые переносятся между inline и anchor как есть
function parts(xml) {
  const p = (/<(\w+):(?:inline|anchor)\b/.exec(xml) || [])[1] || 'wp';
  const pick = (name) => (new RegExp(`<${p}:${name}\\b[^>]*(?:/>|>[\\s\\S]*?</${p}:${name}>)`).exec(xml) || [''])[0];
  const graphic = (/<a:graphic\b[\s\S]*<\/a:graphic>/.exec(xml) || [''])[0];
  return { p, extent: pick('extent'), effectExtent: pick('effectExtent'), docPr: pick('docPr'), frameLocks: pick('cNvGraphicFramePr'), graphic };
}

let relHeight = 251659264;
// Новый wp:anchor вокруг частей рисунка
function anchorFrame(pt, a, old) {
  const { p } = pt;
  const d = a.dist || {};
  const rh = (old && attr(old, 'relativeHeight')) || String(relHeight++);
  return `<${p}:anchor distT="${emu(d.t)}" distB="${emu(d.b)}" distL="${emu(d.l)}" distR="${emu(d.r)}" simplePos="0" relativeHeight="${rh}" ` +
    `behindDoc="${a.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><${p}:simplePos x="0" y="0"/>` +
    posXml(p, 'H', a.h) + posXml(p, 'V', a.v) + pt.extent + (pt.effectExtent || `<${p}:effectExtent l="0" t="0" r="0" b="0"/>`) +
    wrapXml(p, a, old) + pt.docPr + (pt.frameLocks || `<${p}:cNvGraphicFramePr/>`) + pt.graphic + `</${p}:anchor>`;
}
function inlineFrame(pt) {
  const { p } = pt;
  return `<${p}:inline distT="0" distB="0" distL="0" distR="0">${pt.extent}${pt.effectExtent || `<${p}:effectExtent l="0" t="0" r="0" b="0"/>`}` +
    `${pt.docPr}${pt.frameLocks || `<${p}:cNvGraphicFramePr/>`}${pt.graphic}</${p}:inline>`;
}

// Поставить в XML рисунка обтекание/положение (anchor) или сделать его встроенным (anchor = null).
// Неизменная часть (graphic: картинка, обрезка, эффекты; docPr) — как была; у плавающего прежние атрибуты рамки
// (relativeHeight, locked…) сохраняются, меняются положение, обтекание, behindDoc и отступы.
export function setAnchorXml(xml, a) {
  const frame = /<(\w+):(inline|anchor)\b[\s\S]*<\/\1:\2>/.exec(xml);
  if (!frame) return xml;
  const pt = parts(frame[0]);
  let next;
  if (!a) next = frame[2] === 'inline' ? frame[0] : inlineFrame(pt);
  else if (frame[2] === 'inline') next = anchorFrame(pt, a, null);
  else {
    const p = pt.p;
    const d = a.dist || {};
    next = frame[0]
      .replace(/^<(\w+):anchor\b[^>]*>/, (tag) => tag
        .replace(/\bdistT="[^"]*"/, `distT="${emu(d.t)}"`).replace(/\bdistB="[^"]*"/, `distB="${emu(d.b)}"`)
        .replace(/\bdistL="[^"]*"/, `distL="${emu(d.l)}"`).replace(/\bdistR="[^"]*"/, `distR="${emu(d.r)}"`)
        .replace(/\bbehindDoc="[^"]*"/, `behindDoc="${a.behind ? 1 : 0}"`)) // как в файле: у Word бывает и при обтекании
      .replace(new RegExp(`<${p}:positionH\\b[\\s\\S]*?</${p}:positionH>`), posXml(p, 'H', a.h))
      .replace(new RegExp(`<${p}:positionV\\b[\\s\\S]*?</${p}:positionV>`), posXml(p, 'V', a.v))
      .replace(new RegExp(`<${p}:(wrapNone|wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom)\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}:\\1>)`), (old) => wrapXml(p, a, old));
  }
  return xml.slice(0, frame.index) + next + xml.slice(frame.index + frame[0].length);
}

// Отступ от текста по умолчанию у Word: слева и справа 0,32 см (114300 EMU), сверху и снизу 0
export const DEFAULT_DIST = { t: 0, b: 0, l: px(114300), r: px(114300) };
// Новое обтекание для рисунка, который был в тексте (положение — там, где он стоял: от колонки и абзаца)
export function makeAnchor(kind, prev) {
  const base = prev || { wrap: 'square', side: 'bothSides', behind: false, h: { rel: 'column', align: null, off: 0 }, v: { rel: 'paragraph', align: null, off: 0 }, dist: { ...DEFAULT_DIST } };
  if (kind === 'inline') return null;
  if (kind === 'behind' || kind === 'front') return { ...base, wrap: 'none', behind: kind === 'behind' };
  return { ...base, wrap: kind, behind: false };
}
