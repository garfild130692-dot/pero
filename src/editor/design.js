// Вкладка «Конструктор»: подложка, цвет страницы, границы страниц, интервал между абзацами и шрифт документа.
//
// Подложка — doc.attrs.watermark { text, font, color, semi, layout: 'diagonal' | 'horizontal', size (pt | null — авто) },
// watermarkChanged — задана/изменена в Пере. В файле — как у Word: надпись VML (v:shape id «PowerPlusWaterMarkObject…»,
// textpath) в колонтитулах; при сохранении она пишется во все верхние колонтитулы (нет ни одного — создаётся свой).
// Цвет страницы — doc.attrs.pageColor ↔ w:background w:color перед телом документа (+ displayBackgroundShape в settings).
// Границы страниц — w:pgBorders в sectPr каждого раздела (атрибуты sectPr, как поля и размер листа).
import { Extension } from '@tiptap/core';

// символы, недопустимые в XML 1.0, выбрасываются (как в export.js) — иначе текст из чужого файла ломает часть
const esc = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// ---------- Подложка ----------
export const WATERMARK_PRESETS = ['ОБРАЗЕЦ', 'ЧЕРНОВИК', 'КОПИЯ', 'КОНФИДЕНЦИАЛЬНО', 'ДЛЯ СЛУЖЕБНОГО ПОЛЬЗОВАНИЯ', 'НЕ КОПИРОВАТЬ', 'СРОЧНО'];
export const WATERMARK_DEFAULT = { text: 'ОБРАЗЕЦ', font: 'Calibri', color: '#c0c0c0', semi: true, layout: 'diagonal', size: null };
const NAMED = { silver: '#c0c0c0', gray: '#808080', black: '#000000', red: '#ff0000', blue: '#0000ff', white: '#ffffff' };
const WM_RE = /PowerPlusWaterMarkObject/;

// Подложка из XML колонтитула (null — нет)
export function readWatermark(xml) {
  const s = String(xml || '');
  const i = s.search(WM_RE);
  if (i < 0) return null;
  const shapeStart = s.lastIndexOf('<v:shape ', i);
  const shapeEnd = s.indexOf('</v:shape>', i);
  const shape = s.slice(shapeStart, shapeEnd > 0 ? shapeEnd : undefined);
  const text = (/string="([^"]*)"/.exec(shape) || [])[1];
  if (text == null) return null;
  const style = (/<v:shape\b[^>]*style="([^"]*)"/.exec(shape) || [])[1] || '';
  const tp = (/<v:textpath\b[^>]*style="([^"]*)"/.exec(shape) || [])[1] || '';
  const font = (/font-family:(?:&quot;|")?([^;&"]+)/.exec(tp) || [])[1] || 'Calibri';
  const fs = parseFloat((/font-size:([\d.]+)pt/.exec(tp) || [])[1] || '1');
  let color = (/fillcolor="([^"]*)"/.exec(shape) || [])[1] || 'silver';
  color = NAMED[color.toLowerCase()] || (/^#[0-9a-f]{6}$/i.test(color) ? color.toLowerCase() : /^#[0-9a-f]{3}$/i.test(color) ? `#${[...color.slice(1)].map((c) => c + c).join('')}`.toLowerCase() : '#c0c0c0');
  const opacity = parseFloat((/<v:fill\b[^>]*opacity="([^"]*)"/.exec(shape) || [])[1] || '1');
  const rot = parseFloat((/rotation:(-?[\d.]+)/.exec(style) || [])[1] || '0');
  return {
    text: unesc(text), font: unesc(font).trim(), color, semi: opacity < 1,
    layout: Math.abs(((rot % 360) + 360) % 360 - 315) < 1 ? 'diagonal' : 'horizontal', size: fs > 1 ? fs : null,
  };
}

// Размер надписи, pt. Авто — как у Word: рамка 18,62 × 4,65 см (при ширине текста A4 с полями по умолчанию — 467,75 pt),
// надпись растягивается в неё (fitshape); для другой ширины текста — пропорционально. Заданный размер — по числу знаков.
export function watermarkBox(wm, textWidthPt) {
  const n = Math.max(1, [...String(wm.text || '')].length);
  if (wm.size) return { w: Math.round(wm.size * n * 0.75 * 10) / 10, h: Math.round(wm.size * 1.1 * 10) / 10 };
  const w = Math.round(527.85 * (textWidthPt / 467.75) * 10) / 10;
  return { w, h: Math.round((w / 4) * 10) / 10 };
}

const SHAPETYPE = '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas>' +
  ['sum #0 0 10800', 'prod #0 2 1', 'sum 21600 0 @1', 'sum 0 0 @2', 'sum 21600 0 @3', 'if @0 @3 0', 'if @0 21600 @1', 'if @0 0 @2', 'if @0 @4 21600',
    'mid @5 @6', 'mid @8 @5', 'mid @7 @8', 'mid @6 @7', 'sum @6 0 @5'].map((e) => `<v:f eqn="${e}"/>`).join('') +
  '</v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>' +
  '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

// Прогон с подложкой (как пишет Word); пространства имён VML — на самом w:pict (в части колонтитула их может не быть)
export function watermarkRunXml(wm, box, id = 1) {
  const style = `position:absolute;margin-left:0;margin-top:0;width:${box.w}pt;height:${box.h}pt;${wm.layout === 'diagonal' ? 'rotation:315;' : ''}` +
    'z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin';
  return '<w:r><w:rPr><w:noProof/></w:rPr><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word">' +
    `${SHAPETYPE}<v:shape id="PowerPlusWaterMarkObject${id}" o:spid="_x0000_s${2048 + id}" type="#_x0000_t136" style="${style}" o:allowincell="f" fillcolor="${esc(wm.color || '#c0c0c0')}" stroked="f">` +
    `${wm.semi ? '<v:fill opacity=".5"/>' : ''}<v:textpath style="font-family:&quot;${esc(wm.font || 'Calibri')}&quot;;font-size:${wm.size || 1}pt" string="${esc(wm.text)}"/>` +
    '<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>';
}

// Убрать подложки из XML колонтитула (прогоны с PowerPlusWaterMarkObject)
export function stripWatermark(xml) {
  let s = String(xml);
  for (let guard = 0; guard < 20; guard++) {
    const i = s.search(WM_RE);
    if (i < 0) break;
    const a = Math.max(s.lastIndexOf('<w:r>', i), s.lastIndexOf('<w:r ', i));
    const b = s.indexOf('</w:r>', i);
    if (a < 0 || b < 0) break;
    s = s.slice(0, a) + s.slice(b + 6);
  }
  return s;
}

// Вставить прогон подложки в первый абзац колонтитула (нет абзаца — новый абзац)
export function addWatermark(xml, run) {
  const s = stripWatermark(xml);
  const empty = /<w:p\b[^>]*\/>/.exec(s);
  const close = s.indexOf('</w:p>');
  if (close >= 0 && (!empty || close < empty.index)) return s.slice(0, close) + run + s.slice(close);
  if (empty) return s.slice(0, empty.index) + empty[0].replace(/\/>$/, '>') + run + '</w:p>' + s.slice(empty.index + empty[0].length);
  return s.replace(/(<\/w:hdr>)\s*$/, `<w:p>${run}</w:p>$1`);
}

// ---------- Цвет страницы ----------
export const readBackground = (preBody) => {
  const m = /<w:background\b[^>]*w:color="([0-9A-Fa-f]{6})"/.exec(preBody || '');
  return m ? `#${m[1].toLowerCase()}` : null;
};
export function patchBackground(preBody, color) {
  const s = String(preBody || '').replace(/<w:background\b[^>]*\/>|<w:background\b[^>]*>[\s\S]*?<\/w:background>/, '');
  return color ? `<w:background w:color="${color.slice(1).toUpperCase()}"/>${s}` : s;
}
// w:displayBackgroundShape — на своём месте по схеме settings (после view/zoom/… doNotDisplayPageBoundaries)
export function patchSettingsBackground(xml, on) {
  const s = String(xml).replace(/<w:displayBackgroundShape\b[^>]*\/>/g, '');
  if (!on) return s;
  let at = -1;
  for (const tag of ['doNotDisplayPageBoundaries', 'removeDateAndTime', 'removePersonalInformation', 'zoom', 'view', 'writeProtection']) {
    const m = new RegExp(`<w:${tag}\\b[^>]*(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(s);
    if (m) { at = m.index + m[0].length; break; }
  }
  if (at < 0) { const m = /<w:settings\b[^>]*>/.exec(s); at = m.index + m[0].length; }
  return s.slice(0, at) + '<w:displayBackgroundShape/>' + s.slice(at);
}

// ---------- Границы страниц ----------
export const PAGE_BORDER_STYLES = [['single', 'Сплошная'], ['double', 'Двойная'], ['dotted', 'Точками'], ['dashed', 'Штрихами'], ['thick', 'Толстая']];
const SIDES = ['top', 'left', 'bottom', 'right'];

// { offsetFrom: 'page' | 'text', style, width (pt), color, space (pt) } — одинаковые стороны (как «Рамка» в Word); null — нет
export function parsePgBorders(sectXml) {
  const m = /<w:pgBorders\b([^>]*)>([\s\S]*?)<\/w:pgBorders>/.exec(sectXml || '');
  if (!m) return null;
  const top = /<w:top\b[^>]*\/>/.exec(m[2]) || /<w:(?:left|bottom|right)\b[^>]*\/>/.exec(m[2]);
  if (!top) return null;
  const a = (n) => (new RegExp(`w:${n}="([^"]*)"`).exec(top[0]) || [])[1];
  if (!a('val') || a('val') === 'none' || a('val') === 'nil') return null;
  const color = a('color');
  return {
    offsetFrom: /w:offsetFrom="page"/.test(m[1]) ? 'page' : 'text',
    style: a('val'),
    width: (+a('sz') || 4) / 8,
    color: color && color !== 'auto' ? `#${color.toLowerCase()}` : null,
    space: +a('space') || 0,
  };
}

export function pgBordersXml(b) {
  if (!b) return '';
  const side = (s) => `<w:${s} w:val="${b.style}" w:sz="${Math.round((b.width || 0.5) * 8)}" w:space="${Math.round(b.space || 0)}" w:color="${b.color ? b.color.slice(1).toUpperCase() : 'auto'}"/>`;
  return `<w:pgBorders${b.offsetFrom === 'page' ? ' w:offsetFrom="page"' : ''}>${SIDES.map(side).join('')}</w:pgBorders>`;
}

// Заменить границы страниц в sectPr: по схеме — после pgMar/paperSrc, перед lnNumType/pgNumType/cols…
export function patchPgBorders(sectXml, b) {
  let s = String(sectXml).replace(/<w:pgBorders\b[^>]*>[\s\S]*?<\/w:pgBorders>|<w:pgBorders\b[^>]*\/>/, '');
  if (!b) return s;
  const x = pgBordersXml(b);
  const after = /<w:paperSrc\b[^>]*\/>/.exec(s) || /<w:pgMar\b[^>]*\/>/.exec(s) || /<w:pgSz\b[^>]*\/>/.exec(s);
  if (after) return s.slice(0, after.index + after[0].length) + x + s.slice(after.index + after[0].length);
  return s.replace(/(<w:sectPr\b[^>]*>)/, `$1${x}`);
}

// Рамка листа на экране/при печати: прямоугольник от края листа (offsetFrom page) или вокруг текста, pt
export function pageBorderRect(b, g) {
  const W = g.w / 20; const H = g.h / 20;
  if (b.offsetFrom === 'page') return { left: b.space, top: b.space, right: W - b.space, bottom: H - b.space };
  return { left: g.left / 20 - b.space, top: g.top / 20 - b.space, right: W - g.right / 20 + b.space, bottom: H - g.bottom / 20 + b.space };
}
export const BORDER_CSS = { single: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', thick: 'solid' };

// ---------- Интервал между абзацами (умолчания документа), шрифт документа ----------
export const SPACING_PRESETS = [
  ['default', 'По умолчанию', { spaceBefore: 0, spaceAfter: 8, lineRule: 'auto', lineSpacing: 1.08 }],
  ['none', 'Без интервала', { spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1 }],
  ['compact', 'Компактный', { spaceBefore: 0, spaceAfter: 4, lineRule: 'auto', lineSpacing: 1 }],
  ['tight', 'Плотный', { spaceBefore: 0, spaceAfter: 6, lineRule: 'auto', lineSpacing: 1.15 }],
  ['open', 'Открытый', { spaceBefore: 0, spaceAfter: 10, lineRule: 'auto', lineSpacing: 1.15 }],
  ['relaxed', 'Свободный', { spaceBefore: 0, spaceAfter: 6, lineRule: 'auto', lineSpacing: 1.5 }],
  ['double', 'Двойной', { spaceBefore: 0, spaceAfter: 8, lineRule: 'auto', lineSpacing: 2 }],
  ['gost', 'Делопроизводство (ГОСТ): 1,5 строки, абзац 1,25 см', { spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1.5, indentFirst: 35.45 }],
];
export const DOC_FONTS = ['Calibri', 'Times New Roman', 'Arial', 'Cambria', 'Georgia', 'Verdana', 'PT Astra Serif', 'PT Sans', 'Liberation Serif'];

export const Design = Extension.create({
  name: 'design',
  addCommands() {
    const setAttrs = (pairs) => ({ tr, dispatch }) => {
      if (dispatch) { for (const [k, v] of pairs) tr.setDocAttribute(k, v); dispatch(tr); }
      return true;
    };
    return {
      setWatermark: (wm) => setAttrs([['watermark', wm ? { ...WATERMARK_DEFAULT, ...wm } : null], ['watermarkChanged', true]]),
      setPageColor: (color) => setAttrs([['pageColor', color || null], ['pageColorChanged', true]]),
      // границы страниц — во всех разделах (как «Применить к: всему документу»)
      setPageBorders: (b) => ({ state, tr, dispatch }) => {
        if (!dispatch) return true;
        state.doc.descendants((node, pos) => {
          if (node.attrs && node.attrs.sectPr) tr.setNodeMarkup(pos, undefined, { ...node.attrs, sectPr: patchPgBorders(node.attrs.sectPr, b) });
          return node.isBlock && !node.isTextblock;
        });
        tr.setDocAttribute('sectPr', patchPgBorders(state.doc.attrs.sectPr || DEFAULT_SECT, b));
        dispatch(tr);
        return true;
      },
      setParaSpacing: (key) => ({ state, tr, dispatch }) => {
        const p = SPACING_PRESETS.find(([k]) => k === key);
        if (!p) return false;
        const para = { ...(state.doc.attrs.para || {}), indentFirst: 0, ...p[2] };
        if (dispatch) dispatch(tr.setDocAttribute('para', para).setDocAttribute('defaultsModified', true));
        return true;
      },
      setDocFont: (name) => setAttrs([['fontFamily', name], ['defaultsModified', true]]),
    };
  },
});

// sectPr нового документа (как в export.js) — чтобы рамку можно было задать и до первого сохранения
const DEFAULT_SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
