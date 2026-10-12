// Фигуры и надписи Word (wps:wsp внутри mc:AlternateContent / w:drawing): показ на экране. Узел остаётся opaque
// (исходный XML пишется в файл как был) — здесь только разбор и вид: геометрия prstGeom — SVG, заливка и линия из
// wps:spPr (иначе — из стиля фигуры: lnRef/fillRef/fontRef с цветами темы документа), текст надписи (w:txbxContent) —
// абзацы с полужирным/курсивом/размером/цветом/выравниванием, поля и вертикальное выравнивание — wps:bodyPr.
// Положение и обтекание — те же, что у рисунков (anchor.js, floatLayout.js).
import { readAnchor } from './anchor.js';

const EMU = 9525; // EMU в px
const px = (v) => (+v || 0) / EMU;

// Цвета темы Office по умолчанию (2013–2022); документ передаёт свои из theme1.xml (setThemeColors)
const OFFICE = { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72' };
let THEME = OFFICE;
const ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
// цвета темы открытого документа (doc.attrs.themeColors; null — тема Office)
export function setThemeColors(c) {
  const next = c && c.accent1 ? { ...OFFICE, ...c } : OFFICE;
  if (JSON.stringify(next) !== JSON.stringify(THEME)) { THEME = next; cache.clear(); }
}
export function themeColorsFromXml(xml) {
  const out = {};
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml || '');
  if (!scheme) return out;
  for (const m of scheme[0].matchAll(/<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr\b[^>]*lastClr="([0-9A-Fa-f]{6})")/g)) out[m[1]] = (m[2] || m[3]).toUpperCase();
  return out;
}
const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };

// ---------- цвет DrawingML ----------
const hex2rgb = (h) => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
const rgb2hex = (c) => c.map((v) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0')).join('');
function rgb2hsl([r, g, b]) {
  const mx = Math.max(r, g, b); const mn = Math.min(r, g, b); const l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn; const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function hsl2rgb([h, s, l]) {
  if (!s) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
}
// Цвет из элемента-заливки (<a:solidFill>…</a:solidFill> или само цветовое содержимое): 'rgba(…)' или null
function colorOf(xml) {
  if (!xml) return null;
  const m = /<a:(srgbClr|schemeClr|prstClr|sysClr)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:\1>)/.exec(xml);
  if (!m) return null;
  const val = (/\bval="([^"]*)"/.exec(m[2]) || [])[1];
  let hex = m[1] === 'srgbClr' ? val : m[1] === 'schemeClr' ? THEME[ALIAS[val] || val] : m[1] === 'prstClr' ? PRESET[val] : ((/lastClr="([^"]*)"/.exec(m[2]) || [])[1] || '000000');
  if (!hex) hex = '000000';
  let rgb = hex2rgb(hex);
  let alpha = 1;
  for (const t of (m[3] || '').matchAll(/<a:(\w+)\s+val="(-?\d+)"\s*\/>/g)) {
    const v = +t[2] / 100000;
    if (t[1] === 'shade') rgb = rgb.map((c) => c * v);
    else if (t[1] === 'tint') rgb = rgb.map((c) => c * v + (1 - v));
    else if (t[1] === 'lumMod' || t[1] === 'lumOff') { const h = rgb2hsl(rgb); h[2] = t[1] === 'lumMod' ? h[2] * v : Math.min(1, h[2] + v); rgb = hsl2rgb(h); }
    else if (t[1] === 'alpha') alpha = v;
  }
  return alpha < 1 ? `rgba(${rgb.map((c) => Math.round(c * 255)).join(',')},${alpha})` : `#${rgb2hex(rgb)}`;
}

// ---------- разбор фигуры ----------
const pick = (xml, re) => { const m = re.exec(xml || ''); return m ? m[0] : null; };

// XML прогона с фигурой → описание для показа или null (не фигура wps: группы, диаграммы, рисунки — значком)
export function parseShape(xml) {
  const choice = pick(xml, /<mc:Choice\b[^>]*Requires="wps"[^>]*>[\s\S]*?<\/mc:Choice>/) || xml;
  const wsp = pick(choice, /<wps:wsp\b[\s\S]*?<\/wps:wsp>/);
  if (!wsp) return null;
  const frame = pick(choice, /<wp:(inline|anchor)\b[\s\S]*<\/wp:\1>/) || choice;
  const ext = /<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(frame);
  if (!ext) return null;
  // wp:effectExtent — запас под линию, тень, поворот: Word резервирует его в строке и при обтекании
  const ee = /<wp:effectExtent\b([^>]*)\/>/.exec(frame);
  const eff = (n) => { const m = ee && new RegExp(`\\b${n}="(-?\\d+)"`).exec(ee[1]); return m ? Math.max(0, px(m[1])) : 0; };
  const sp = pick(wsp, /<wps:spPr\b[\s\S]*?<\/wps:spPr>/) || '';
  const style = pick(wsp, /<wps:style\b[\s\S]*?<\/wps:style>/) || '';
  const geom = (/<a:prstGeom\b[^>]*prst="([^"]+)"/.exec(sp) || [])[1] || (/<a:custGeom\b/.test(sp) ? 'custom' : 'rect');
  const adj = {};
  for (const g of (pick(sp, /<a:avLst\b[\s\S]*?<\/a:avLst>/) || '').matchAll(/<a:gd\b[^>]*name="(\w+)"[^>]*fmla="val (-?\d+)"/g)) adj[g[1]] = +g[2];
  // заливка: spPr (noFill / solidFill) или стиль (fillRef idx > 0)
  const spNoLn = pick(sp, /<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/);
  const spFill = sp.replace(spNoLn || '', '').replace(/<a:extLst\b[\s\S]*?<\/a:extLst>/g, '');
  let fill = null;
  if (/<a:noFill\/>/.test(spFill)) fill = null;
  else if (/<a:solidFill>/.test(spFill)) fill = colorOf(pick(spFill, /<a:solidFill>[\s\S]*?<\/a:solidFill>/));
  else if (/<a:gradFill\b/.test(spFill)) fill = colorOf(pick(spFill, /<a:gs\b[\s\S]*?<\/a:gs>/)); // градиент — первым цветом
  else { const fr = pick(style, /<a:fillRef\b[^>]*idx="(\d+)"[\s\S]*?<\/a:fillRef>/); if (fr && !/idx="0"/.test(fr)) fill = colorOf(fr); }
  // линия: spPr/a:ln (noFill, solidFill, w) или стиль lnRef
  let line = null;
  const ln = spNoLn || '';
  const lnW = (/<a:ln\b[^>]*\bw="(\d+)"/.exec(ln) || [])[1];
  if (!/<a:noFill\/>/.test(ln)) {
    let color = /<a:solidFill>/.test(ln) ? colorOf(pick(ln, /<a:solidFill>[\s\S]*?<\/a:solidFill>/)) : null;
    if (!color) { const lr = pick(style, /<a:lnRef\b[^>]*idx="(\d+)"[\s\S]*?<\/a:lnRef>/); if (lr && !/idx="0"/.test(lr)) color = colorOf(lr); }
    if (color) {
      line = {
        color, width: lnW ? px(lnW) : 1,
        dash: (/<a:prstDash val="(\w+)"/.exec(ln) || [])[1] || 'solid',
        head: (/<a:headEnd\b[^>]*type="(\w+)"/.exec(ln) || [])[1] || null,
        tail: (/<a:tailEnd\b[^>]*type="(\w+)"/.exec(ln) || [])[1] || null,
      };
    }
  }
  const xf = /<a:xfrm\b([^>]*)>/.exec(sp);
  const rot = xf && /\brot="(-?\d+)"/.exec(xf[1]);
  const body = (/<wps:bodyPr\b([^>]*)/.exec(wsp) || [])[1] || '';
  const ins = (n, d) => { const m = new RegExp(`\\b${n}="(\\d+)"`).exec(body); return px(m ? m[1] : d); };
  const fontColor = colorOf(pick(style, /<a:fontRef\b[\s\S]*?<\/a:fontRef>/)) || '#000000';
  return {
    w: px(ext[1]), h: px(ext[2]),
    effect: { l: eff('l'), t: eff('t'), r: eff('r'), b: eff('b') },
    anchor: readAnchor(frame),
    geom, adj, fill, line,
    rotate: rot ? ((+rot[1] / 60000) % 360 + 360) % 360 : 0,
    flipH: !!(xf && /flipH="1"/.test(xf[1])), flipV: !!(xf && /flipV="1"/.test(xf[1])),
    inset: { l: ins('lIns', 91440), t: ins('tIns', 45720), r: ins('rIns', 91440), b: ins('bIns', 45720) },
    vAnchor: (/\banchor="(\w+)"/.exec(body) || [])[1] || 't',
    vert: (/\bvert="(\w+)"/.exec(body) || [])[1] || 'horz',
    paras: readTextbox(pick(wsp, /<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/), fontColor),
  };
}

const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
// Текст надписи: [{ align, runs: [{ text, bold, italic, underline, size (pt), color }] }]
function readTextbox(xml, fontColor) {
  if (!xml) return [];
  const paras = [];
  for (const p of xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>|<w:p\b[^>]*\/>/g)) {
    const pPr = pick(p[0], /<w:pPr\b[\s\S]*?<\/w:pPr>/) || '';
    const jc = (/<w:jc w:val="(\w+)"/.exec(pPr) || [])[1];
    const runs = [];
    for (const r of p[0].matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)) {
      const rPr = pick(r[0], /<w:rPr\b[\s\S]*?<\/w:rPr>/) || '';
      const text = [...r[0].matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:br\/>/g)].map((t) => (t[1] != null ? unesc(t[1]) : t[0] === '<w:tab/>' ? '\t' : '\n')).join('');
      if (!text) continue;
      const on = (tag) => new RegExp(`<w:${tag}(?: w:val="(?:1|true|on)")?\\s*/>`).test(rPr);
      const sz = /<w:sz w:val="(\d+)"/.exec(rPr);
      const col = /<w:color w:val="([0-9A-Fa-f]{6})"/.exec(rPr);
      runs.push({ text, bold: on('b'), italic: on('i'), underline: /<w:u w:val="(?!none)/.test(rPr), size: sz ? +sz[1] / 2 : null, color: col ? `#${col[1]}` : fontColor });
    }
    paras.push({ align: jc === 'both' ? 'justify' : jc === 'end' ? 'right' : jc === 'start' ? 'left' : jc || 'left', runs });
  }
  return paras;
}

// ---------- геометрия ----------
// Путь SVG фигуры w×h по prstGeom (основные фигуры; прочие — прямоугольник)
function shapePath(g, w, h, adj) {
  const a = (name, d) => (adj[name] ?? d) / 100000;
  const ss = Math.min(w, h);
  const poly = (pts) => `M${pts.map((p) => p.map((v) => v.toFixed(2)).join(' ')).join('L')}Z`;
  switch (g) {
    case 'ellipse': return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
    case 'roundRect': { const r = ss * a('adj', 16667); return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w} ${r}V${h - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 0 ${h - r}V${r}A${r} ${r} 0 0 1 ${r} 0Z`; }
    case 'triangle': return poly([[w * a('adj', 50000), 0], [w, h], [0, h]]);
    case 'rtTriangle': return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond': return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': { const d = ss * a('adj', 25000); return poly([[d, 0], [w, 0], [w - d, h], [0, h]]); }
    case 'trapezoid': { const d = ss * a('adj', 25000); return poly([[d, 0], [w - d, 0], [w, h], [0, h]]); }
    case 'hexagon': { const d = ss * a('adj', 25000); return poly([[d, 0], [w - d, 0], [w, h / 2], [w - d, h], [d, h], [0, h / 2]]); }
    case 'octagon': { const d = ss * a('adj', 29289); return poly([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]); }
    case 'pentagon': return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'homePlate': { const d = ss * a('adj', 50000); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    case 'chevron': { const d = ss * a('adj', 50000); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case 'plus': { const d = ss * a('adj', 25000); return poly([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]); }
    case 'star5': {
      const pts = [];
      for (let i = 0; i < 10; i++) {
        const r = i % 2 ? 0.382 : 1; const t = -Math.PI / 2 + i * Math.PI / 5;
        pts.push([w / 2 + (w / 2) * r * Math.cos(t), h / 2 + (h / 2) * r * Math.sin(t) * 1.05 + h * 0.025]);
      }
      return poly(pts);
    }
    case 'rightArrow': case 'leftArrow': {
      const t = h * a('adj1', 50000); const hl = ss * a('adj2', 50000);
      const p = [[0, (h - t) / 2], [w - hl, (h - t) / 2], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, (h + t) / 2], [0, (h + t) / 2]];
      return poly(g === 'leftArrow' ? p.map(([x, y]) => [w - x, y]) : p);
    }
    case 'upArrow': case 'downArrow': {
      const t = w * a('adj1', 50000); const hl = ss * a('adj2', 50000);
      const p = [[(w - t) / 2, h], [(w - t) / 2, hl], [0, hl], [w / 2, 0], [w, hl], [(w + t) / 2, hl], [(w + t) / 2, h]];
      return poly(g === 'downArrow' ? p.map(([x, y]) => [x, h - y]) : p);
    }
    default: return `M0 0H${w}V${h}H0Z`;
  }
}

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const DASH = { dash: '4 3', sysDash: '3 1', dot: '1 2', sysDot: '1 1', lgDash: '8 3', dashDot: '4 3 1 3', lgDashDot: '8 3 1 3' };

// HTML фигуры: SVG геометрии + текст надписи поверх (поля и вертикальное выравнивание bodyPr)
export function shapeHtml(s) {
  const { w, h } = s;
  const isLine = s.geom === 'line' || /Connector/.test(s.geom);
  const lw = s.line ? Math.max(0.5, s.line.width) : 0;
  const dash = s.line && DASH[s.line.dash] ? ` stroke-dasharray="${DASH[s.line.dash].split(' ').map((v) => v * Math.max(1, lw)).join(' ')}"` : '';
  const marker = (id, which) => `<marker id="${id}" viewBox="0 0 10 10" refX="${which === 'end' ? 8 : 2}" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse"><path d="M0 0L10 5L0 10Z" fill="${s.line.color}"/></marker>`;
  let svg;
  if (isLine) {
    const id = `pm${Math.random().toString(36).slice(2, 8)}`;
    const defs = s.line ? `${s.line.tail ? marker(`${id}t`, 'end') : ''}${s.line.head ? marker(`${id}h`, 'start') : ''}` : '';
    const [x1, y1, x2, y2] = [s.flipH ? w : 0, s.flipV ? h : 0, s.flipH ? 0 : w, s.flipV ? 0 : h];
    svg = `<svg class="pero-shape-svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible"><defs>${defs}</defs>` +
      `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${s.line ? s.line.color : '#000'}" stroke-width="${lw || 1}"${dash}` +
      `${s.line && s.line.tail ? ` marker-end="url(#${id}t)"` : ''}${s.line && s.line.head ? ` marker-start="url(#${id}h)"` : ''}/></svg>`;
  } else {
    const flip = s.flipH || s.flipV ? ` transform="translate(${s.flipH ? w : 0} ${s.flipV ? h : 0}) scale(${s.flipH ? -1 : 1} ${s.flipV ? -1 : 1})"` : '';
    svg = `<svg class="pero-shape-svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible"><path d="${shapePath(s.geom, w, h, s.adj)}"${flip}` +
      ` fill="${s.fill || 'none'}" stroke="${s.line ? s.line.color : 'none'}" stroke-width="${lw}"${dash}/></svg>`;
  }
  let text = '';
  if (s.paras.some((p) => p.runs.length)) {
    const just = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[s.vAnchor] || 'flex-start';
    const ps = s.paras.map((p) => `<div style="text-align:${p.align}">${p.runs.map((r) => `<span style="${[
      r.bold ? 'font-weight:bold' : '', r.italic ? 'font-style:italic' : '', r.underline ? 'text-decoration:underline' : '',
      r.size ? `font-size:${r.size}pt` : '', `color:${r.color}`].filter(Boolean).join(';')}">${esc(r.text)}</span>`).join('') || '&nbsp;'}</div>`).join('');
    text = `<div class="pero-shape-text" style="padding:${s.inset.t}px ${s.inset.r}px ${s.inset.b}px ${s.inset.l}px;justify-content:${just}${s.vert !== 'horz' ? ';writing-mode:vertical-rl' : ''}">${ps}</div>`;
  }
  return `${svg}${text}`;
}

// Кеш разбора по XML (renderHTML значка вызывается при каждой перерисовке абзаца)
const cache = new Map();
export function shapeOf(xml) {
  if (!cache.has(xml)) { if (cache.size > 300) cache.clear(); cache.set(xml, parseShape(xml)); }
  return cache.get(xml);
}
