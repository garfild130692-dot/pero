// Обрезка и поворот рисунка: pic:blipFill/a:srcRect (l t r b — тысячные доли процента: 100000 = 100 %) и a:xfrm рисунка
// (rot — 1/60000 градуса, flipH/flipV). Модель: crop { l, t, r, b } — доли 0…1 от исходной картинки (null — нет),
// rotate — градусы 0…360, flipH/flipV. Порядок DrawingML: сначала отражение, потом поворот.
// width/height узла — размер рамки (уже обрезанной картинки), как wp:extent.

const frac = (v) => Math.round(((+v || 0) / 100000) * 100000) / 100000;

export function readPicFormat(xml) {
  const out = { crop: null, rotate: 0, flipH: false, flipV: false };
  const fill = /<pic:blipFill\b[\s\S]*?<\/pic:blipFill>/.exec(xml || '');
  const sr = fill && /<a:srcRect\b([^>]*)\/?>/.exec(fill[0]);
  if (sr) {
    const g = (n) => { const m = new RegExp(`\\b${n}="(-?\\d+)"`).exec(sr[1]); return m ? frac(m[1]) : 0; };
    const c = { l: g('l'), t: g('t'), r: g('r'), b: g('b') };
    if (c.l || c.t || c.r || c.b) out.crop = c;
  }
  const sp = /<pic:spPr\b[\s\S]*?<\/pic:spPr>/.exec(xml || '');
  const x = sp && /<a:xfrm\b([^>]*)>/.exec(sp[0]);
  if (x) {
    const rot = /\brot="(-?\d+)"/.exec(x[1]);
    if (rot) out.rotate = ((Math.round(+rot[1] / 60000 * 100) / 100) % 360 + 360) % 360;
    out.flipH = /\bflipH="(1|true)"/.test(x[1]);
    out.flipV = /\bflipV="(1|true)"/.test(x[1]);
  }
  return out;
}

const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

// Поставить обрезку и поворот в XML рисунка (только если отличаются от того, что в нём уже есть)
export function setPicFormatXml(xml, fmt) {
  const cur = readPicFormat(xml);
  const want = { crop: fmt.crop || null, rotate: fmt.rotate || 0, flipH: !!fmt.flipH, flipV: !!fmt.flipV };
  let out = xml;
  if (!same(cur.crop, want.crop)) {
    out = out.replace(/<pic:blipFill\b[^>]*>[\s\S]*?<\/pic:blipFill>/, (fill) => {
      const f = fill.replace(/<a:srcRect\b[^>]*\/>|<a:srcRect\b[^>]*>[\s\S]*?<\/a:srcRect>/g, '');
      if (!want.crop) return f;
      const c = want.crop;
      const v = (k) => Math.round(c[k] * 100000);
      const tag = `<a:srcRect${['l', 't', 'r', 'b'].filter((k) => v(k)).map((k) => ` ${k}="${v(k)}"`).join('')}/>`;
      // по схеме CT_BlipFillProperties: blip, srcRect, (tile | stretch)
      return /<a:blip\b[^>]*\/>|<a:blip\b[\s\S]*?<\/a:blip>/.test(f) ? f.replace(/(<a:blip\b[^>]*\/>|<a:blip\b[\s\S]*?<\/a:blip>)/, `$1${tag}`) : f.replace(/(<pic:blipFill\b[^>]*>)/, `$1${tag}`);
    });
  }
  if (cur.rotate !== want.rotate && fmt.width && fmt.height) {
    // запас под повёрнутую картинку (EMU) — wp:effectExtent рамки
    const p = rotatedPad(fmt.width, fmt.height, want.rotate);
    const e = (v) => Math.round(v * 9525);
    out = out.replace(/<(\w+):effectExtent\b[^>]*\/>/, (all, ns) => `<${ns}:effectExtent l="${e(p.x)}" t="${e(p.y)}" r="${e(p.x)}" b="${e(p.y)}"/>`);
  }
  if (cur.rotate !== want.rotate || cur.flipH !== want.flipH || cur.flipV !== want.flipV) {
    out = out.replace(/(<pic:spPr\b[^>]*>[\s\S]*?)<a:xfrm\b([^>]*?)(\/?)>/, (all, head, attrs, slash) => {
      let a = attrs.replace(/\s*\b(rot|flipH|flipV)="[^"]*"/g, '');
      if (want.rotate) a += ` rot="${Math.round(want.rotate * 60000)}"`;
      if (want.flipH) a += ' flipH="1"';
      if (want.flipV) a += ' flipV="1"';
      return `${head}<a:xfrm${a}${slash}>`;
    });
  }
  return out;
}

// Запас вокруг рамки под повёрнутую картинку (габарит повёрнутого прямоугольника минус рамка, пополам), px —
// как wp:effectExtent у Word: место в строке и при обтекании резервируется по габариту
export function rotatedPad(w, h, deg) {
  const a = ((deg || 0) % 180) * Math.PI / 180;
  if (!a || !w || !h) return { x: 0, y: 0 };
  const c = Math.abs(Math.cos(a)); const s = Math.abs(Math.sin(a));
  return { x: Math.max(0, (w * c + h * s - w) / 2), y: Math.max(0, (w * s + h * c - h) / 2) };
}

// CSS поворота/отражения (сначала отражение, потом поворот — как DrawingML)
export const transformCss = ({ rotate, flipH, flipV }) => {
  const parts = [];
  if (rotate) parts.push(`rotate(${rotate}deg)`);
  if (flipH || flipV) parts.push(`scale(${flipH ? -1 : 1}, ${flipV ? -1 : 1})`);
  return parts.join(' ');
};
