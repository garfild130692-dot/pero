// Рисунки: узел редактора, изменение размера мышью, общие преобразования для импорта/экспорта.
//
// Узел image (встроенный в строку): src — data:-URL картинки (для показа и для записи новых рисунков),
// width/height — размер на странице в px (1 px = 9525 EMU), alt — замещающий текст,
// xml/target — исходный <w:r> с <w:drawing> и путь картинки в пакете (у рисунков из файла):
// при сохранении в тот же пакет пишется исходный XML с новым размером — сохраняются обтекание, обрезка, эффекты;
// anchor — обтекание и положение плавающего рисунка (anchor.js; null — в строке текста), показ — floatLayout.js;
// crop/rotate/flipH/flipV — обрезка и поворот (picFormat.js): рамка width×height показывает часть картинки.
import { Node, mergeAttributes } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import { rotatedPad, transformCss } from './picFormat.js';

export const EMU_PER_PX = 9525;
export const MAX_INSERT_WIDTH = 623; // ширина текста A4 с полями по умолчанию (16,5 см) в px

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', webp: 'image/webp' };
export const mimeOf = (path) => MIME[String(path).split('.').pop().toLowerCase()] || 'application/octet-stream';
// Форматы, которые Word понимает без преобразования
export const WORD_FORMATS = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp'];
export const extOf = (mime) => ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' })[mime] || 'png';

// base64 ↔ байты без Buffer (работает и в браузере, и в Node)
export function bytesToBase64(bytes) {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(bin);
}
export function dataUrlToBytes(url) {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url || '');
  if (!m) return null;
  const bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return { mime: m[1] || 'application/octet-stream', bytes: out };
}

// Изменение размера: рамка с маркером в правом нижнем углу, пропорции сохраняются (Shift — свободно)
class ImageView {
  constructor(node, view, getPos) {
    this.node = node;
    this.view = view;
    this.getPos = getPos;
    this.dom = document.createElement('span');
    this.dom.className = 'pero-image';
    this.img = document.createElement('img');
    this.img.draggable = false;
    this.handle = document.createElement('span');
    this.handle.className = 'pero-image-handle';
    this.handle.title = 'Потяните, чтобы изменить размер (Shift — без сохранения пропорций)';
    this.handle.addEventListener('mousedown', (e) => this.startResize(e));
    this.img.addEventListener('mousedown', (e) => this.startMove(e));
    // у плавающего — своё перетаскивание; родное (перенос узла в тексте) не запускать: ProseMirror сам ставит draggable
    this.dom.addEventListener('dragstart', (e) => { if (this.node.attrs.anchor) e.preventDefault(); });
    // рамка: обрезка (overflow: hidden, картинка сдвинута и увеличена) и поворот; в режиме обрезки — вся картинка бледно
    this.frame = document.createElement('span');
    this.frame.className = 'pero-frame';
    this.ghost = null;
    this.frame.append(this.img);
    this.dom.append(this.frame, this.handle);
    this.dom.peroView = this; // для кнопки «Обрезка» на ленте
    this.cropping = null;
    this.render();
  }

  // Размер всей картинки и её сдвиг в рамке по обрезке
  geom(crop = this.cropping || this.node.attrs.crop) {
    const { width, height } = this.node.attrs;
    const base = this.node.attrs.crop || { l: 0, t: 0, r: 0, b: 0 };
    const w = width || this.img.naturalWidth || 96;
    const h = height || this.img.naturalHeight || 96;
    // вся картинка — по текущей (сохранённой) обрезке, рамка — по показываемой
    const W = w / Math.max(0.01, 1 - base.l - base.r);
    const H = h / Math.max(0.01, 1 - base.t - base.b);
    const c = crop || { l: 0, t: 0, r: 0, b: 0 };
    return { W, H, fw: W * (1 - c.l - c.r), fh: H * (1 - c.t - c.b), ox: -c.l * W, oy: -c.t * H };
  }

  render() {
    const { src, width, height, alt, anchor } = this.node.attrs;
    if (this.img.getAttribute('src') !== src) this.img.src = src || '';
    this.img.alt = alt || '';
    const g = this.geom();
    const cropped = !!(this.node.attrs.crop || this.cropping);
    this.frame.style.width = width ? `${g.fw}px` : '';
    this.frame.style.height = height ? `${g.fh}px` : '';
    this.img.style.width = width ? `${cropped ? g.W : width}px` : '';
    this.img.style.height = height ? `${cropped ? g.H : height}px` : '';
    this.img.style.marginLeft = cropped ? `${g.ox}px` : '';
    this.img.style.marginTop = cropped ? `${g.oy}px` : '';
    this.frame.style.transform = transformCss(this.node.attrs);
    // повёрнутая картинка занимает свой габарит (как wp:effectExtent у Word)
    const pad = rotatedPad(g.fw, g.fh, this.node.attrs.rotate);
    this.frame.style.margin = pad.x || pad.y ? `${pad.y.toFixed(1)}px ${pad.x.toFixed(1)}px` : '';
    // плавающий: класс по виду обтекания, положение ставит floatLayout (стили — прямо в DOM, ignoreMutation)
    const kind = !anchor ? '' : anchor.wrap === 'none' ? (anchor.behind ? 'behind' : 'front') : anchor.wrap;
    this.dom.className = `pero-image${kind ? ` pero-float pf-${kind}` : ''}${this.dom.classList.contains('selected') ? ' selected' : ''}`;
    if (!kind) { this.dom.removeAttribute('style'); delete this.dom.dataset.floatCss; }
  }

  startResize(e) {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startY = e.clientY;
    // размер рамки (при обрезке картинка внутри больше рамки — масштабируется вместе с ней)
    const w0 = this.frame.offsetWidth;
    const h0 = this.frame.offsetHeight;
    const iw0 = this.img.offsetWidth; const ih0 = this.img.offsetHeight;
    const ml0 = parseFloat(this.img.style.marginLeft) || 0; const mt0 = parseFloat(this.img.style.marginTop) || 0;
    // offsetWidth — без учёта масштаба страницы (CSS zoom), смещение мыши — с ним
    // масштаб страницы — по узлу рисунка (он не повёрнут, в отличие от рамки)
    const scale = this.dom.getBoundingClientRect().width / (this.dom.offsetWidth || 1) || 1;
    let w = w0;
    let h = h0;
    const move = (ev) => {
      const dx = (ev.clientX - startX) / scale;
      const dy = (ev.clientY - startY) / scale;
      w = Math.max(16, w0 + dx);
      h = ev.shiftKey ? Math.max(16, h0 + dy) : (w * h0) / (w0 || 1);
      const kx = w / (w0 || 1); const ky = h / (h0 || 1);
      this.frame.style.width = `${w}px`;
      this.frame.style.height = `${h}px`;
      this.img.style.width = `${iw0 * kx}px`;
      this.img.style.height = `${ih0 * ky}px`;
      this.img.style.marginLeft = `${ml0 * kx}px`;
      this.img.style.marginTop = `${mt0 * ky}px`;
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const pos = this.getPos();
      if (typeof pos !== 'number') return;
      const r = (v) => Math.round(v * 100) / 100;
      this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, width: r(w), height: r(h) }));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  // Плавающий рисунок: тянуть мышью — новое место от колонки и абзаца (как Word при перетаскивании)
  startMove(e) {
    const a = this.node.attrs.anchor;
    if (!a || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const pos = this.getPos();
    if (typeof pos !== 'number') return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
    const scale = this.dom.getBoundingClientRect().width / (this.dom.offsetWidth || 1) || 1;
    const sx = e.clientX; const sy = e.clientY;
    let dx = 0; let dy = 0;
    const move = (ev) => {
      dx = (ev.clientX - sx) / scale; dy = (ev.clientY - sy) / scale;
      this.dom.style.translate = `${dx}px ${dy}px`;
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      this.dom.style.translate = '';
      if (Math.abs(dx) < 2 && Math.abs(dy) < 2) return;
      const p = this.getPos();
      if (typeof p !== 'number') return;
      const r = (v) => Math.round(v * 100) / 100;
      const relx = parseFloat(this.dom.dataset.relx) || 0;
      const rely = parseFloat(this.dom.dataset.rely) || 0;
      const anchor = { ...a, h: { rel: 'column', align: null, off: r(relx + dx) }, v: { rel: 'paragraph', align: null, off: r(rely + dy) } };
      this.view.dispatch(this.view.state.tr.setNodeMarkup(p, undefined, { ...this.node.attrs, anchor }).setMeta('floatLayout$', { pending: p }));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  // Режим обрезки: вся картинка бледно, рамка — сохраняемая часть, маркеры на сторонах тянут края; Enter/щелчок мимо —
  // применить, Esc — отменить. Размер рамки после обрезки = вся картинка × (1 − обрезанное).
  toggleCrop() {
    if (this.cropping) { this.finishCrop(true); return; }
    this.cropping = { ...(this.node.attrs.crop || { l: 0, t: 0, r: 0, b: 0 }) };
    this.dom.classList.add('cropping');
    this.ghost = this.img.cloneNode();
    this.ghost.className = 'pero-crop-ghost';
    this.frame.before(this.ghost);
    this.cropHandles = ['l', 't', 'r', 'b'].map((side) => {
      const hd = document.createElement('span');
      hd.className = `pero-crop-handle ch-${side}`;
      hd.addEventListener('mousedown', (e) => this.dragCrop(e, side));
      this.dom.append(hd);
      return hd;
    });
    this.onKey = (e) => { if (e.key === 'Enter') { e.preventDefault(); this.finishCrop(true); } else if (e.key === 'Escape') this.finishCrop(false); };
    this.onDown = (e) => { if (!this.dom.contains(e.target)) this.finishCrop(true); };
    window.addEventListener('keydown', this.onKey, true);
    window.addEventListener('mousedown', this.onDown, true);
    this.renderCrop();
  }

  renderCrop() {
    const g = this.geom();
    // бледная вся картинка — от левого верхнего угла рамки назад на обрезанное
    this.ghost.style.cssText = `position:absolute;left:${g.ox}px;top:${g.oy}px;width:${g.W}px;height:${g.H}px;opacity:.35;pointer-events:none`;
    this.render();
    const pos = { l: [0, g.fh / 2], r: [g.fw, g.fh / 2], t: [g.fw / 2, 0], b: [g.fw / 2, g.fh] };
    for (const hd of this.cropHandles) {
      const side = hd.className.slice(-1);
      hd.style.left = `${pos[side][0]}px`;
      hd.style.top = `${pos[side][1]}px`;
    }
  }

  dragCrop(e, side) {
    e.preventDefault();
    e.stopPropagation();
    const g = this.geom();
    const scale = this.dom.getBoundingClientRect().width / (this.dom.offsetWidth || 1) || 1;
    const start = { ...this.cropping };
    const sx = e.clientX; const sy = e.clientY;
    const move = (ev) => {
      const dx = (ev.clientX - sx) / scale / g.W;
      const dy = (ev.clientY - sy) / scale / g.H;
      const c = { ...start };
      if (side === 'l') c.l = Math.min(Math.max(0, start.l + dx), 0.95 - c.r);
      if (side === 'r') c.r = Math.min(Math.max(0, start.r - dx), 0.95 - c.l);
      if (side === 't') c.t = Math.min(Math.max(0, start.t + dy), 0.95 - c.b);
      if (side === 'b') c.b = Math.min(Math.max(0, start.b - dy), 0.95 - c.t);
      this.cropping = c;
      this.renderCrop();
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  finishCrop(apply) {
    const c = this.cropping;
    const g = this.geom();
    this.cropping = null;
    window.removeEventListener('keydown', this.onKey, true);
    window.removeEventListener('mousedown', this.onDown, true);
    this.dom.classList.remove('cropping');
    if (this.ghost) this.ghost.remove();
    (this.cropHandles || []).forEach((hd) => hd.remove());
    this.ghost = null; this.cropHandles = null;
    const pos = this.getPos();
    if (!apply || typeof pos !== 'number') { this.render(); return; }
    const r = (v) => Math.round(v * 100000) / 100000;
    const crop = c.l || c.t || c.r || c.b ? { l: r(c.l), t: r(c.t), r: r(c.r), b: r(c.b) } : null;
    const r2 = (v) => Math.round(v * 100) / 100;
    this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, crop, width: r2(g.fw), height: r2(g.fh) }));
  }

  update(node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() { this.dom.classList.add('selected'); }
  deselectNode() { this.dom.classList.remove('selected'); }
  stopEvent(e) { return e.target === this.handle || (!!this.node.attrs.anchor && e.target === this.img && e.type === 'mousedown'); }
  ignoreMutation() { return true; }
}

export const PeroImage = Node.create({
  name: 'image',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return {
      src: { default: null },
      width: { default: null },
      height: { default: null },
      alt: { default: '' },
      xml: { default: null, rendered: false },
      target: { default: null, rendered: false },
      anchor: { default: null, rendered: false },
      crop: { default: null, rendered: false },
      rotate: { default: 0, rendered: false },
      flipH: { default: false, rendered: false },
      flipV: { default: false, rendered: false },
    };
  },

  parseHTML() {
    return [{
      tag: 'img[src]',
      getAttrs: (el) => ({
        src: el.getAttribute('src'),
        alt: el.getAttribute('alt') || '',
        width: parseFloat(el.style.width || el.getAttribute('width')) || null,
        height: parseFloat(el.style.height || el.getAttribute('height')) || null,
      }),
    }];
  },

  renderHTML({ HTMLAttributes }) {
    const { width, height, ...rest } = HTMLAttributes;
    const style = [width ? `width: ${width}px` : '', height ? `height: ${height}px` : ''].filter(Boolean).join('; ');
    return ['img', mergeAttributes(rest, style ? { style } : {})];
  },

  addNodeView() {
    return ({ node, view, getPos }) => new ImageView(node, view, getPos);
  },

  addCommands() {
    return {
      setImageSize: (width, height) => ({ state, tr, dispatch }) => {
        const sel = state.selection;
        if (!sel.node || sel.node.type.name !== 'image') return false;
        if (dispatch) dispatch(tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, width, height }));
        return true;
      },
      // обтекание и положение (anchor.js); null — в строке текста
      setImageAnchor: (anchor) => ({ state, tr, dispatch }) => {
        const sel = state.selection;
        if (!sel.node || sel.node.type.name !== 'image') return false;
        if (dispatch) dispatch(tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, anchor }).setMeta('floatLayout$', { pending: sel.from }));
        return true;
      },
      // поворот/отражение: { rotate, flipH, flipV } — заданные поля; сброс обрезки — crop: null с прежним размером всей картинки
      setImageFormat: (fmt) => ({ state, tr, dispatch }) => {
        const sel = state.selection;
        if (!sel.node || sel.node.type.name !== 'image') return false;
        if (dispatch) dispatch(tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, ...fmt }));
        return true;
      },
      setImageAlt: (alt) => ({ state, tr, dispatch }) => {
        const sel = state.selection;
        if (!sel.node || sel.node.type.name !== 'image') return false;
        if (dispatch) dispatch(tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, alt }));
        return true;
      },
    };
  },
});

// Подготовить картинку к вставке: Word-совместимый формат (иначе — PNG через canvas) и размер не шире текста
export async function prepareImage(bytes, mime, maxWidth = MAX_INSERT_WIDTH) {
  // грузим из data:-URL — CSP окна разрешает картинки только из self и data: (blob: — нет)
  const dataUrl = `data:${mime};base64,${bytesToBase64(new Uint8Array(bytes))}`;
  const img = await new Promise((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = () => reject(new Error('Не удалось прочитать изображение'));
    i.src = dataUrl;
  });
  let src = dataUrl;
  if (!WORD_FORMATS.includes(mime)) {
    const c = document.createElement('canvas');
    c.width = img.naturalWidth || 300;
    c.height = img.naturalHeight || 150;
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    src = c.toDataURL('image/png');
  }
  const w0 = img.naturalWidth || 300;
  const h0 = img.naturalHeight || 150;
  const k = w0 > maxWidth ? maxWidth / w0 : 1;
  return { src, width: Math.round(w0 * k * 100) / 100, height: Math.round(h0 * k * 100) / 100 };
}
