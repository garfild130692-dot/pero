// Рисунки: узел редактора, изменение размера мышью, общие преобразования для импорта/экспорта.
//
// Узел image (встроенный в строку): src — data:-URL картинки (для показа и для записи новых рисунков),
// width/height — размер на странице в px (1 px = 9525 EMU), alt — замещающий текст,
// xml/target — исходный <w:r> с <w:drawing> и путь картинки в пакете (у рисунков из файла):
// при сохранении в тот же пакет пишется исходный XML с новым размером — сохраняются обтекание, обрезка, эффекты;
// anchor — обтекание и положение плавающего рисунка (anchor.js; null — в строке текста), показ — floatLayout.js.
import { Node, mergeAttributes } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';

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
    this.dom.append(this.img, this.handle);
    this.render();
  }

  render() {
    const { src, width, height, alt, anchor } = this.node.attrs;
    if (this.img.getAttribute('src') !== src) this.img.src = src || '';
    this.img.alt = alt || '';
    this.img.style.width = width ? `${width}px` : '';
    this.img.style.height = height ? `${height}px` : '';
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
    const w0 = this.img.offsetWidth;
    const h0 = this.img.offsetHeight;
    // offsetWidth — без учёта масштаба страницы (CSS zoom), смещение мыши — с ним
    const scale = this.img.getBoundingClientRect().width / (w0 || 1);
    let w = w0;
    let h = h0;
    const move = (ev) => {
      const dx = (ev.clientX - startX) / scale;
      const dy = (ev.clientY - startY) / scale;
      w = Math.max(16, w0 + dx);
      h = ev.shiftKey ? Math.max(16, h0 + dy) : (w * h0) / (w0 || 1);
      this.img.style.width = `${w}px`;
      this.img.style.height = `${h}px`;
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
