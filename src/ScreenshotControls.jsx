// «Вставка → Снимок»: доступные окна (эскизы) — снимок окна целиком; «Вырезка экрана» — окно Перо сворачивается,
// снимается экран, затем мышью выделяется область (Esc — отмена), как в Word.
import React, { useEffect, useRef, useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { insertImageFiles } from './insertImage.js';
import * as fileApi from './fileApi.js';

const pngFile = (data, name = 'Снимок экрана.png') => ({ name, type: 'image/png', data });

export function ScreenshotButton({ editor, Btn, onClip }) {
  const pop = usePopover();
  const [list, setList] = useState(null);
  const open = async () => {
    if (pop.open) { pop.setOpen(false); return; }
    setList(null);
    pop.setOpen(true);
    setList(await fileApi.screenList());
  };
  const pick = async (w) => {
    pop.setOpen(false);
    const data = await fileApi.screenCapture(w.id);
    if (data) insertImageFiles(editor, [pngFile(data, `${w.name || 'Снимок экрана'}.png`)]);
  };
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big icon="screenshot" label="Снимок" title="Снимок окна или области экрана" active={pop.open} onClick={open} />
      {pop.open && (
        <div className="menu-pop shot-pop" style={pop.style} role="menu">
          <div className="pop-title">Доступные окна</div>
          {!list && <div className="recent-empty">Получаю список окон…</div>}
          {list && !list.length && <div className="recent-empty">Других открытых окон нет</div>}
          <div className="shot-grid">
            {(list || []).map((w) => (
              <button key={w.id} type="button" className="shot-cell" title={w.name} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(w)}>
                <img src={w.thumb} alt="" /><span>{w.name}</span>
              </button>
            ))}
          </div>
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onClip(); }}>
            <span className="check">✂</span>Вырезка экрана
          </button>
        </div>
      )}
    </div>
  );
}

// Вырезка: снимок экрана во весь экран окна, выделение прямоугольника, вставка вырезанной части в полном разрешении
export function ClipOverlay({ editor, data, onClose }) {
  // CSP окна разрешает картинки только из data:-URL (blob: запрещён)
  const [url, setUrl] = useState(null);
  const [rect, setRect] = useState(null);
  const imgRef = useRef(null);
  useEffect(() => {
    const r = new FileReader();
    r.onload = () => setUrl(r.result);
    r.readAsDataURL(new Blob([data], { type: 'image/png' }));
  }, [data]);
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);
  const down = (e) => {
    if (!imgRef.current) return;
    const b =imgRef.current.getBoundingClientRect();
    const x0 = e.clientX - b.left; const y0 = e.clientY - b.top;
    const move = (ev) => {
      const x1 = Math.min(Math.max(ev.clientX - b.left, 0), b.width); const y1 = Math.min(Math.max(ev.clientY - b.top, 0), b.height);
      setRect({ x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) });
    };
    const up = async (ev) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const x1 = Math.min(Math.max(ev.clientX - b.left, 0), b.width); const y1 = Math.min(Math.max(ev.clientY - b.top, 0), b.height);
      const r = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) };
      if (r.w < 4 || r.h < 4) { setRect(null); return; }
      // вырезать в пикселях исходного снимка
      const img = imgRef.current;
      const k = img.naturalWidth / b.width;
      const c = document.createElement('canvas');
      c.width = Math.round(r.w * k); c.height = Math.round(r.h * k);
      c.getContext('2d').drawImage(img, r.x * k, r.y * k, r.w * k, r.h * k, 0, 0, c.width, c.height);
      const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
      onClose();
      if (blob) insertImageFiles(editor, [pngFile(await blob.arrayBuffer(), 'Вырезка экрана.png')]);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  return (
    <div className="clip-overlay" onMouseDown={(e) => { e.preventDefault(); down(e); }}>
      {url && <img ref={imgRef} src={url} alt="Снимок экрана" draggable={false} />}
      <div className="clip-hint">Выделите область мышью · Esc — отмена</div>
      {rect && <div className="clip-rect" style={{ left: rect.x + imgRef.current.getBoundingClientRect().left, top: rect.y + imgRef.current.getBoundingClientRect().top, width: rect.w, height: rect.h }} />}
    </div>
  );
}
