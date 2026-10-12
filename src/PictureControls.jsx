// Контекстная вкладка «Рисунок»: размер (см, с сохранением пропорций), обрезка и поворот, упорядочение, замещающий текст.
import React, { useEffect, useState } from 'react';
import { fmtNum } from './editor/paragraph.js';
import { ArrangeGroup } from './ArrangeControls.jsx';
import { usePopover } from './FontControls.jsx';

const PX_PER_CM = 96 / 2.54;

// Поле размера в сантиметрах: применяется по Enter или при уходе из поля
function CmField({ label, value, onCommit }) {
  const shown = value ? fmtNum(value / PX_PER_CM, 2) : '';
  const [text, setText] = useState(shown);
  useEffect(() => setText(shown), [shown]);
  const commit = () => {
    const v = parseFloat(text.replace(',', '.'));
    if (isFinite(v) && v > 0.1 && v <= 55) onCommit(v * PX_PER_CM);
    else setText(shown);
  };
  return (
    <div className="rb-field">
      <span>{label}</span>
      <span className="pd-input small">
        <input value={text} aria-label={`${label}, см`} onChange={(e) => setText(e.target.value)} onBlur={commit}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } if (e.key === 'Escape') setText(shown); }} />
        <em>см</em>
      </span>
    </div>
  );
}

// «Обрезка» (режим с маркерами на сторонах — NodeView рисунка), «Повернуть ▾», «Сбросить»
function CropRotateGroup({ editor, image, Group }) {
  const pop = usePopover();
  const view = () => { const dom = editor.view.nodeDOM(editor.state.selection.from); return dom && dom.peroView; };
  const rot = (d) => editor.chain().focus().setImageFormat({ rotate: (((image.rotate || 0) + d) % 360 + 360) % 360 }).run();
  const item = (label, fn) => (
    <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}><span className="check" />{label}</button>
  );
  // сбросить: вся картинка (размер всей картинки по текущей обрезке), без поворота и отражений
  const reset = () => {
    const c = image.crop || { l: 0, t: 0, r: 0, b: 0 };
    const r2 = (v) => Math.round(v * 100) / 100;
    editor.chain().focus().setImageFormat({
      crop: null, rotate: 0, flipH: false, flipV: false,
      width: r2(image.width / Math.max(0.01, 1 - c.l - c.r)), height: r2(image.height / Math.max(0.01, 1 - c.t - c.b)),
    }).run();
  };
  return (
    <Group title="Обрезка и поворот">
      <button type="button" className="rb-btn big" title="Обрезка: потяните маркеры на сторонах рисунка, Enter — применить, Esc — отменить"
        onMouseDown={(e) => e.preventDefault()} onClick={() => { const v = view(); if (v) v.toggleCrop(); }}>
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true"><path d="M6 2v16h16M2 6h16v16" /></svg>
        <span className="rb-label">Обрезка</span>
      </button>
      <div className="rb-popover-anchor" ref={pop.ref}>
        <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title="Повернуть или отразить рисунок"
          onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true"><path d="M20 12a8 8 0 1 1-2.3-5.7" /><path d="M20 3v4h-4" /></svg>
          <span className="rb-label">Повернуть ▾</span>
        </button>
        {pop.open && (
          <div className="menu-pop" style={pop.style} role="menu">
            {item('Повернуть вправо на 90°', () => rot(90))}
            {item('Повернуть влево на 90°', () => rot(-90))}
            {item('Отразить сверху вниз', () => editor.chain().focus().setImageFormat({ flipV: !image.flipV }).run())}
            {item('Отразить слева направо', () => editor.chain().focus().setImageFormat({ flipH: !image.flipH }).run())}
            <div className="menu-sep" />
            {item('Без поворота', () => editor.chain().focus().setImageFormat({ rotate: 0, flipH: false, flipV: false }).run())}
          </div>
        )}
      </div>
      <div className="rb-rows">
        <RotField value={image.rotate || 0} onCommit={(d) => editor.chain().focus().setImageFormat({ rotate: d }).run()} />
        <button type="button" className="rb-btn small" title="Вернуть исходный рисунок: без обрезки и поворота" onMouseDown={(e) => e.preventDefault()} onClick={reset}>Сбросить</button>
      </div>
    </Group>
  );
}

// Угол поворота, градусы
function RotField({ value, onCommit }) {
  const shown = fmtNum(value, 1);
  const [text, setText] = useState(shown);
  useEffect(() => setText(shown), [shown]);
  const commit = () => {
    const v = parseFloat(String(text).replace(',', '.'));
    if (isFinite(v)) onCommit(((Math.round(v * 100) / 100) % 360 + 360) % 360); else setText(shown);
  };
  return (
    <div className="rb-field">
      <span>Поворот</span>
      <span className="pd-input small">
        <input value={text} aria-label="Поворот, градусы" onChange={(e) => setText(e.target.value)} onBlur={commit}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } if (e.key === 'Escape') setText(shown); }} />
        <em>°</em>
      </span>
    </div>
  );
}

export function PicturePanel({ editor, image, Group }) {
  const [alt, setAlt] = useState(image.alt || '');
  useEffect(() => setAlt(image.alt || ''), [image.alt]);
  const ratio = image.width && image.height ? image.height / image.width : 1;
  const r2 = (v) => Math.round(v * 100) / 100;
  const setSize = (w, h) => editor.chain().focus().setImageSize(r2(w), r2(h)).run();
  return (
    <>
      <Group title="Размер">
        <div className="rb-rows">
          <CmField label="Высота" value={image.height} onCommit={(h) => setSize(h / ratio, h)} />
          <CmField label="Ширина" value={image.width} onCommit={(w) => setSize(w, w * ratio)} />
        </div>
      </Group>
      <CropRotateGroup editor={editor} image={image} Group={Group} />
      <ArrangeGroup editor={editor} image={image} Group={Group} />
      <Group title="Замещающий текст">
        <div className="rb-field alt-field">
          <input value={alt} placeholder="Описание рисунка" aria-label="Замещающий текст"
            title="Текст для программ чтения с экрана и на случай, если рисунок не загрузится"
            onChange={(e) => setAlt(e.target.value)}
            onBlur={() => alt !== (image.alt || '') && editor.chain().setImageAlt(alt).run()}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.chain().focus().setImageAlt(alt).run(); } }} />
        </div>
      </Group>
    </>
  );
}
