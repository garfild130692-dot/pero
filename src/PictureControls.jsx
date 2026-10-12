// Контекстная вкладка «Рисунок»: размер (см, с сохранением пропорций) и замещающий текст.
import React, { useEffect, useState } from 'react';
import { fmtNum } from './editor/paragraph.js';
import { ArrangeGroup } from './ArrangeControls.jsx';

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
