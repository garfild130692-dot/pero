// Группа «Абзац»: границы и заливка абзаца (как кнопки Word «Границы» и «Заливка»)
import React, { useState } from 'react';
import { ColorPalette, SplitColor, usePopover } from './FontControls.jsx';
import { BorderIcon, BucketIcon } from './TableControls.jsx';

// [режим команды setParagraphBorders, подпись, значок]
const MODES = [
  ['bottom', 'Нижняя граница', 'bottom'], ['top', 'Верхняя граница', 'top'], ['left', 'Левая граница', 'left'], ['right', 'Правая граница', 'right'],
  null,
  ['none', 'Нет границы', 'none'], ['all', 'Все границы', 'paraAll'], ['outer', 'Внешние границы', 'outer'],
  ['between', 'Граница между абзацами', 'between'],
];

export function ParaBordersButton({ editor }) {
  const pop = usePopover();
  const [last, setLast] = useState('bottom');
  const apply = (m) => { setLast(m); editor.chain().focus().setParagraphBorders(m).run(); };
  const icon = (MODES.find((x) => x && x[0] === last) || [])[2] || 'bottom';
  const label = (MODES.find((x) => x && x[0] === last) || [])[1];
  return (
    <div className="split" ref={pop.ref}>
      <button type="button" className="rb-btn split-main" title={`Границы абзаца: ${label}`} onMouseDown={(e) => e.preventDefault()} onClick={() => apply(last)}>
        <BorderIcon mode={icon} />
      </button>
      <button type="button" className="rb-btn split-arrow" title="Границы абзаца: выбрать" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop borders-pop" style={pop.style} role="menu">
          {MODES.map((m, i) => (m ? (
            <button key={m[0]} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); apply(m[0]); }}>
              <BorderIcon mode={m[2]} size={16} />{m[1]}
            </button>
          ) : <div key={i} className="menu-sep" />))}
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); editor.chain().focus().insertHorizontalLine().run(); }}>
            <BorderIcon mode="bottom" size={16} />Горизонтальная линия
          </button>
        </div>
      )}
    </div>
  );
}

export function ParaShadingButton({ editor }) {
  const [last, setLast] = useState('#dcefe9');
  const apply = (c) => { setLast(c); editor.chain().focus().setParagraphShading(c).run(); };
  return (
    <SplitColor title="Заливка абзаца" color={last} icon={<BucketIcon />}>
      {{
        apply: () => apply(last),
        palette: (close) => (
          <ColorPalette autoLabel="Нет цвета" autoClass="none" current={last}
            onAuto={() => { close(); editor.chain().focus().setParagraphShading(null).run(); }}
            onPick={(x) => { close(); apply(x); }} />
        ),
      }}
    </SplitColor>
  );
}
