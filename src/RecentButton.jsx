// Кнопка «Последние» на вкладке «Файл»: список недавно открытых и сохранённых файлов (общий для всех окон).
import React, { useState } from 'react';
import { Icon } from './icons.jsx';
import { usePopover } from './FontControls.jsx';
import * as fileApi from './fileApi.js';

const dirOf = (p) => p.replace(/[\\/][^\\/]*$/, '');

// onOpen(file) — открыть в этом окне (программа вернула данные, т.к. окно пустое); иначе файл откроется в новом окне
export default function RecentButton({ onOpen }) {
  const pop = usePopover();
  const [items, setItems] = useState([]);
  const toggle = async () => {
    if (!pop.open) setItems(await fileApi.recentList());
    pop.setOpen(!pop.open);
  };
  const open = async (p) => {
    pop.setOpen(false);
    const f = await fileApi.openPath(p);
    if (f) onOpen(f);
  };
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title="Недавно открытые файлы"
        onMouseDown={(e) => e.preventDefault()} onClick={toggle}>
        <Icon.recent size={28} />
        <span className="rb-label">Последние</span>
      </button>
      {pop.open && (
        <div className="menu-pop recent-pop" style={pop.style} role="menu">
          {!items.length && <div className="recent-empty">Список пуст</div>}
          {items.map((it) => (
            <button key={it.path} type="button" className="recent-item" title={it.path} onClick={() => open(it.path)}>
              <span className="recent-name">{it.name}</span>
              <span className="recent-dir">{dirOf(it.path)}</span>
            </button>
          ))}
          {items.length > 0 && (
            <>
              <div className="menu-sep" />
              <button type="button" className="menu-item" onClick={async () => { await fileApi.recentClear(); setItems([]); }}>
                <span className="check" />Очистить список
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
}
