// Заголовок окна как в Word/Excel: одна полоса — значок, панель быстрого доступа, имя файла по центру;
// справа — системные кнопки окна (Electron рисует их поверх полосы: titleBarStyle 'hidden' + titleBarOverlay).
// Полоса перетаскивает окно (двойной щелчок — развернуть), кнопки — нет.
import React from 'react';
import { Icon } from './icons.jsx';

export function QuickButton({ icon, title, onClick, disabled }) {
  const I = Icon[icon];
  return (
    <button type="button" className="rb-btn" title={title} disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      <I size={16} />
    </button>
  );
}

export default function TitleBar({ name, dirty, children }) {
  return (
    <div className="titlebar">
      <div className="quick-access">
        <span className="app-mark" aria-hidden="true">П</span>
        {children}
      </div>
      <div className="doc-title" title={name}>{name}{dirty ? ' •' : ''} — Перо</div>
      <div className="titlebar-right" />
    </div>
  );
}
