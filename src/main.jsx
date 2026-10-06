import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';
import SheetApp from './sheet/SheetApp.jsx';
import './styles.css';
import './sheet/sheet.css';

// Одно окно — один файл: текстовый документ (.docx) или книга (.xlsx). Root переключает редакторы.
const isSheet = (f) => f && /\.xlsx$/i.test(f.name);

function Root() {
  const [mode, setMode] = useState({ kind: 'doc', file: null, n: 0 });
  const openDoc = (file) => setMode((m) => ({ kind: 'doc', file, n: m.n + 1 }));
  const openSheet = (file) => setMode((m) => ({ kind: 'sheet', file, n: m.n + 1 }));
  // в браузерном режиме отладки: ?sheet — сразу книга
  useEffect(() => { if (/[?&]sheet\b/.test(window.location.search)) openSheet(null); }, []);
  const openOther = (file) => (isSheet(file) ? openSheet(file) : openDoc(file));
  return mode.kind === 'sheet'
    ? <SheetApp key={mode.n} initialFile={mode.file} onOpenOther={openOther} onNewDoc={() => openDoc(null)} />
    : <App key={mode.n} initialFile={mode.file} onOpenOther={openOther} onNewSheet={() => openSheet(null)} />;
}

createRoot(document.getElementById('root')).render(<Root />);
