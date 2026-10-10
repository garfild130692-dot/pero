import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';
import SheetApp from './sheet/SheetApp.jsx';
import RecoveryPanel from './RecoveryPanel.jsx';
import * as fileApi from './fileApi.js';
import './styles.css';
import './sheet/sheet.css';

// Ошибки интерфейса — в журнал программы (%APPDATA%\Перо\logs), чтобы их было видно в «Сообщить о проблеме»
window.addEventListener('error', (e) => fileApi.logError(`${e.message} (${e.filename}:${e.lineno})${e.error && e.error.stack ? `\n${e.error.stack}` : ''}`));
window.addEventListener('unhandledrejection', (e) => fileApi.logError(`Необработанный отказ: ${e.reason && e.reason.stack ? e.reason.stack : e.reason}`));

// Одно окно — один файл: текстовый документ (.docx) или книга (.xlsx). Root переключает редакторы.
const isSheet = (f) => f && /\.(xlsx|csv)$/i.test(f.name);

function Root() {
  const [boot, setBoot] = useState(null); // { autosaveSec, recovery, notice }
  const [mode, setMode] = useState(null);
  const open = (kind, file) => setMode((m) => ({ kind, file, n: (m ? m.n : 0) + 1 }));
  const openOther = (file) => open(isSheet(file) ? 'sheet' : 'doc', file);

  useEffect(() => {
    fileApi.ready().then((r) => {
      setBoot({ autosaveSec: r.autosaveSec || 0, recovery: r.recovery, notice: r.notice });
      if (r.file) openOther(r.file); else open(r.kind === 'sheet' ? 'sheet' : 'doc', null);
    }, (e) => { fileApi.logError(e); open('doc', null); });
  }, []);

  if (!mode) return null;
  const props = { initialFile: mode.file, onOpenOther: openOther, autosaveSec: boot.autosaveSec };
  return (
    <>
      {mode.kind === 'sheet'
        ? <SheetApp key={mode.n} {...props} onNewDoc={() => open('doc', null)} />
        : <App key={mode.n} {...props} onNewSheet={() => open('sheet', null)} />}
      {boot.recovery && (
        <RecoveryPanel items={boot.recovery} onOpenHere={openOther} onClose={() => setBoot((b) => ({ ...b, recovery: null }))} />
      )}
      {boot.notice && (
        <div className="notice-bar" role="status">
          <span>{boot.notice}</span>
          <button type="button" onClick={() => setBoot((b) => ({ ...b, notice: null }))} title="Закрыть">×</button>
        </div>
      )}
    </>
  );
}

// Ошибка в интерфейсе не должна оставлять белое окно: показываем, что случилось, и даём перезапустить окно —
// документ вернётся из автосохранения (если оно успело сделаться).
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { fileApi.logError(`Ошибка интерфейса: ${error && error.stack ? error.stack : error}\n${info.componentStack || ''}`); }
  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div className="crash-screen">
        <h1>В окне Перо произошла ошибка</h1>
        <p>Сведения записаны в журнал. Перезапустите окно — несохранённые изменения восстановятся из автосохранения
          (оно делается раз в минуту).</p>
        <pre>{String(this.state.error && this.state.error.message)}</pre>
        <div className="crash-actions">
          <button type="button" className="btn primary" onClick={() => window.location.reload()}>Перезапустить окно</button>
          <button type="button" className="btn" onClick={() => fileApi.reportProblem()}>Сообщить о проблеме</button>
        </div>
      </div>
    );
  }
}

createRoot(document.getElementById('root')).render(<ErrorBoundary><Root /></ErrorBoundary>);
