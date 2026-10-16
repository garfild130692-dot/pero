// Начальный экран (запуск без файла, как в Word/Excel): создать документ или книгу, восстановить несохранённое после
// сбоя, открыть последний файл (поиск, закрепление, удаление из списка) или другой файл. Выбор открывается в этом окне.
import React, { useEffect, useMemo, useState } from 'react';
import TitleBar from './TitleBar.jsx';
import * as fileApi from './fileApi.js';

const dirOf = (p) => p.replace(/[\\/][^\\/]*$/, '');
const kindOf = (name) => (/\.(xlsx|xls|csv)$/i.test(name) ? 'sheet' : 'doc');

// «Сегодня, 14:05», «Вчера, 09:12», иначе дата
export function fmtWhen(t, now = new Date()) {
  if (!t) return '';
  const d = new Date(t);
  const hm = d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  const day = (x) => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const diff = Math.round((day(now) - day(d)) / 86400000);
  if (diff === 0) return `Сегодня, ${hm}`;
  if (diff === 1) return `Вчера, ${hm}`;
  return d.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// Приветствие по времени суток
export function greeting(h = new Date().getHours()) {
  if (h >= 5 && h < 12) return 'Доброе утро';
  if (h >= 12 && h < 18) return 'Добрый день';
  if (h >= 18 && h < 23) return 'Добрый вечер';
  return 'Доброй ночи';
}

// Свои значки файлов (не эмблемы Office)
const FileIcon = ({ kind, size = 22 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" aria-hidden="true"
    className={`start-ficon ${kind}`}>
    <path d="M6 2h8l5 5v15H6z" /><path d="M14 2v5h5" />
    {kind === 'sheet'
      ? <path d="M8.5 11h8M8.5 14.5h8M8.5 18h8M12.5 11v7" />
      : <path d="M8.5 11h8M8.5 14h8M8.5 17h5" />}
  </svg>
);
const PinIcon = ({ on }) => (
  <svg width="15" height="15" viewBox="0 0 24 24" fill={on ? 'currentColor' : 'none'} stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" aria-hidden="true">
    <path d="M9 3h6l-1 6 4 4H6l4-4z" /><path d="M12 13v8" />
  </svg>
);

// Эскизы плиток «Создать»
const DocThumb = () => (
  <div className="start-thumb doc">
    <div className="tl w60" /><div className="tl" /><div className="tl" /><div className="tl w80" /><div className="tl" /><div className="tl w40" />
  </div>
);
const SheetThumb = () => (
  <div className="start-thumb sheet">
    {Array.from({ length: 24 }, (_, i) => <span key={i} className={i < 4 ? 'hd' : ''} />)}
  </div>
);

export default function StartScreen({ onCreate, onOpen }) {
  const [recent, setRecent] = useState(null);
  const [recovery, setRecovery] = useState([]);
  const [query, setQuery] = useState('');
  const [showAtStart, setShowAtStart] = useState(true);
  const [busy, setBusy] = useState(false);

  const reload = async () => {
    setRecent(await fileApi.recentList(true));
    setRecovery(await fileApi.recoveryList());
  };
  useEffect(() => {
    fileApi.setTitle('Перо');
    fileApi.setDoc({ path: null, kind: 'start' });
    reload();
    fileApi.settingsGet().then((s) => setShowAtStart(s.startScreen !== false));
  }, []);

  // открыть файл (данные пришли этому окну) — в нём же
  const here = (f) => { if (f) onOpen(f); };
  const run = async (fn) => { setBusy(true); try { await fn(); } finally { setBusy(false); } };

  const openRecent = (it) => run(async () => {
    if (it.exists === false) {
      if (window.confirm(`Файл не найден:\n${it.path}\n\nУбрать его из списка последних?`)) { await fileApi.recentRemove(it.path); reload(); }
      return;
    }
    here(await fileApi.openPath(it.path));
    reload(); // файл мог открыться в другом окне (уже открыт) — список обновится
  });
  const openOther = () => run(async () => here(await fileApi.openFile()));
  const restore = (it) => run(async () => here(await fileApi.recoveryOpen(it.id)));
  const discard = async (it) => {
    if (!window.confirm(`Удалить несохранённую копию «${it.name}»? Восстановить её будет нельзя.`)) return;
    await fileApi.recoveryDiscard(it.id);
    reload();
  };

  // Ctrl+N (новый документ здесь), Ctrl+O, файл, присланный программой (двойной щелчок по файлу в Проводнике)
  useEffect(() => fileApi.onMenu((cmd, payload) => {
    if (cmd === 'newHere' || cmd === 'newDoc') onCreate('doc');
    else if (cmd === 'newSheet') onCreate('sheet');
    else if (cmd === 'openData') onOpen(payload);
    else if (cmd === 'open') openOther();
  }), []);

  const onDrop = async (e) => {
    const f = [...e.dataTransfer.files].find((x) => /\.(docx|doc|xlsx|xls|csv)$/i.test(x.name));
    if (!f) return;
    e.preventDefault();
    const p = fileApi.pathForFile(f);
    if (p) here(await fileApi.openPath(p));
    else here({ path: null, name: f.name, data: await f.arrayBuffer() });
  };

  const list = useMemo(() => {
    const q = query.trim().toLowerCase();
    return (recent || []).filter((it) => !q || it.name.toLowerCase().includes(q) || it.path.toLowerCase().includes(q));
  }, [recent, query]);

  return (
    <div className="app start-app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <TitleBar name={null} />
      <div className="start">
        <nav className="start-nav" aria-label="Начальный экран">
          <div className="start-brand">Перо</div>
          <button type="button" className="start-nav-item active">Главная</button>
          <button type="button" className="start-nav-item" onClick={openOther} disabled={busy}>Открыть… <kbd>Ctrl+O</kbd></button>
          <div className="spacer" />
          <label className="start-opt" title="При запуске без файла показывать этот экран (иначе сразу новый документ)">
            <input type="checkbox" checked={showAtStart} onChange={(e) => { setShowAtStart(e.target.checked); fileApi.settingsSet('startScreen', e.target.checked); }} />
            Показывать при запуске
          </label>
          <button type="button" className="start-nav-item small" onClick={() => fileApi.about()}>О программе</button>
        </nav>
        <main className="start-main">
          <h1>{greeting()}</h1>

          <section className="start-section">
            <h2>Создать</h2>
            <div className="start-tiles">
              <button type="button" className="start-tile" onClick={() => onCreate('doc')} title="Новый текстовый документ (Ctrl+N)">
                <DocThumb /><span>Новый документ</span>
              </button>
              <button type="button" className="start-tile" onClick={() => onCreate('sheet')} title="Новая книга с листами и формулами">
                <SheetThumb /><span>Новая книга</span>
              </button>
            </div>
          </section>

          {recovery.length > 0 && (
            <section className="start-section start-recovery">
              <h2>Восстановление документов</h2>
              <p className="start-note">Перо было завершено, не сохранив изменения. Эти копии сохранены автоматически —
                откройте нужные и сохраните их.</p>
              {recovery.map((it) => (
                <div key={it.id} className="start-row">
                  <FileIcon kind={it.kind === 'sheet' ? 'sheet' : 'doc'} />
                  <div className="start-row-main">
                    <div className="start-row-name">{it.name}</div>
                    <div className="start-row-dir">{it.path ? it.path : 'новый, ещё не сохранявшийся'}</div>
                  </div>
                  <div className="start-row-when">{fmtWhen(it.time)}</div>
                  <button type="button" className="btn small primary" disabled={busy} onClick={() => restore(it)}>Открыть</button>
                  <button type="button" className="btn small danger" disabled={busy} onClick={() => discard(it)}>Удалить</button>
                </div>
              ))}
            </section>
          )}

          <section className="start-section">
            <div className="start-head">
              <h2>Последние</h2>
              <input className="start-search" type="search" placeholder="Поиск по имени и папке" value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>
            {recent && !recent.length && <div className="start-empty">Здесь появятся файлы, которые вы открывали и сохраняли.</div>}
            {recent && recent.length > 0 && !list.length && <div className="start-empty">Ничего не найдено</div>}
            <div className="start-list" role="list">
              {list.map((it) => (
                <div key={it.path} role="listitem" className={`start-row clickable${it.exists === false ? ' missing' : ''}`} title={it.exists === false ? `Файл не найден: ${it.path}` : it.path}
                  onClick={() => openRecent(it)} onKeyDown={(e) => { if (e.key === 'Enter') openRecent(it); }} tabIndex={0}>
                  <FileIcon kind={kindOf(it.name)} />
                  <div className="start-row-main">
                    <div className="start-row-name">{it.name}{it.exists === false && <span className="start-miss"> — не найден</span>}</div>
                    <div className="start-row-dir">{dirOf(it.path)}</div>
                  </div>
                  <div className="start-row-when">{fmtWhen(it.time)}</div>
                  <button type="button" className={`start-icon-btn${it.pinned ? ' on' : ''}`} title={it.pinned ? 'Открепить' : 'Закрепить в списке'}
                    onClick={async (e) => { e.stopPropagation(); await fileApi.recentPin(it.path, !it.pinned); reload(); }}><PinIcon on={it.pinned} /></button>
                  <button type="button" className="start-icon-btn" title="Убрать из списка (файл не удаляется)"
                    onClick={async (e) => { e.stopPropagation(); await fileApi.recentRemove(it.path); reload(); }}>✕</button>
                </div>
              ))}
            </div>
            <button type="button" className="btn start-more" onClick={openOther} disabled={busy}>Открыть другой файл…</button>
          </section>
        </main>
      </div>
    </div>
  );
}
