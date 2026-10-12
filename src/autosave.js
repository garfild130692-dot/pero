// Автосохранение для восстановления после сбоя: раз в intervalSec секунд, если есть несохранённые правки
// и они изменились с прошлой копии, документ сериализуется и отдаётся главному процессу (autosave:write).
// Когда правок не стало (файл сохранён) — копия удаляется.
// Большие файлы (этап 17): сериализация книги на 100 000 строк — секунды в главном потоке. Поэтому копия делается
// в паузе набора (2 с без правок; если правки идут непрерывно — всё равно не реже чем раз в 5 интервалов), а после
// долгой копии следующая — не раньше чем через 20 её длительностей (6 с → раз в 2 минуты).
import { useEffect, useRef } from 'react';
import * as fileApi from './fileApi.js';

const IDLE_MS = 2000;
const COST_FACTOR = 20;
const FORCE_FACTOR = 5;

// opts: { intervalSec, dirty, revision (значение или функция; меняется при правке), kind, name, path, getBytes: async () => bytes }
export function useAutosave(opts) {
  const cur = useRef(opts);
  cur.current = opts;
  const saved = useRef(undefined); // revision последней копии
  const busy = useRef(false);

  useEffect(() => {
    const sec = opts.intervalSec;
    if (!sec) return undefined;
    const period = sec * 1000;
    const idle = Math.min(IDLE_MS, period / 2);
    let lastAt = Date.now(); // конец последней копии (или начало работы)
    let cost = 0; // длительность последней сериализации, мс
    let seenRev; let seenAt = Date.now(); // последняя замеченная правка
    const t = setInterval(async () => {
      const o = cur.current;
      const rev = typeof o.revision === 'function' ? o.revision() : o.revision;
      const now = Date.now();
      if (rev !== seenRev) { seenRev = rev; seenAt = now; }
      if (!o.dirty || busy.current || rev === saved.current) return;
      const gap = Math.max(period, COST_FACTOR * cost);
      if (now - lastAt < gap) return;
      if (now - seenAt < idle && now - lastAt < FORCE_FACTOR * gap) return; // пользователь печатает — подождать паузы
      busy.current = true;
      const t0 = Date.now();
      try {
        const bytes = await o.getBytes();
        cost = Date.now() - t0;
        if (await fileApi.autosave({ kind: o.kind, name: o.name, path: o.path }, bytes)) saved.current = rev;
      } catch (e) {
        fileApi.logError(`Автосохранение: ${e && e.stack ? e.stack : e}`);
      } finally {
        busy.current = false;
        lastAt = Date.now();
      }
    }, Math.min(period, 1000));
    return () => clearInterval(t);
  }, [opts.intervalSec]);

  // правки были и пропали (сохранено) — копия больше не нужна. При загрузке окна не чистим:
  // в копии может быть восстановленный документ, который ещё не успели переписать.
  const wasDirty = useRef(false);
  useEffect(() => {
    if (wasDirty.current && !opts.dirty) { fileApi.autosaveClear(); saved.current = undefined; }
    wasDirty.current = opts.dirty;
  }, [opts.dirty]);
}
