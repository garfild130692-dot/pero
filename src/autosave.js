// Автосохранение для восстановления после сбоя: раз в intervalSec секунд, если есть несохранённые правки
// и они изменились с прошлой копии, документ сериализуется и отдаётся главному процессу (autosave:write).
// Когда правок не стало (файл сохранён) — копия удаляется.
import { useEffect, useRef } from 'react';
import * as fileApi from './fileApi.js';

// opts: { intervalSec, dirty, revision (значение или функция; меняется при правке), kind, name, path, getBytes: async () => bytes }
export function useAutosave(opts) {
  const cur = useRef(opts);
  cur.current = opts;
  const saved = useRef(undefined); // revision последней копии
  const busy = useRef(false);

  useEffect(() => {
    const sec = opts.intervalSec;
    if (!sec) return undefined;
    const t = setInterval(async () => {
      const o = cur.current;
      const rev = typeof o.revision === 'function' ? o.revision() : o.revision;
      if (!o.dirty || busy.current || rev === saved.current) return;
      busy.current = true;
      try {
        const bytes = await o.getBytes();
        if (await fileApi.autosave({ kind: o.kind, name: o.name, path: o.path }, bytes)) saved.current = rev;
      } catch (e) {
        fileApi.logError(`Автосохранение: ${e && e.stack ? e.stack : e}`);
      } finally {
        busy.current = false;
      }
    }, sec * 1000);
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
