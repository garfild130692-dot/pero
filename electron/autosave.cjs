// Автосохранение для восстановления после сбоя.
// Каждое окно пишет копию несохранённого документа в %APPDATA%\Перо\autosave\<pid>-<окно>.bin + .json (сведения).
// Нормальное закрытие окна и успешное сохранение файла копию удаляют; оставшиеся после «умершего» процесса
// копии при следующем запуске предлагаются к восстановлению.
const fs = require('fs');
const path = require('path');

// Жив ли процесс: kill(pid, 0) ничего не посылает, только проверяет
function pidAlive(pid) {
  try { process.kill(pid, 0); return true; } catch (e) { return e.code === 'EPERM'; }
}

function createAutosave(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const bin = (key) => path.join(dir, `${key}.bin`);
  const json = (key) => path.join(dir, `${key}.json`);
  const rm = (f) => { try { fs.rmSync(f, { force: true }); } catch { /* занято — удалится при следующем запуске */ } };

  // запись через временный файл: прерванная запись не портит прошлую копию
  function write(key, meta, data) {
    const tmp = `${bin(key)}.tmp`;
    fs.writeFileSync(tmp, Buffer.from(data));
    fs.renameSync(tmp, bin(key));
    fs.writeFileSync(json(key), JSON.stringify({ ...meta, key, time: Date.now() }), 'utf8');
  }

  function readMeta(key) {
    try { return JSON.parse(fs.readFileSync(json(key), 'utf8')); } catch { return null; }
  }

  // Копия конкретного окна (перезапуск окна после сбоя в том же процессе)
  function readOwn(key) {
    const meta = readMeta(key);
    if (!meta || !fs.existsSync(bin(key))) return null;
    return { meta, data: fs.readFileSync(bin(key)) };
  }

  // Копии, оставшиеся от завершившихся процессов: [{ id, name, path, kind, time }]
  function scan({ currentPid = process.pid, isAlive = pidAlive } = {}) {
    const out = [];
    for (const f of fs.readdirSync(dir)) {
      if (f.endsWith('.tmp')) { rm(path.join(dir, f)); continue; }
      if (!f.endsWith('.json')) continue;
      const key = f.slice(0, -5);
      const meta = readMeta(key);
      if (!meta || !fs.existsSync(bin(key))) { rm(json(key)); continue; }
      if (meta.pid === currentPid || isAlive(meta.pid)) continue;
      out.push({ id: key, name: meta.name, path: meta.path || null, kind: meta.kind, time: meta.time });
    }
    return out.sort((a, b) => b.time - a.time);
  }

  // Передать копию другому окну (восстановленный документ дальше автосохраняется под ключом этого окна)
  function adopt(id, key) {
    const meta = readMeta(id);
    if (!meta || !fs.existsSync(bin(id))) return null;
    const data = fs.readFileSync(bin(id));
    if (id !== key) {
      write(key, { ...meta, pid: process.pid }, data);
      clear(id);
    }
    return { meta, data };
  }

  function clear(key) { rm(bin(key)); rm(json(key)); }

  return { dir, write, readOwn, scan, adopt, clear };
}

module.exports = { createAutosave, pidAlive };
