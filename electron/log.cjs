// Журнал ошибок и событий: %APPDATA%\Перо\logs\pero.log, при превышении размера — в pero.1.log (одна копия).
const fs = require('fs');
const path = require('path');

const pad = (n, w = 2) => String(n).padStart(w, '0');
// Местное время: в журнале и имени отчёта — как на часах пользователя
function stamp(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function createLog(dir, { maxBytes = 1_000_000 } = {}) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, 'pero.log');
  const old = path.join(dir, 'pero.1.log');
  const write = (level, msg) => {
    try {
      const st = fs.statSync(file, { throwIfNoEntry: false });
      if (st && st.size > maxBytes) fs.renameSync(file, old); // на Windows перезаписывает старую копию
      const text = String(msg instanceof Error ? msg.stack || msg.message : msg).replace(/\r?\n/g, '\n    ');
      fs.appendFileSync(file, `${stamp()} [${level}] ${text}\n`, 'utf8');
    } catch { /* журнал не должен ронять программу */ }
  };
  // Последние n строк (с учётом предыдущей копии, если текущий файл короткий)
  const tail = (n = 300) => {
    const read = (f) => { try { return fs.readFileSync(f, 'utf8').split('\n').filter(Boolean); } catch { return []; } };
    let lines = read(file);
    if (lines.length < n) lines = [...read(old), ...lines];
    return lines.slice(-n);
  };
  return {
    file,
    dir,
    info: (m) => write('INFO', m),
    warn: (m) => write('WARN', m),
    error: (m) => write('ERROR', m),
    tail,
  };
}

module.exports = { createLog, stamp };
