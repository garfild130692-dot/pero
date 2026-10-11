// Последние файлы: recent.json в папке настроек; новые сверху, без повторов (путь без учёта регистра, как в Windows).
// Закреплённые (pinned) — всегда сверху и не вытесняются новыми; лимит max — для незакреплённых.
const fs = require('fs');
const path = require('path');

const same = (a, b) => path.resolve(a).toLowerCase() === path.resolve(b).toLowerCase();

class RecentFiles {
  constructor(file, max = 15) {
    this.file = file;
    this.max = max;
    this.items = [];
    try {
      const data = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (Array.isArray(data)) this.items = this.trim(data.filter((x) => x && typeof x.path === 'string'));
    } catch { /* нет файла или испорчен — начинаем с пустого списка */ }
  }

  // закреплённые — первыми (в своём порядке), затем не больше max остальных
  trim(items) {
    const pinned = items.filter((x) => x.pinned);
    return [...pinned, ...items.filter((x) => !x.pinned).slice(0, this.max)];
  }

  list() { return this.items.map((x) => ({ ...x })); }

  add(p) {
    const old = this.items.find((x) => same(x.path, p));
    const item = { path: path.resolve(p), name: path.basename(p), time: Date.now(), ...(old && old.pinned ? { pinned: true } : {}) };
    this.items = this.trim([item, ...this.items.filter((x) => !same(x.path, p))]);
    this.save();
    return item;
  }

  pin(p, value) {
    const it = this.items.find((x) => same(x.path, p));
    if (!it) return false;
    if (value) it.pinned = true; else delete it.pinned;
    this.items = this.trim([...this.items.filter((x) => x.pinned), ...this.items.filter((x) => !x.pinned).sort((a, b) => (b.time || 0) - (a.time || 0))]);
    this.save();
    return true;
  }

  remove(p) {
    const n = this.items.length;
    this.items = this.items.filter((x) => !same(x.path, p));
    if (this.items.length !== n) this.save();
  }

  // «Очистить список» — закреплённые остаются (как в Word)
  clear() { this.items = this.items.filter((x) => x.pinned); this.save(); }

  save() {
    try {
      fs.mkdirSync(path.dirname(this.file), { recursive: true });
      fs.writeFileSync(this.file, JSON.stringify(this.items, null, 1), 'utf8');
    } catch { /* список последних — не критичен */ }
  }
}

module.exports = { RecentFiles, samePath: same };
