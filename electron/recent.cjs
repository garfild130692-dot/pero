// Последние файлы: recent.json в папке настроек; новые сверху, без повторов (путь без учёта регистра, как в Windows).
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
      if (Array.isArray(data)) this.items = data.filter((x) => x && typeof x.path === 'string').slice(0, max);
    } catch { /* нет файла или испорчен — начинаем с пустого списка */ }
  }

  list() { return this.items.map((x) => ({ ...x })); }

  add(p) {
    const item = { path: path.resolve(p), name: path.basename(p), time: Date.now() };
    this.items = [item, ...this.items.filter((x) => !same(x.path, p))].slice(0, this.max);
    this.save();
    return item;
  }

  remove(p) {
    const n = this.items.length;
    this.items = this.items.filter((x) => !same(x.path, p));
    if (this.items.length !== n) this.save();
  }

  clear() { this.items = []; this.save(); }

  save() {
    try {
      fs.mkdirSync(path.dirname(this.file), { recursive: true });
      fs.writeFileSync(this.file, JSON.stringify(this.items, null, 1), 'utf8');
    } catch { /* список последних — не критичен */ }
  }
}

module.exports = { RecentFiles, samePath: same };
