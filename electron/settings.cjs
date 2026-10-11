// Настройки программы: settings.json в папке настроек. Неизвестные ключи и неверные значения — умолчания.
const fs = require('fs');
const path = require('path');

const DEFAULTS = {
  startScreen: true, // при запуске без файла — начальный экран (иначе сразу новый документ)
};

class Settings {
  constructor(file) {
    this.file = file;
    this.values = { ...DEFAULTS };
    try {
      const data = JSON.parse(fs.readFileSync(file, 'utf8'));
      for (const k of Object.keys(DEFAULTS)) if (data && typeof data[k] === typeof DEFAULTS[k]) this.values[k] = data[k];
    } catch { /* нет файла или испорчен — умолчания */ }
  }

  get(k) { return k ? this.values[k] : { ...this.values }; }

  set(k, v) {
    if (!(k in DEFAULTS) || typeof v !== typeof DEFAULTS[k]) return false;
    this.values[k] = v;
    try {
      fs.mkdirSync(path.dirname(this.file), { recursive: true });
      fs.writeFileSync(this.file, JSON.stringify(this.values, null, 1), 'utf8');
    } catch { return false; }
    return true;
  }
}

module.exports = { Settings, SETTINGS_DEFAULTS: DEFAULTS };
