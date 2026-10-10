// Защита листа (как «Рецензирование → Защитить лист» Excel): sheet.protection = { attrs } — атрибуты <sheetProtection>
// из файла или созданные Пером (хеш пароля SHA-512 с солью, как пишет Excel). Ячейки по умолчанию заблокированы;
// стиль { unlocked: true } — можно править на защищённом листе, { formulaHidden: true } — формула не показывается.

export const isProtected = (sheet) => !!(sheet && sheet.protection);
export const isLocked = (cell) => !(cell && cell.s && cell.s.unlocked);

// Правка заблокированного на защищённом листе? cur/next — книги до и после операции.
// Пересчёт формул (изменилось только значение формулы) правкой не считается.
export function protectionViolation(cur, next) {
  // добавление/удаление листов защита листа не запрещает (это защита книги)
  if (cur.sheets.length !== next.sheets.length) return false;
  for (let i = 0; i < next.sheets.length; i++) {
    const a = cur.sheets[i];
    const b = next.sheets[i];
    if (!a || a === b || !isProtected(a) || !isProtected(b)) continue;
    // строки/столбцы, объединения, размеры — заблокированы целиком
    if (a.merges !== b.merges || a.colW !== b.colW || a.rowH !== b.rowH || a.rowInfo !== b.rowInfo || a.colInfo !== b.colInfo) return true;
    if (a.cells === b.cells) continue;
    const keys = new Set([...a.cells.keys(), ...b.cells.keys()]);
    for (const k of keys) {
      const x = a.cells.get(k); const y = b.cells.get(k);
      if (x === y) continue;
      const sameF = (x && x.f) === (y && y.f);
      const sameS = JSON.stringify(x && x.s) === JSON.stringify(y && y.s);
      if (sameF && sameS && x && x.f != null) continue; // пересчёт формулы
      if (isLocked(x)) return true;
    }
  }
  return false;
}

export function setProtection(wb, si, protection) {
  const sheet = { ...wb.sheets[si], protection: protection || null };
  if (!protection) delete sheet.protection;
  return { ...wb, sheets: wb.sheets.map((s, i) => (i === si ? sheet : s)) };
}

// ---------- Пароль ----------
const b64 = (bytes) => btoa(String.fromCharCode(...bytes));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const utf16le = (s) => { const out = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 255; out[i * 2 + 1] = s.charCodeAt(i) >> 8; } return out; };
const ALGO = { 'SHA-512': 'SHA-512', 'SHA-384': 'SHA-384', 'SHA-256': 'SHA-256', 'SHA-1': 'SHA-1' };

// Хеш пароля по ISO/IEC 29500: H0 = H(соль + пароль UTF-16LE), Hn = H(Hn-1 + n (4 байта LE))
export async function hashPassword(password, saltB64, spinCount, algorithm = 'SHA-512') {
  const subtle = globalThis.crypto.subtle;
  const salt = unb64(saltB64);
  const pw = utf16le(password);
  const first = new Uint8Array(salt.length + pw.length);
  first.set(salt); first.set(pw, salt.length);
  let h = new Uint8Array(await subtle.digest(ALGO[algorithm] || 'SHA-512', first));
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    buf[h.length] = i & 255; buf[h.length + 1] = (i >>> 8) & 255; buf[h.length + 2] = (i >>> 16) & 255; buf[h.length + 3] = (i >>> 24) & 255;
    h = new Uint8Array(await subtle.digest(ALGO[algorithm] || 'SHA-512', buf));
  }
  return b64(h);
}
// Старый 16-битный хеш (атрибут password) — книги из старых версий Excel
export function legacyHash(password) {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= password.charCodeAt(i);
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= password.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}
export const hasPassword = (protection) => !!(protection && protection.attrs && (protection.attrs.hashValue || protection.attrs.password));
// Подходит ли пароль к защите (без пароля — любой, в т. ч. пустой)
export async function checkPassword(protection, password) {
  const a = (protection && protection.attrs) || {};
  if (a.hashValue) return (await hashPassword(password, a.saltValue || '', +(a.spinCount || 0), a.algorithmName)) === a.hashValue;
  if (a.password) return legacyHash(password) === a.password.toUpperCase();
  return true;
}
// Новая защита: как у Excel — SHA-512, соль 16 байт, 100 000 повторов; без пароля — только флаги
export async function makeProtection(password) {
  const attrs = { sheet: '1', objects: '1', scenarios: '1' };
  if (password) {
    const salt = new Uint8Array(16);
    globalThis.crypto.getRandomValues(salt);
    const saltValue = b64(salt);
    Object.assign(attrs, { algorithmName: 'SHA-512', hashValue: await hashPassword(password, saltValue, 100000), saltValue, spinCount: '100000' });
  }
  return { attrs };
}
