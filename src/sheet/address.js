// Адреса ячеек: A1 ↔ (строка, столбец), диапазоны, ключи. Строки и столбцы внутри — с 0.

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export function colName(c) {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colIndex(name) {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r, c) => `${colName(c)}${r + 1}`;
export const key = (r, c) => `${r},${c}`;
export const unkey = (k) => k.split(',').map(Number);

// "A1", "$B$2" → { r, c, absR, absC } или null
export function parseRef(s) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(String(s).trim());
  if (!m) return null;
  const c = colIndex(m[2]);
  const r = +m[4] - 1;
  if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null;
  return { r, c, absC: !!m[1], absR: !!m[3] };
}

// "A1:B3" (или одна ячейка) → { r1, c1, r2, c2 } нормализованный
export function parseRange(s) {
  const [a, b] = String(s).split(':');
  const p = parseRef(a);
  const q = b ? parseRef(b) : p;
  if (!p || !q) return null;
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
}

export const rangeName = ({ r1, c1, r2, c2 }) => (r1 === r2 && c1 === c2 ? addr(r1, c1) : `${addr(r1, c1)}:${addr(r2, c2)}`);
export const inRange = (rg, r, c) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
