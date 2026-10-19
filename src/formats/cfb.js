// Составной файл Microsoft ([MS-CFB], «OLE2»): контейнер старых .doc и .xls — маленькая файловая система внутри файла
// (хранилища и потоки, таблица размещения секторов FAT, мини-поток для потоков меньше 4096 байт).
// readCfb(bytes) → { entries: [{ path, name, type: 'storage'|'stream', data? }], stream(path) }
// writeCfb([{ path: 'Workbook', data: Uint8Array }, { path: 'ObjectPool/_1', data }]) → Uint8Array (версия 3, сектор 512 байт)
// Пути — через «/», без корня: 'Workbook', '\u0005SummaryInformation', 'ObjectPool/_123/\u0001Ole'.

const SIG = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const FREE = 0xFFFFFFFF; const END = 0xFFFFFFFE; const FATSECT = 0xFFFFFFFD; const DIFSECT = 0xFFFFFFFC;
const NOSTREAM = 0xFFFFFFFF;

export const isCfb = (bytes) => bytes && bytes.length >= 512 && SIG.every((b, i) => bytes[i] === b);

export function readCfb(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (!isCfb(bytes)) throw new Error('Это не составной файл Microsoft (OLE2)');
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (o) => dv.getUint16(o, true); const u32 = (o) => dv.getUint32(o, true);
  const shift = u16(30); const size = 1 << shift;           // 512 (версия 3) или 4096 (версия 4)
  const miniSize = 1 << u16(32);
  const cutoff = u32(56);
  // заголовок по [MS-CFB] 2.2: сектор 512/4096, мини-сектор 64, порог мини-потока 4096 — иначе файл повреждён
  if ((shift !== 9 && shift !== 12) || u16(32) !== 6 || cutoff !== 4096) throw new Error('Составной файл повреждён (заголовок)');
  const nSec = Math.max(0, Math.floor(bytes.length / size) - 1); // секторов в файле — границы для всех чисел из заголовка
  const nFat = Math.min(u32(44), nSec); const dirStart = u32(48);
  const miniFatStart = u32(60); const nMiniFat = u32(64); let difat = u32(68); const nDifat = Math.min(u32(72), nSec);
  const off = (s) => (s + 1) * size;
  // секторы FAT: первые 109 — в заголовке, дальше — цепочка секторов DIFAT (без петель и в пределах файла)
  const fatSecs = [];
  // сектор FAT за пределами файла — заглушка (-1 → свободные записи), чтобы следующие секторы FAT не съехали
  for (let i = 0; i < 109 && fatSecs.length < nFat; i++) { const s = u32(76 + i * 4); if (s !== FREE) fatSecs.push(s < nSec ? s : -1); }
  const difatSeen = new Set();
  for (let k = 0; k < nDifat && difat < nSec && !difatSeen.has(difat); k++) {
    difatSeen.add(difat);
    const base = off(difat); const per = size / 4 - 1;
    for (let i = 0; i < per && fatSecs.length < nFat; i++) { const s = u32(base + i * 4); if (s !== FREE) fatSecs.push(s < nSec ? s : -1); }
    difat = u32(base + per * 4);
  }
  const fat = new Uint32Array(fatSecs.length * size / 4);
  fatSecs.forEach((s, i) => { for (let j = 0; j < size / 4; j++) fat[i * size / 4 + j] = s < 0 ? FREE : u32(off(s) + j * 4); });
  const chain = (start, table) => {
    const out = []; const seen = new Set();
    for (let s = start; s !== END && s !== FREE && s < table.length && !seen.has(s); s = table[s]) { seen.add(s); out.push(s); }
    return out;
  };
  const readChain = (start, len) => {
    const secs = chain(start, fat);
    // размер из каталога — не больше, чем данных в цепочке (враждебный размер иначе — гигантское выделение памяти)
    const out = new Uint8Array(Math.min(len ?? Infinity, secs.length * size));
    let p = 0;
    for (const s of secs) {
      const n = Math.min(size, out.length - p);
      if (n <= 0) break;
      out.set(bytes.subarray(off(s), off(s) + n), p); p += n;
    }
    return out;
  };
  // каталог: записи по 128 байт
  const dir = readChain(dirStart);
  const ddv = new DataView(dir.buffer);
  const raw = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const nl = ddv.getUint16(o + 64, true);
    let name = '';
    for (let i = 0; i + 2 < nl && i < 64; i += 2) name += String.fromCharCode(ddv.getUint16(o + i, true));
    const hi = ddv.getUint32(o + 124, true);
    raw.push({
      name, type: dir[o + 66], left: ddv.getUint32(o + 68, true), right: ddv.getUint32(o + 72, true), child: ddv.getUint32(o + 76, true),
      start: ddv.getUint32(o + 116, true), size: ddv.getUint32(o + 120, true) + (shift === 12 ? hi * 2 ** 32 : 0),
    });
  }
  const root = raw[0];
  // мини-поток (в цепочке корня) и мини-FAT
  const miniStream = root && root.type === 5 ? readChain(root.start, root.size) : new Uint8Array(0);
  const mfBytes = nMiniFat && miniFatStart !== END ? readChain(miniFatStart) : new Uint8Array(0);
  const miniFat = new Uint32Array(mfBytes.buffer, 0, Math.floor(mfBytes.length / 4));
  const readMini = (start, len) => {
    const secs = chain(start, miniFat);
    const out = new Uint8Array(Math.min(len, secs.length * miniSize)); let p = 0;
    for (const s of secs) {
      const n = Math.min(miniSize, out.length - p);
      if (n <= 0) break;
      out.set(miniStream.subarray(s * miniSize, s * miniSize + n), p); p += n;
    }
    return out;
  };
  // обход дерева каталога (красно-чёрное дерево братьев + child)
  const entries = [];
  const visit = (id, prefix, seen) => {
    if (id === NOSTREAM || id >= raw.length || seen.has(id)) return;
    seen.add(id);
    const e = raw[id];
    visit(e.left, prefix, seen);
    const path = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 1) { entries.push({ path, name: e.name, type: 'storage' }); visit(e.child, path, seen); }
    else if (e.type === 2) {
      let data;
      entries.push({ path, name: e.name, type: 'stream', get data() { return data || (data = e.size < cutoff ? readMini(e.start, e.size) : readChain(e.start, e.size)); } });
    }
    visit(e.right, prefix, seen);
  };
  if (root) visit(root.child, '', new Set([0]));
  const byPath = new Map(entries.map((e) => [e.path.toLowerCase(), e]));
  return { entries, stream: (p) => { const e = byPath.get(String(p).toLowerCase()); return e && e.type === 'stream' ? e.data : null; } };
}

// ---------- Запись ----------
// Порядок имён в дереве каталога по спецификации: сначала длина, потом без учёта регистра (верхний регистр)
const cmpName = (a, b) => (a.length !== b.length ? a.length - b.length : a.toUpperCase() < b.toUpperCase() ? -1 : a.toUpperCase() > b.toUpperCase() ? 1 : 0);

export function writeCfb(files) {
  const SIZE = 512; const MINI = 64; const CUTOFF = 4096;
  // дерево: корень → хранилища → потоки
  const nodes = [{ name: 'Root Entry', type: 5, kids: [] }];
  const storageOf = new Map([['', 0]]);
  const ensure = (path) => {
    if (storageOf.has(path)) return storageOf.get(path);
    const i = path.lastIndexOf('/');
    const parent = ensure(i < 0 ? '' : path.slice(0, i));
    const id = nodes.length;
    nodes.push({ name: path.slice(i + 1), type: 1, kids: [] });
    nodes[parent].kids.push(id);
    storageOf.set(path, id);
    return id;
  };
  for (const f of files) {
    const i = f.path.lastIndexOf('/');
    const parent = ensure(i < 0 ? '' : f.path.slice(0, i));
    const id = nodes.length;
    nodes.push({ name: f.path.slice(i + 1), type: 2, data: f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data) });
    nodes[parent].kids.push(id);
  }
  // братья — цепочкой вправо по порядку имён (все чёрные): читатели обходят дерево, сбалансированность не проверяют
  for (const n of nodes) {
    if (!n.kids) continue;
    const sorted = n.kids.slice().sort((a, b) => cmpName(nodes[a].name, nodes[b].name));
    n.child = sorted.length ? sorted[0] : NOSTREAM;
    sorted.forEach((id, k) => { nodes[id].right = k + 1 < sorted.length ? sorted[k + 1] : NOSTREAM; });
  }
  // мини-поток: потоки меньше 4096 байт
  const mini = []; let miniLen = 0;
  for (const n of nodes) if (n.type === 2 && n.data.length < CUTOFF) { n.miniStart = miniLen / MINI; n.size = n.data.length; mini.push(n); miniLen += Math.ceil(n.data.length / MINI) * MINI; }
  const miniFat = new Uint32Array(Math.max(0, miniLen / MINI));
  for (const n of mini) {
    const cnt = Math.ceil(n.size / MINI);
    if (!cnt) { n.miniStart = END; continue; }
    for (let k = 0; k < cnt; k++) miniFat[n.miniStart + k] = k + 1 < cnt ? n.miniStart + k + 1 : END;
  }
  const miniData = new Uint8Array(miniLen);
  for (const n of mini) if (n.size) miniData.set(n.data, n.miniStart * MINI);
  // размещение в секторах: обычные потоки, мини-поток, мини-FAT, каталог, затем FAT (и DIFAT)
  const big = nodes.filter((n) => n.type === 2 && n.data.length >= CUTOFF);
  const secCount = (len) => Math.ceil(len / SIZE);
  const dirSecs = secCount(nodes.length * 128);
  const miniFatSecs = secCount(miniFat.length * 4);
  let used = 0;
  for (const n of big) { n.start = used; n.size = n.data.length; used += secCount(n.size); }
  const miniStreamStart = miniLen ? used : END; used += secCount(miniLen);
  const miniFatStart = miniFatSecs ? used : END; used += miniFatSecs;
  const dirStart = used; used += dirSecs;
  // FAT должна покрыть и собственные секторы (и DIFAT)
  let fatSecs = 0; let difSecs = 0;
  for (;;) {
    const total = used + fatSecs + difSecs;
    const needFat = Math.ceil(total / (SIZE / 4));
    const needDif = needFat > 109 ? Math.ceil((needFat - 109) / (SIZE / 4 - 1)) : 0;
    if (needFat === fatSecs && needDif === difSecs) break;
    fatSecs = needFat; difSecs = needDif;
  }
  const fatStart = used; const difStart = used + fatSecs;
  const total = used + fatSecs + difSecs;
  const fat = new Uint32Array(fatSecs * SIZE / 4).fill(FREE);
  const link = (start, cnt) => { for (let k = 0; k < cnt; k++) fat[start + k] = k + 1 < cnt ? start + k + 1 : END; };
  for (const n of big) link(n.start, secCount(n.size));
  if (miniLen) link(miniStreamStart, secCount(miniLen));
  if (miniFatSecs) link(miniFatStart, miniFatSecs);
  link(dirStart, dirSecs);
  for (let k = 0; k < fatSecs; k++) fat[fatStart + k] = FATSECT;
  for (let k = 0; k < difSecs; k++) fat[difStart + k] = DIFSECT;
  const out = new Uint8Array((total + 1) * SIZE);
  const dv = new DataView(out.buffer);
  // заголовок
  out.set(SIG, 0);
  dv.setUint16(24, 0x003E, true); dv.setUint16(26, 3, true); dv.setUint16(28, 0xFFFE, true);
  dv.setUint16(30, 9, true); dv.setUint16(32, 6, true);
  dv.setUint32(44, fatSecs, true); dv.setUint32(48, dirStart, true); dv.setUint32(56, CUTOFF, true);
  dv.setUint32(60, miniFatStart, true); dv.setUint32(64, miniFatSecs, true);
  dv.setUint32(68, difSecs ? difStart : END, true); dv.setUint32(72, difSecs, true);
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i < fatSecs ? fatStart + i : FREE, true);
  const sec = (s) => (s + 1) * SIZE;
  // DIFAT: продолжение списка секторов FAT
  for (let d = 0; d < difSecs; d++) {
    const base = sec(difStart + d); const per = SIZE / 4 - 1;
    for (let i = 0; i < per; i++) { const idx = 109 + d * per + i; dv.setUint32(base + i * 4, idx < fatSecs ? fatStart + idx : FREE, true); }
    dv.setUint32(base + per * 4, d + 1 < difSecs ? difStart + d + 1 : END, true);
  }
  for (const n of big) out.set(n.data, sec(n.start));
  if (miniLen) out.set(miniData, sec(miniStreamStart));
  if (miniFatSecs) { const b = new Uint8Array(miniFatSecs * SIZE).fill(0xFF); b.set(new Uint8Array(miniFat.buffer)); out.set(b, sec(miniFatStart)); }
  // каталог
  const dir = new Uint8Array(dirSecs * SIZE);
  const ddv = new DataView(dir.buffer);
  for (let o = 0; o < dir.length; o += 128) { ddv.setUint32(o + 68, NOSTREAM, true); ddv.setUint32(o + 72, NOSTREAM, true); ddv.setUint32(o + 76, NOSTREAM, true); }
  nodes.forEach((n, id) => {
    const o = id * 128;
    const name = n.name.slice(0, 31);
    for (let i = 0; i < name.length; i++) ddv.setUint16(o + i * 2, name.charCodeAt(i), true);
    ddv.setUint16(o + 64, (name.length + 1) * 2, true);
    dir[o + 66] = n.type; dir[o + 67] = 1; // чёрный
    ddv.setUint32(o + 68, NOSTREAM, true);
    ddv.setUint32(o + 72, n.right ?? NOSTREAM, true);
    ddv.setUint32(o + 76, n.kids ? n.child : NOSTREAM, true);
    if (n.type === 5) { ddv.setUint32(o + 116, miniStreamStart, true); ddv.setUint32(o + 120, miniLen, true); }
    else if (n.type === 2) { ddv.setUint32(o + 116, n.data.length < CUTOFF ? n.miniStart : n.start, true); ddv.setUint32(o + 120, n.data.length, true); }
    else ddv.setUint32(o + 116, 0, true);
  });
  out.set(dir, sec(dirStart));
  out.set(new Uint8Array(fat.buffer), sec(fatStart));
  return out;
}
