// Байты файла из загруженного JSZip. Большие сжатые части (фотографии) JSZip распаковывает на JS по кускам —
// 100 МБ фото ~1,7 с в главном потоке. Где есть встроенный DecompressionStream (окно Electron), сжатые deflate данные
// распаковываются им — вне главного потока и параллельно. Сжатые данные — внутреннее поле JSZip 3 (_data:
// CompressedObject); если его нет или формат другой — обычный путь JSZip.
const DEFLATE = '\x08\x00';
export async function zipBytes(file) {
  const d = file && file._data;
  if (typeof DecompressionStream === 'function' && d && d.compressedContent && d.compression && d.compression.magic === DEFLATE
    && d.uncompressedSize > 256 * 1024) {
    try {
      const stream = new Blob([d.compressedContent]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      const out = new Uint8Array(await new Response(stream).arrayBuffer());
      if (out.length === d.uncompressedSize) return out;
    } catch { /* повреждённые данные — пусть разбирается JSZip */ }
  }
  return file.async('uint8array');
}

// ---------- Сжатие больших частей при сохранении ----------
// JSZip сжимает deflate на JS (pako): лист книги на 60 МБ XML — ~3 с в главном потоке. Где есть встроенный
// CompressionStream, большие части сжимаются им (вне главного потока, параллельно) и отдаются JSZip уже сжатыми — его
// собственным CompressedObject (так JSZip хранит неизменённые части загруженного архива и пишет их без пересжатия).
// Конструктор и описание DEFLATE берутся из крошечного архива, прочитанного самим JSZip: отдельный импорт
// jszip/lib/compressedObject в сборке Vite дал бы другой класс, и проверка instanceof внутри JSZip не сработала бы.
let jszipInternals = null;
async function internals(JSZip) {
  if (!jszipInternals) {
    const probe = new JSZip();
    probe.file('x', 'xxxxxxxxxx');
    const back = await JSZip.loadAsync(await probe.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    const d = back.file('x')._data;
    // прототип задан литералом (constructor у него — Object): новый объект — от прототипа образца
    const proto = Object.getPrototypeOf(d);
    jszipInternals = {
      make: (compressedSize, uncompressedSize, crc, data) => Object.assign(Object.create(proto), { compressedSize, uncompressedSize, crc32: crc, compression: d.compression, compressedContent: data }),
    };
  }
  return jszipInternals;
}
let crcTable = null;
function crc32(bytes) {
  if (!crcTable) {
    crcTable = new Int32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; crcTable[n] = c; }
  }
  let crc = -1;
  for (let i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ -1) >>> 0;
}
// Сжать большие части (строки и байты, ещё не сжатые) встроенным deflate; JSZip — конструктор (тот же, что у zip)
export async function precompress(JSZip, zip, min = 512 * 1024) {
  if (typeof CompressionStream !== 'function') return;
  const parts = [];
  // добавленное через zip.file() JSZip хранит промисом (prepareContent) — дождаться значения
  zip.forEach((path, f) => { if (!f.dir) parts.push(f); });
  const big = [];
  for (const f of parts) {
    const d = f._data && typeof f._data.then === 'function' ? await f._data : f._data;
    if ((typeof d === 'string' || d instanceof Uint8Array) && d.length >= min) big.push([f, d]);
  }
  if (!big.length) return;
  const { make } = await internals(JSZip);
  await Promise.all(big.map(async ([f, d]) => {
    try {
      const bytes = typeof d === 'string' ? new TextEncoder().encode(d) : d;
      const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
      const comp = new Uint8Array(await new Response(stream).arrayBuffer());
      f._data = make(comp.length, bytes.length, crc32(bytes), comp);
      f._dataBinary = true;
    } catch { /* не вышло — пусть сожмёт JSZip */ }
  }));
}
