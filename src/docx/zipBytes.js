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
