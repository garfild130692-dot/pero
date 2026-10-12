// Общие объекты Range для замеров. Каждый document.createRange() — «живой» диапазон: пока его не соберёт сборщик мусора,
// Chromium обновляет его при каждом изменении DOM. Замер разбиения создавал их тысячами, и снятие распорок на 200+ листах
// стоило ~0,5 с (removeChild). Здесь — несколько постоянных диапазонов по номеру, которые переиспользуются.
const slots = [];
export function scratchRange(i = 0) {
  if (!slots[i]) slots[i] = document.createRange();
  return slots[i];
}
