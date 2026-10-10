# Эталон разбиения на страницы от Word: начало каждой страницы — номер абзаца тела документа (с 0, абзацы таблиц
# считаются по ячейкам, как их видит Word) и смещение символа в абзаце. Вывод — JSON.
# Запуск: powershell -NoProfile -ExecutionPolicy Bypass -File tests/word_pages.ps1 -path <docx>
param([string]$path)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$w = New-Object -ComObject Word.Application
try {
  $d = $w.Documents.Open((Resolve-Path $path).Path, $false, $true)
  $d.Repaginate()
  $pages = $d.ComputeStatistics(2)
  # начала абзацев (по символам) для перевода позиции в (абзац, смещение)
  $starts = New-Object System.Collections.Generic.List[int]
  $texts = New-Object System.Collections.Generic.List[string]
  foreach ($p in $d.Paragraphs) { $starts.Add($p.Range.Start); $texts.Add($p.Range.Text) }
  $out = @()
  for ($n = 2; $n -le $pages; $n++) {
    $r = $d.GoTo(1, 1, $n)  # wdGoToPage, wdGoToAbsolute
    $pos = $r.Start
    $i = 0
    $lo = 0; $hi = $starts.Count - 1
    while ($lo -le $hi) { $mid = [int](($lo + $hi) / 2); if ($starts[$mid] -le $pos) { $i = $mid; $lo = $mid + 1 } else { $hi = $mid - 1 } }
    $off = $pos - $starts[$i]
    $snip = $texts[$i].Substring($off, [Math]::Min(25, $texts[$i].Length - $off)) -replace "[`r`a`f]", ''
    $out += [pscustomobject]@{ page = $n; para = $i; offset = $off; text = $snip }
  }
  [pscustomobject]@{ pages = $pages; paragraphs = $starts.Count; breaks = $out } | ConvertTo-Json -Depth 4 -Compress
  $d.Close(0)
} finally { $w.Quit() }
