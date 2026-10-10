# Генерирует tests/fixtures/hf.docx средствами Word (COM): особый колонтитул первой страницы, разные колонтитулы
# чётных и нечётных страниц, номер страницы полем PAGE; 5 страниц текста. Автор — нейтральный «Перо Тест».
# Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_hf_fixture.ps1 tests/fixtures/hf.docx
param([string]$Out = 'tests/fixtures/hf.docx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName
$oldInitials = $w.UserInitials
$oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true
$w.UserName = 'Перо Тест'
$w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $text = ('Абзац текста для заполнения страниц. ' * 12).Trim()
  $body = (1..40 | ForEach-Object { "$_. $text" }) -join "`r"
  $d.Content.Text = $body
  $ps = $d.Sections.Item(1).PageSetup
  $ps.DifferentFirstPageHeaderFooter = -1
  $ps.OddAndEvenPagesHeaderFooter = -1
  $s = $d.Sections.Item(1)
  # 1 — основной (нечётные), 2 — первая страница, 3 — чётные
  $s.Headers.Item(2).Range.Text = 'Титульный лист'
  $s.Headers.Item(1).Range.Text = 'Нечётная страница'
  $s.Headers.Item(1).Range.ParagraphFormat.Alignment = 2
  $s.Headers.Item(3).Range.Text = 'Чётная страница'
  foreach ($k in 1, 3) {
    $r = $s.Footers.Item($k).Range
    $r.Text = 'Стр. '
    $r.Collapse(0)
    [void]$d.Fields.Add($r, 33)   # wdFieldPage
    $s.Footers.Item($k).Range.ParagraphFormat.Alignment = 1
  }
  $d.SaveAs2($path, 16)
  'pages=' + $d.ComputeStatistics(2)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
