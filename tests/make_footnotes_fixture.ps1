# Генерирует tests/fixtures/footnotes.docx средствами Word (COM): три абзаца и две сноски внизу страницы.
# Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_footnotes_fixture.ps1 tests/fixtures/footnotes.docx
param([string]$Out = 'tests/fixtures/footnotes.docx')
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
  $d.Content.Text = "Договор поставки.`rСрок оплаты — 10 дней.`rПорядок приёмки."
  $p1 = $d.Paragraphs.Item(1).Range
  $d.Footnotes.Add($d.Range($p1.End - 2, $p1.End - 2), [Type]::Missing, 'Типовая форма компании.') | Out-Null
  $p2 = $d.Paragraphs.Item(2).Range
  $d.Footnotes.Add($d.Range($p2.End - 2, $p2.End - 2), [Type]::Missing, 'Календарных дней с даты счёта.') | Out-Null
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
