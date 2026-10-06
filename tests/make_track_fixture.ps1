# Генерирует tests/fixtures/track.docx средствами Word (COM): документ с исправлениями (вставка, удаление)
# от нейтрального автора «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_track_fixture.ps1 tests/fixtures/track.docx
param([string]$Out = 'tests/fixtures/track.docx')
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
  $d.Content.Text = 'Договор поставки товара заключён сегодня.'
  $d.TrackRevisions = $true
  # вставка: «срочной » перед «поставки»
  $r = $d.Content; $f = $r.Find; $f.Execute('поставки') | Out-Null
  $r.InsertBefore('срочной ')
  # удаление: слово «сегодня»
  $r = $d.Content; $f = $r.Find; $f.Execute(' сегодня') | Out-Null
  $r.Delete() | Out-Null
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
