# Эталоны .doc (этап 19.4): сам Word (COM) пересохраняет эталонные документы .docx в формат Word 97–2003
# (wdFormatDocument97 = 0), а полученный .doc — снова в .docx (wdFormatXMLDocument = 12): <имя>.doc и <имя>.doc.docx.
# Перо читает .doc своим кодом — модель должна совпасть с моделью .doc.docx (то, как сам Word понял свой .doc).
# Автор — нейтральный «Перо Тест» (UserName возвращается). Файл в UTF-8 с BOM (PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_doc_fixtures.ps1 [имя ...] (без имён — все)
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Names)
$ErrorActionPreference = 'Stop'
$dir = [IO.Path]::GetFullPath('tests/fixtures')
$log = [IO.Path]::GetFullPath('tests/out/make_doc_fixtures.log')
function Step($m) { Add-Content -Path $log -Value "$(Get-Date -Format HH:mm:ss) $m" -Encoding UTF8 }
Step 'старт'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName
$oldInitials = $w.UserInitials
$oldLocal = $w.Options.UseLocalUserInfo
# присваивать только при отличии: присваивание тех же значений вешает последующий SaveAs2 в .doc (проверено)
if (-not $oldLocal) { $w.Options.UseLocalUserInfo = $true }
if ($oldName -ne 'Перо Тест') { $w.UserName = 'Перо Тест' }
if ($oldInitials -ne 'ПТ') { $w.UserInitials = 'ПТ' }
Step 'имя пользователя'
try {
  $all = @('sample', 'fonts', 'para', 'tables', 'layout', 'lists', 'numbering', 'hf', 'sections', 'footnotes', 'preserve', 'toc', 'columns', 'track')
  foreach ($name in $(if ($Names) { $Names } else { $all })) {
    # [string]: Join-Path отдаёт строку в обёртке PowerShell — с ней SaveAs2 в .doc зависает (проверено)
    $src = [string](Join-Path $dir "$name.docx")
    $dst = [string](Join-Path $dir "$name.doc")
    $back = [string](Join-Path $dir "$name.doc.docx")
    foreach ($p in @($dst, $back)) { if (Test-Path $p) { Remove-Item $p } }
    $d = $w.Documents.Open($src, $false, $true, $false)
    Step "${name}: открыт"
    $d.SaveAs2($dst, 0)
    Step "${name}: .doc"
    $d.Close()
    $d = $w.Documents.Open($dst, $false, $true, $false)
    $d.SaveAs2($back, 12)
    Step "${name}: .doc.docx"
    $d.Close()
    Write-Output "$name.doc"
  }
} finally {
  if ($w.UserName -ne $oldName) { $w.UserName = $oldName }
  if ($w.UserInitials -ne $oldInitials) { $w.UserInitials = $oldInitials }
  if ($w.Options.UseLocalUserInfo -ne $oldLocal) { $w.Options.UseLocalUserInfo = $oldLocal }
  $w.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
