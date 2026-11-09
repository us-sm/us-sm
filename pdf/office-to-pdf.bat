@echo off
rem ============================================================
rem  office-to-pdf.bat
rem  Converts Word / Excel / PowerPoint files to PDF using the
rem  Microsoft Office installed on this PC.
rem
rem  Usage:
rem    - Drag files (or a folder) onto this .bat, or
rem    - Double-click it and pick files.
rem  PDFs are saved next to the original files.
rem ============================================================
setlocal
title Office to PDF
set "W2P_ARGS=%*"
set "W2P_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:W2P_SELF); $p=($s -split '(?m)^#__PS__\r?\n',2)[1]; & ([scriptblock]::Create($p))"
echo.
pause
exit /b
#__PS__
$ErrorActionPreference = 'Stop'
$types = @{
  '.doc'='word'; '.docx'='word'; '.docm'='word'; '.rtf'='word'; '.odt'='word'
  '.xls'='excel'; '.xlsx'='excel'; '.xlsm'='excel'; '.csv'='excel'
  '.ppt'='ppt'; '.pptx'='ppt'; '.pptm'='ppt'
}

# ---- collect input paths ----
$inputs = @()
if ($env:W2P_ARGS) {
  foreach ($m in [regex]::Matches($env:W2P_ARGS, '"([^"]+)"|(\S+)')) {
    if ($m.Groups[1].Success) { $inputs += $m.Groups[1].Value } else { $inputs += $m.Groups[2].Value }
  }
} else {
  Add-Type -AssemblyName System.Windows.Forms
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Multiselect = $true
  $dlg.Title = 'Choose Word / Excel / PowerPoint files to convert to PDF'
  $dlg.Filter = 'Office files|*.doc;*.docx;*.docm;*.rtf;*.odt;*.xls;*.xlsx;*.xlsm;*.csv;*.ppt;*.pptx;*.pptm|All files|*.*'
  if ($dlg.ShowDialog() -ne 'OK') { Write-Host 'Nothing selected.'; return }
  $inputs = $dlg.FileNames
}

$files = @()
foreach ($i in $inputs) {
  if (Test-Path -LiteralPath $i -PathType Container) {
    $files += Get-ChildItem -LiteralPath $i -File | Where-Object { $types.ContainsKey($_.Extension.ToLower()) }
  } elseif (Test-Path -LiteralPath $i -PathType Leaf) {
    $files += Get-Item -LiteralPath $i
  } else {
    Write-Host "Not found: $i" -ForegroundColor Yellow
  }
}
$files = $files | Where-Object { -not $_.Name.StartsWith('~$') }
if (-not $files) { Write-Host 'No Word / Excel / PowerPoint files found.' -ForegroundColor Yellow; return }

function Get-OutPath($f) {
  $base = Join-Path $f.DirectoryName ([IO.Path]::GetFileNameWithoutExtension($f.Name))
  $out = "$base.pdf"; $n = 1
  while (Test-Path -LiteralPath $out) { $out = "$base ($n).pdf"; $n++ }
  return $out
}

$word = $null; $excel = $null; $ppt = $null
$ok = 0; $fail = 0
try {
  foreach ($f in $files) {
    $kind = $types[$f.Extension.ToLower()]
    if (-not $kind) { Write-Host "Skipped (not an Office file): $($f.Name)" -ForegroundColor Yellow; continue }
    $out = Get-OutPath $f
    Write-Host "Converting $($f.Name) ... " -NoNewline
    try {
      switch ($kind) {
        'word' {
          if (-not $word) { $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0 }
          # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument)
          $doc = $word.Documents.Open($f.FullName, $false, $true, $false, '__nopass__')
          $doc.ExportAsFixedFormat($out, 17)   # 17 = PDF
          $doc.Close(0)
        }
        'excel' {
          if (-not $excel) { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false }
          # Open(Filename, UpdateLinks, ReadOnly, Format, Password)
          $wb = $excel.Workbooks.Open($f.FullName, 0, $true, 5, '__nopass__')
          $wb.ExportAsFixedFormat(0, $out)     # 0 = PDF, all sheets
          $wb.Close($false)
        }
        'ppt' {
          if (-not $ppt) { $ppt = New-Object -ComObject PowerPoint.Application }
          # Open(FileName, ReadOnly, Untitled, WithWindow)
          $pres = $ppt.Presentations.Open($f.FullName, -1, 0, 0)
          $pres.SaveAs($out, 32)               # 32 = PDF
          $pres.Close()
        }
      }
      Write-Host 'OK' -ForegroundColor Green
      $ok++
    } catch {
      Write-Host 'FAILED' -ForegroundColor Red
      Write-Host "   $($_.Exception.Message)" -ForegroundColor DarkGray
      Write-Host '   (Password-protected or damaged files cannot be converted.)' -ForegroundColor DarkGray
      $fail++
    }
  }
} finally {
  if ($word)  { $word.Quit(0);  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
  if ($excel) { $excel.Quit();  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  if ($ppt)   { $ppt.Quit();    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) }
  [GC]::Collect()
}

Write-Host ''
Write-Host "Done: $ok converted, $fail failed. PDFs are saved next to the original files." -ForegroundColor Cyan
Write-Host 'To merge them into one PDF, use the PDF Tools page.'
