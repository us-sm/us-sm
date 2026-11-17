@echo off
rem ============================================================
rem  video-compress.bat
rem  Compresses videos (MOV, MP4, AVI, MKV, ...) to small MP4 files
rem  using ffmpeg. Much faster than the web page for AVI / large files.
rem
rem  Usage:
rem    - Drag videos (or a folder) onto this .bat, or
rem    - Double-click it and pick files.
rem  First run downloads ffmpeg (~90 MB) to %LOCALAPPDATA%\us-sm\ffmpeg.
rem  No install or admin rights needed. If the download is blocked,
rem  put ffmpeg.exe + ffprobe.exe next to this .bat.
rem  New files are saved next to the originals as *_compressed.mp4.
rem ============================================================
setlocal
title Compress Video
set "VC_ARGS=%*"
set "VC_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText($env:VC_SELF); $p=($s -split '(?m)^#__PS__\r?\n',2)[1]; & ([scriptblock]::Create($p))"
echo.
pause
exit /b
#__PS__
$ErrorActionPreference = 'Stop'
$exts = '.mov','.mp4','.m4v','.avi','.mkv','.webm','.wmv','.flv','.mpg','.mpeg','.3gp','.mts','.m2ts','.ts'

# ---- find or download ffmpeg (no install, no admin) ----
$selfDir = if ($env:VC_SELF) { Split-Path -Parent $env:VC_SELF } else { (Get-Location).Path }
$cacheDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'us-sm\ffmpeg'
function Find-Tool($name) {
  foreach ($d in @($selfDir, $cacheDir)) {
    $p = Join-Path $d "$name.exe"
    if (Test-Path -LiteralPath $p) { return $p }
  }
  $c = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($c) { return $c.Source }
  return $null
}
$ffmpeg = Find-Tool 'ffmpeg'; $ffprobe = Find-Tool 'ffprobe'
if (-not $ffmpeg -or -not $ffprobe) {
  $url = 'https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl-shared.zip'
  Write-Host 'First run: downloading ffmpeg (about 90 MB, one time only)...' -ForegroundColor Cyan
  New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
  $zip = Join-Path ([IO.Path]::GetTempPath()) 'us-sm-ffmpeg.zip'
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) { & $curl.Source -L --fail --progress-bar -o $zip $url; if ($LASTEXITCODE -ne 0) { throw "download failed (curl exit $LASTEXITCODE)" } }
    else { $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $z = [IO.Compression.ZipFile]::OpenRead($zip)
    try {
      foreach ($e in $z.Entries) {
        if ($e.FullName -match '/bin/(ffmpeg\.exe|ffprobe\.exe|[^/]+\.dll)$') {
          [IO.Compression.ZipFileExtensions]::ExtractToFile($e, (Join-Path $cacheDir $e.Name), $true)
        }
      }
    } finally { $z.Dispose() }
  } catch {
    Write-Host "Could not download ffmpeg: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Fix: download ffmpeg yourself and put ffmpeg.exe + ffprobe.exe next to this .bat, then run it again.' -ForegroundColor Yellow
    Write-Host '     Or use the web page (https://us-sm.github.io/video/) instead.' -ForegroundColor Yellow
    return
  } finally { Remove-Item -LiteralPath $zip -ErrorAction SilentlyContinue }
  $ffmpeg = Find-Tool 'ffmpeg'; $ffprobe = Find-Tool 'ffprobe'
  if (-not $ffmpeg -or -not $ffprobe) { Write-Host 'ffmpeg was downloaded but could not be found. Try again.' -ForegroundColor Red; return }
  Write-Host 'ffmpeg is ready.' -ForegroundColor Green
  Write-Host ''
}

# ---- collect input files ----
$inputs = @()
if ($env:VC_ARGS) {
  foreach ($m in [regex]::Matches($env:VC_ARGS, '"([^"]+)"|(\S+)')) {
    if ($m.Groups[1].Success) { $inputs += $m.Groups[1].Value } else { $inputs += $m.Groups[2].Value }
  }
} else {
  Add-Type -AssemblyName System.Windows.Forms
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Multiselect = $true
  $dlg.Title = 'Choose videos to compress'
  $dlg.Filter = 'Videos|' + (($exts | ForEach-Object { "*$_" }) -join ';') + '|All files|*.*'
  if ($dlg.ShowDialog() -ne 'OK') { Write-Host 'Nothing selected.'; return }
  $inputs = $dlg.FileNames
}
$files = @()
foreach ($i in $inputs) {
  if (Test-Path -LiteralPath $i -PathType Container) {
    $files += Get-ChildItem -LiteralPath $i -File | Where-Object { $exts -contains $_.Extension.ToLower() -and $_.BaseName -notlike '*_compressed' }
  } elseif (Test-Path -LiteralPath $i -PathType Leaf) {
    $files += Get-Item -LiteralPath $i
  } else {
    Write-Host "Not found: $i" -ForegroundColor Yellow
  }
}
if (-not $files) { Write-Host 'No video files found.' -ForegroundColor Yellow; return }

# ---- choose a preset ----
Write-Host "$($files.Count) video(s) selected." -ForegroundColor Cyan
Write-Host '  1) Email      - fit under 25 MB each'
Write-Host '  2) Normal     - 1080p, good quality (recommended)'
Write-Host '  3) High       - keep resolution, high quality'
Write-Host '  4) Small      - 720p, smallest'
Write-Host '  5) Custom     - fit under a size you choose (MB)'
$choice = if ($env:VC_PRESET) { $env:VC_PRESET } else { Read-Host 'Choose 1-5 [2]' }
if (-not $choice) { $choice = '2' }
$targetMb = 0
switch ($choice) {
  '1' { $targetMb = 25 }
  '3' { $crf = 22; $maxShort = 0 }
  '4' { $crf = 30; $maxShort = 720 }
  '5' {
    $v = if ($env:VC_MB) { $env:VC_MB } else { Read-Host 'Max size in MB' }
    $targetMb = [double]($v -replace '[^\d.]', '')
    if ($targetMb -le 0) { Write-Host 'Invalid size.' -ForegroundColor Red; return }
  }
  default { $crf = 26; $maxShort = 1080 }
}

function Get-OutPath($f) {
  $base = Join-Path $f.DirectoryName ($f.BaseName + '_compressed')
  $out = "$base.mp4"; $n = 1
  while (Test-Path -LiteralPath $out) { $out = "$base ($n).mp4"; $n++ }
  return $out
}
function Fmt($b) { if ($b -ge 1GB) { '{0:N2} GB' -f ($b / 1GB) } elseif ($b -ge 1MB) { '{0:N1} MB' -f ($b / 1MB) } else { '{0:N0} KB' -f ($b / 1KB) } }
function Probe($path, $entries, $sel) {
  $ErrorActionPreference = 'Continue'   # PS 5.1 turns redirected stderr into errors under 'Stop'
  $a = @('-v', 'error') + $sel + @('-show_entries', $entries, '-of', 'default=nw=1:nk=1', $path)
  $r = & $ffprobe @a 2>$null
  return @($r | Where-Object { $_ -and $_ -ne 'N/A' })
}

$ok = 0; $fail = 0
$nul = if ($IsLinux -or $IsMacOS) { '/dev/null' } else { 'NUL' }
foreach ($f in $files) {
  $out = Get-OutPath $f
  Write-Host ''
  Write-Host "Compressing $($f.Name) ($(Fmt $f.Length))" -ForegroundColor White
  try {
    $dur = 0.0; $d = Probe $f.FullName 'format=duration' @(); if ($d) { $dur = [double]$d[0] }
    $hasAudio = [bool](Probe $f.FullName 'stream=index' @('-select_streams', 'a:0'))
    $fpsRaw = Probe $f.FullName 'stream=avg_frame_rate' @('-select_streams', 'v:0')
    $fps = 0.0
    if ($fpsRaw -and $fpsRaw[0] -match '^(\d+)/(\d+)$' -and [double]$Matches[2] -gt 0) { $fps = [double]$Matches[1] / [double]$Matches[2] }

    $rate = @(); $twoPass = $false; $short = $maxShort
    if ($targetMb -gt 0) {
      if ($dur -le 0) { throw 'Could not read the video length.' }
      $aBps = if ($hasAudio) { 128000 } else { 0 }
      $vBps = [int][Math]::Max(100000, ($targetMb * 1e6 * 8 * 0.95 / $dur) - $aBps)
      # Lower the resolution automatically when the bitrate is too small for 1080p.
      $short = if ($vBps -lt 700000) { 480 } elseif ($vBps -lt 1800000) { 720 } else { 1080 }
      $rate = @('-b:v', "$vBps", '-maxrate', "$([int]($vBps * 1.5))", '-bufsize', "$($vBps * 2)")
      $twoPass = $true
    } else {
      $rate = @('-crf', "$crf")
    }

    $vf = @()
    if ($short -gt 0) {
      # Cap the SHORT side (works for portrait phone videos too); keep sizes even for H.264.
      $vf += "scale=w='if(gt(iw,ih),-2,trunc(min($short,iw)/2)*2)':h='if(gt(iw,ih),trunc(min($short,ih)/2)*2,-2)'"
    } else {
      $vf += "scale=trunc(iw/2)*2:trunc(ih/2)*2"
    }
    if ($fps -gt 31) { $vf += 'fps=30' }
    $common = @('-hide_banner', '-loglevel', 'error', '-stats', '-y', '-i', $f.FullName,
                '-map', '0:v:0', '-vf', ($vf -join ','),
                '-c:v', 'libx264', '-preset', 'veryfast') + $rate + @('-pix_fmt', 'yuv420p')

    if ($twoPass) {
      $log = Join-Path ([IO.Path]::GetTempPath()) ("us-sm-vc-" + [guid]::NewGuid().ToString('N'))
      try {
        Write-Host '  pass 1/2' -ForegroundColor DarkGray
        $a1 = $common + @('-pass', '1', '-passlogfile', $log, '-an', '-f', 'mp4', $nul)
        & $ffmpeg @a1
        if ($LASTEXITCODE -ne 0) { throw "ffmpeg failed (pass 1, exit $LASTEXITCODE)" }
        Write-Host '  pass 2/2' -ForegroundColor DarkGray
        $a2 = $common + @('-pass', '2', '-passlogfile', $log, '-map', '0:a:0?', '-c:a', 'aac', '-b:a', '128k', '-ac', '2', '-movflags', '+faststart', $out)
        & $ffmpeg @a2
        if ($LASTEXITCODE -ne 0) { throw "ffmpeg failed (exit $LASTEXITCODE)" }
      } finally { Get-ChildItem -Path "$log*" -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue }
    } else {
      $a = $common + @('-map', '0:a:0?', '-c:a', 'aac', '-b:a', '128k', '-ac', '2', '-movflags', '+faststart', $out)
      & $ffmpeg @a
      if ($LASTEXITCODE -ne 0) { throw "ffmpeg failed (exit $LASTEXITCODE)" }
    }

    $after = (Get-Item -LiteralPath $out).Length
    $pct = [int][Math]::Round((1 - $after / $f.Length) * 100)
    Write-Host "  OK  $(Fmt $f.Length) -> $(Fmt $after)  ($pct% smaller)" -ForegroundColor Green
    Write-Host "  Saved: $out" -ForegroundColor DarkGray
    if ($targetMb -gt 0 -and $after -gt $targetMb * 1e6) { Write-Host "  Note: slightly over $targetMb MB." -ForegroundColor Yellow }
    $ok++
  } catch {
    Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -ErrorAction SilentlyContinue }
    $fail++
  }
}

Write-Host ''
Write-Host "Done: $ok compressed, $fail failed. New files end with _compressed.mp4, next to the originals." -ForegroundColor Cyan
