<#
  Builds dist\MoveTA.exe (single portable file) from Move-TA.ps1 + usp_IPS_MoveTA.sql.
  Uses only what ships with Windows: csc.exe from .NET Framework 4.x and WPF to draw the icon.
  Run:    powershell -NoProfile -ExecutionPolicy Bypass -STA -File build\Build-Exe.ps1
#>
$ErrorActionPreference = 'Stop'
$build = Split-Path -Parent $MyInvocation.MyCommand.Path
$root  = Split-Path -Parent $build
$dist  = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null

# ---------------- icon (PNG inside ICO, 256/48/32/16)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
function New-IconPng([int]$Size) {
    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()
    $r = [double]$Size * 0.19
    $bg = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(0x0F, 0x6C, 0xBD))
    $dc.DrawRoundedRectangle($bg, $null, (New-Object System.Windows.Rect(0, 0, $Size, $Size)), $r, $r)
    $tf = New-Object System.Windows.Media.Typeface('Segoe Fluent Icons')
    $ft = New-Object System.Windows.Media.FormattedText([string][char]0xE8DE, [Globalization.CultureInfo]::InvariantCulture, [System.Windows.FlowDirection]::LeftToRight, $tf, ([double]$Size * 0.58), [System.Windows.Media.Brushes]::White, 1.0)
    $dc.DrawText($ft, (New-Object System.Windows.Point((($Size - $ft.Width) / 2), (($Size - $ft.Height) / 2))))
    $dc.Close()
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($Size, $Size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    if ($Size -eq 256) {
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fsP = [IO.File]::Create((Join-Path $build 'MoveTA_256.png')); $enc.Save($fsP); $fsP.Close()
    }
    # BGRA without premultiplication, rows bottom-up (ICO DIB format)
    $conv = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap($rtb, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0)
    $stride = $Size * 4; $px = New-Object byte[] ($stride * $Size); $conv.CopyPixels($px, $stride, 0)
    $ms = New-Object IO.MemoryStream; $w = New-Object IO.BinaryWriter($ms)
    $w.Write([uint32]40); $w.Write([int32]$Size); $w.Write([int32]($Size * 2)); $w.Write([uint16]1); $w.Write([uint16]32)
    $w.Write([uint32]0); $w.Write([uint32]0); $w.Write([int32]0); $w.Write([int32]0); $w.Write([uint32]0); $w.Write([uint32]0)
    for ($y = $Size - 1; $y -ge 0; $y--) { $w.Write($px, $y * $stride, $stride) }
    $maskRow = [int]([Math]::Ceiling($Size / 32.0) * 4); $w.Write((New-Object byte[] ($maskRow * $Size)))
    $w.Flush(); $ms.ToArray()
}
$sizes = @(256, 48, 32, 16)
$imgs = @(); foreach ($s in $sizes) { $imgs += ,(New-IconPng $s) }
$ico = Join-Path $build 'MoveTA.ico'
$fs = [IO.File]::Create($ico); $bw = New-Object IO.BinaryWriter($fs)
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $len = $imgs[$i].Length
    $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s }))); $bw.Write([byte]($(if ($s -ge 256) { 0 } else { $s })))
    $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$len); $bw.Write([uint32]$offset); $offset += $len
}
foreach ($p in $imgs) { $bw.Write([byte[]]$p) }
$bw.Close()

# ---------------- compile
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$out = Join-Path $dist 'MoveTA.exe'
$args = @('/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', "/out:$out", "/win32icon:$ico",
          '/reference:System.Windows.Forms.dll',
          "/resource:$(Join-Path $root 'Move-TA.ps1'),Move-TA.ps1",
          "/resource:$(Join-Path $root 'usp_IPS_MoveTA.sql'),usp_IPS_MoveTA.sql",
          (Join-Path $build 'MoveTA.Launcher.cs'))
& $csc @args
if ($LASTEXITCODE -ne 0) { throw "csc falhou ($LASTEXITCODE)" }
Get-Item $out | Select-Object FullName, Length, LastWriteTime
