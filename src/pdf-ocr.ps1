param([Parameter(Mandatory = $true)][string]$PdfPath)

$ErrorActionPreference = "Stop"
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]

$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object {
    $_.Name -eq "AsTask" -and
    $_.IsGenericMethod -and
    $_.GetParameters().Count -eq 1
  } |
  Select-Object -First 1

$asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object {
    $_.Name -eq "AsTask" -and
    -not $_.IsGenericMethod -and
    $_.GetParameters().Count -eq 1
  } |
  Select-Object -First 1

function Await-Result {
  param($Operation, [Type]$ResultType)
  $task = $script:asTaskGeneric.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
  $task.Wait()
  return $task.Result
}

function Await-Action {
  param($Operation)
  $task = $script:asTaskAction.Invoke($null, @($Operation))
  $task.Wait()
}

function Get-OcrTextScore {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return 0
  }

  $score = 0
  foreach ($keyword in @(
    "账号编码",
    "账号名",
    "短信业务客户结算单",
    "短信类型",
    "成功条数",
    "对接人",
    "单价",
    "金额",
    "合计"
  )) {
    if ($Text.Contains($keyword)) {
      $score += 100 + $keyword.Length * 10
    }
  }

  $score += [regex]::Matches($Text, "(?:s|sn)_[a-zA-Z0-9_]+", "IgnoreCase").Count * 60
  $score += [Math]::Min(200, [regex]::Matches($Text, "[\u4e00-\u9fff]").Count) * 0.2
  return $score
}

function Get-OcrText {
  param(
    [Windows.Graphics.Imaging.BitmapDecoder]$Decoder,
    [Windows.Media.Ocr.OcrEngine]$Engine
  )

  $bestText = ""
  $bestScore = -1
  $rotations = @(
    [Windows.Graphics.Imaging.BitmapRotation]::None,
    [Windows.Graphics.Imaging.BitmapRotation]::Clockwise90Degrees,
    [Windows.Graphics.Imaging.BitmapRotation]::Clockwise180Degrees,
    [Windows.Graphics.Imaging.BitmapRotation]::Clockwise270Degrees
  )

  foreach ($rotation in $rotations) {
    $transform = [Windows.Graphics.Imaging.BitmapTransform]::new()
    $transform.Rotation = $rotation
    $bitmap = Await-Result (
      $Decoder.GetSoftwareBitmapAsync(
        [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
        [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
        $transform,
        [Windows.Graphics.Imaging.ExifOrientationMode]::IgnoreExifOrientation,
        [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage
      )
    ) ([Windows.Graphics.Imaging.SoftwareBitmap])

    try {
      $result = Await-Result ($Engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
      $score = Get-OcrTextScore $result.Text
      if ($score -gt $bestScore) {
        $bestText = $result.Text
        $bestScore = $score
      }
    } finally {
      $bitmap.Dispose()
    }
  }

  return $bestText
}

$file = Await-Result ([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)) ([Windows.Storage.StorageFile])
$document = Await-Result ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
$language = [Windows.Globalization.Language]::new("zh-CN")
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if (-not $engine) {
  throw "Windows OCR Chinese language is unavailable."
}

$parts = @()
for ($index = 0; $index -lt $document.PageCount; $index++) {
  $page = $document.GetPage($index)
  $stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
  try {
    $options = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
    $options.DestinationWidth = [uint32]2400
    Await-Action ($page.RenderToStreamAsync($stream, $options))
    $stream.Seek(0)
    $decoder = Await-Result ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $parts += Get-OcrText -Decoder $decoder -Engine $engine
  } finally {
    $stream.Dispose()
    $page.Dispose()
  }
}

$parts -join "`n"
