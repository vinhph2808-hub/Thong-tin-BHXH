param (
    [string]$ListFile
)

$Images = Get-Content -LiteralPath $ListFile -Encoding UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { 
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' 
}[0]

function Await($winRtTask, $resultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($resultType)
    $netTask = $asTask.Invoke($null, @($winRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}

[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime] | Out-Null

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) {
    $lang = [Windows.Globalization.Language]::new('en-US')
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
}

$allResults = [System.Collections.Generic.List[Object]]::new()

foreach ($imgPath in $Images) {
    $imgPathStr = "$imgPath".Trim()
    if ([string]::IsNullOrWhiteSpace($imgPathStr)) { continue }
    if (-not (Test-Path $imgPathStr)) { continue }
    $absPath = (Resolve-Path $imgPathStr).Path
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($absPath)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

    $words = [System.Collections.Generic.List[Object]]::new()
    foreach ($line in $result.Lines) {
        foreach ($w in $line.Words) {
            $words.Add(@{
                t = $w.Text
                x = [int]$w.BoundingRect.X
                y = [int]$w.BoundingRect.Y
                w = [int]$w.BoundingRect.Width
                h = [int]$w.BoundingRect.Height
            })
        }
    }

    $allResults.Add(@{
        image = $imgPathStr
        words = $words
    })
}

ConvertTo-Json -InputObject @($allResults) -Depth 5 -Compress
