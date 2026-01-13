
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.RandomAccessStream, Windows.Foundation, ContentType = WindowsRuntime]

function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | 
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and 
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

$imagePath = $args[0]
$stream = [System.IO.File]::OpenRead($imagePath)
$randomAccessStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($stream)

$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($randomAccessStream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$softwareBitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($ocrEngine -eq $null) {
    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage("en-US")
}

$ocrResult = Await ($ocrEngine.RecognizeAsync($softwareBitmap)) ([Windows.Media.Ocr.OcrResult])

$results = @()
foreach ($line in $ocrResult.Lines) {
    foreach ($word in $line.Words) {
        $rect = $word.BoundingRect
        $results += @{
            text = $word.Text
            left = [int]$rect.X
            top = [int]$rect.Y
            width = [int]$rect.Width
            height = [int]$rect.Height
        }
    }
}

$stream.Close()
ConvertTo-Json -InputObject $results -Compress
