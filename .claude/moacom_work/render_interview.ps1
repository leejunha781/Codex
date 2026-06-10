# Render PDF pages to PNG via WinRT (Windows PowerShell 5.1 required)
$pdfPath = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_전문면접_기술이론_전략_수식포함_v2.pdf"
$outDir  = "C:\Users\namma\.claude\moacom_work\png_interview"
$pages   = @(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19)   # zero-based page indices

if ($PSVersionTable.PSVersion.Major -gt 5) { Write-Output "ERROR: requires Windows PowerShell 5.1"; exit 1 }
New-Item -ItemType Directory -Force $outDir | Out-Null

Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
function AwaitAction($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)) ([Windows.Storage.StorageFile])
$pdfDoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
Write-Output "PageCount: $($pdfDoc.PageCount)"

foreach ($i in $pages) {
    if ($i -ge $pdfDoc.PageCount) { continue }
    $page = $pdfDoc.GetPage([uint32]$i)
    $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = [uint32]1240
    AwaitAction ($page.RenderToStreamAsync($stream, $opts))
    $size = $stream.Size
    $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
    Await ($reader.LoadAsync([uint32]$size)) ([uint32]) | Out-Null
    $bytes = New-Object byte[] $size
    $reader.ReadBytes($bytes)
    $png = Join-Path $outDir ("page{0:d2}.png" -f ($i+1))
    [System.IO.File]::WriteAllBytes($png, $bytes)
    Write-Output "Wrote $png"
    $reader.Dispose(); $stream.Dispose(); $page.Dispose()
}
