$ErrorActionPreference = "Stop"
$pdf = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.pdf"
$outdir = "C:\Users\namma\.claude\plm_slide_work\guide_qa_render"
New-Item -ItemType Directory -Force -Path $outdir | Out-Null

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op,$t){ $task=$asTaskGeneric.MakeGenericMethod($t).Invoke($null,@($op)); $task.Wait(-1)|Out-Null; $task.Result }
$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function AwaitAction($act){ $task=$asTaskAction.Invoke($null,@($act)); $task.Wait(-1)|Out-Null }

[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.DataReader,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
$pdoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
Write-Output ("PDF pages = " + $pdoc.PageCount)

$targets = @(0,3,7,9)   # 0-based: pages 1,4,8,10
foreach($idx in $targets){
    if($idx -ge $pdoc.PageCount){ continue }
    $page = $pdoc.GetPage([uint32]$idx)
    $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = [uint32]1400
    AwaitAction ($page.RenderToStreamAsync($stream, $opts))
    $size = [uint32]$stream.Size
    $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
    Await ($reader.LoadAsync($size)) ([uint32]) | Out-Null
    $bytes = New-Object byte[] $size
    $reader.ReadBytes($bytes)
    $out = Join-Path $outdir ("guide_p{0:00}.png" -f ($idx+1))
    [System.IO.File]::WriteAllBytes($out, $bytes)
    $reader.Dispose(); $stream.Dispose(); $page.Dispose()
    Write-Output ("wrote " + $out)
}
