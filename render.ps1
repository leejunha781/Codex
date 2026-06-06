$ErrorActionPreference = 'Stop'
$pdf = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"
$out = "C:\Users\namma\.claude\itt_work"

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$wse = [System.WindowsRuntimeSystemExtensions].GetMethods()
$asTaskOp = ($wse | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$asTaskAct = ($wse | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function AwaitOp($op,$t){ $m=$asTaskOp.MakeGenericMethod($t); $task=$m.Invoke($null,@($op)); $task.Wait(-1)|Out-Null; $task.Result }
function AwaitAct($act){ $task=$asTaskAct.Invoke($null,@($act)); $task.Wait(-1)|Out-Null }

[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.DataReader,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null

$sf = AwaitOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
$pd = AwaitOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
Write-Output ("PAGECOUNT=" + $pd.PageCount)

$pages = @(0,1,7,8)
foreach ($i in $pages) {
  if ($i -ge $pd.PageCount) { continue }
  $pg = $pd.GetPage($i)
  $ras = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $opt.DestinationWidth = [uint32]1240
  AwaitAct ($pg.RenderToStreamAsync($ras, $opt))
  $size = [uint32]$ras.Size
  $reader = New-Object Windows.Storage.Streams.DataReader($ras.GetInputStreamAt(0))
  AwaitOp ($reader.LoadAsync($size)) ([uint32]) | Out-Null
  $bytes = New-Object byte[] $size
  $reader.ReadBytes($bytes)
  $reader.Dispose(); $ras.Dispose(); $pg.Dispose()
  [System.IO.File]::WriteAllBytes((Join-Path $out ("page" + ($i+1) + ".png")), $bytes)
  Write-Output ("WROTE page" + ($i+1) + ".png bytes=" + $size)
}
Write-Output "RENDER_DONE"
