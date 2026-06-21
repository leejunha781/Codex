param([string]$pdf = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.pdf",
      [string]$out = "C:\Users\namma\.claude\genohco_system_work\res")
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskOp = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$asTaskAct = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.FullName -eq 'Windows.Foundation.IAsyncAction' })[0]
function Await($op,$t){ $task=$asTaskOp.MakeGenericMethod($t).Invoke($null,@($op)); $task.Wait(-1)|Out-Null; $task.Result }
function AwaitAction($a){ $task=$asTaskAct.Invoke($null,@($a)); $task.Wait(-1)|Out-Null }
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]|Out-Null
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]|Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]|Out-Null
[Windows.Storage.Streams.DataReader,Windows.Storage.Streams,ContentType=WindowsRuntime]|Out-Null
$file=Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
$pd=Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
Write-Output ("pages=" + $pd.PageCount)
for($i=0;$i -lt $pd.PageCount;$i++){
  $pg=$pd.GetPage([uint32]$i)
  $o=New-Object Windows.Data.Pdf.PdfPageRenderOptions; $o.DestinationWidth=[uint32]1240
  $st=New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitAction ($pg.RenderToStreamAsync($st,$o))
  $sz=[uint32]$st.Size
  $rd=New-Object Windows.Storage.Streams.DataReader($st.GetInputStreamAt(0))
  Await ($rd.LoadAsync($sz)) ([uint32]) | Out-Null
  $bytes=New-Object byte[] $sz; $rd.ReadBytes($bytes)
  [System.IO.File]::WriteAllBytes(($out+($i+1)+".png"),$bytes)
}
Write-Output "done"
