$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output "FILE: $file"
Write-Output "EXISTS: $(Test-Path $file)"
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
$opened=$false
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $true, $false, $false); $opened=$true }

$out = New-Object System.Text.StringBuilder
function Walk($shapes,$sb,$indent){
  foreach($sh in $shapes){
    $t=""
    try { if($sh.HasTextFrame -and $sh.TextFrame.HasText){ $t=$sh.TextFrame.TextRange.Text } } catch {}
    if($t){ $t=($t -replace "[\r\n\x0b]+"," / ").Trim(); if($t.Length -gt 0){ [void]$sb.AppendLine("$indent[$($sh.Name)] $t") } }
    try { if($sh.HasTable){ $tb=$sh.Table; for($r=1;$r -le $tb.Rows.Count;$r++){ $row=@(); for($c=1;$c -le $tb.Columns.Count;$c++){ $ct=$tb.Cell($r,$c).Shape.TextFrame.TextRange.Text; $row+=(($ct -replace "[\r\n\x0b]+"," ").Trim()) }; [void]$sb.AppendLine("$indent  TBLROW: " + ($row -join " | ")) } } } catch {}
    try { if($sh.Type -eq 6){ Walk $sh.GroupItems $sb "$indent  " } } catch {}
  }
}
$n=$pres.Slides.Count
for($i=1;$i -le $n;$i++){ [void]$out.AppendLine(""); [void]$out.AppendLine("========== SLIDE $i =========="); Walk $pres.Slides.Item($i).Shapes $out "" }
$outPath="C:\Users\namma\.claude\plm_slide_work\dump_v4.txt"
[System.IO.File]::WriteAllText($outPath,$out.ToString(),(New-Object System.Text.UTF8Encoding $false))
if($opened){ $pres.Close() }
Write-Output "SLIDES: $n"
Write-Output "WROTE: $outPath"
