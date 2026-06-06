$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
$opened = $false
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $true, $false, $false); $opened = $true }

$terms = "linux","windows","host","embed","fastapi","e3d","marine","control plane","on-prem","data center","data-center","authoring","adapter","agent","server","cloud","connect","runs on","native"
$out = New-Object System.Text.StringBuilder
function Walk($shapes,$sidx){
  foreach($sh in $shapes){
    $t=""
    try { if($sh.HasTextFrame -and $sh.TextFrame.HasText){ $t=$sh.TextFrame.TextRange.Text } } catch {}
    if($t){
      $norm = ($t -replace "[\r\n\x0b]+"," | ").Trim()
      $low = $norm.ToLower()
      foreach($term in $terms){ if($low.Contains($term)){ [void]$out.AppendLine("s$sidx [$($sh.Name)]  $norm"); break } }
    }
    try { if($sh.HasTable){ $tb=$sh.Table; for($r=1;$r -le $tb.Rows.Count;$r++){ for($c=1;$c -le $tb.Columns.Count;$c++){ $ct=$tb.Cell($r,$c).Shape.TextFrame.TextRange.Text; if($ct){ $cl=$ct.ToLower(); foreach($term in $terms){ if($cl.Contains($term)){ [void]$out.AppendLine("s$sidx [tbl r$r c$c]  " + (($ct -replace "[\r\n\x0b]+"," ").Trim())); break } } } } } } } catch {}
    try { if($sh.Type -eq 6){ Walk $sh.GroupItems $sidx } } catch {}
  }
}
for($i=1;$i -le $pres.Slides.Count;$i++){ Walk $pres.Slides.Item($i).Shapes $i }
$outPath="C:\Users\namma\.claude\plm_slide_work\scan_terms.txt"
[System.IO.File]::WriteAllText($outPath,$out.ToString(),(New-Object System.Text.UTF8Encoding $false))
if($opened){ $pres.Close() }
Write-Output "WROTE $outPath"
