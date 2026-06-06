$ErrorActionPreference = "Stop"
$dir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$ko = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$koName = "Future_Industrial_PLM_Meeting_Deck_KO_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_ko"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$pil = [char]0x00B6

# build map from the 3 UTF-8 files
$map = @{}
foreach($f in @("build9_map.txt","build9_map2.txt","build9_map3.txt")){
  $path = "C:\Users\namma\.claude\plm_slide_work\$f"
  foreach($line in (Get-Content -LiteralPath $path -Encoding UTF8)){
    if($line.Length -lt 1){ continue }
    $idx = $line.IndexOf("|||")
    if($idx -lt 1){ continue }
    $k = $line.Substring(0,$idx); $v = $line.Substring($idx+3)
    $map[$k] = $v
  }
}
Write-Output ("MAP_ENTRIES=" + $map.Count)

$ppt = New-Object -ComObject PowerPoint.Application
foreach($p in @($ppt.Presentations)){ if($p.Name -eq $koName){ try{ $p.Saved=$true; $p.Close() }catch{} } }
$pres = $ppt.Presentations.Open($ko,0,0,0)

$script:hits = 0; $script:misses = @{}
function ApplyText($tr){
  try{
    $t = $tr.Text
    if($t.Trim().Length -lt 1){ return }
    $k = $t.Replace([string][char]13,[string]$pil).Replace([string][char]10,[string]$pil)
    if($map.ContainsKey($k)){
      $new = $map[$k].Replace([string]$pil,[string][char]13)
      $tr.Text = $new
      $script:hits++
    } else {
      if(-not $script:misses.ContainsKey($k)){ $script:misses[$k]=1 }
    }
  }catch{}
}
function ApplyShape($sh){
  try{
    if($sh.Type -eq 6){ foreach($c in $sh.GroupItems){ ApplyShape $c }; return }
    if($sh.HasTable -eq -1){ $tb=$sh.Table; for($r=1;$r -le $tb.Rows.Count;$r++){ for($c=1;$c -le $tb.Columns.Count;$c++){ ApplyText $tb.Cell($r,$c).Shape.TextFrame.TextRange } }; return }
    if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ ApplyText $sh.TextFrame.TextRange }
  }catch{}
}
for($i=1;$i -le $pres.Slides.Count;$i++){ foreach($sh in @($pres.Slides.Item($i).Shapes)){ ApplyShape $sh } }

Write-Output ("HITS=" + $script:hits + "  UNIQUE_MISSES=" + $script:misses.Count)
Write-Output "--- sample unmatched (kept as-is) ---"
$n=0; foreach($k in $script:misses.Keys){ if($n -ge 60){break}; Write-Output ("  · " + ($k -replace [string]$pil," / ")); $n++ }

$pres.Save()
for($i=1;$i -le $pres.Slides.Count;$i++){ $num="{0:D2}" -f $i; $pres.Slides.Item($i).Export("$outDir\slide-$num.png","PNG",1280,720) }
$pres.Close()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "APPLY_DONE"
