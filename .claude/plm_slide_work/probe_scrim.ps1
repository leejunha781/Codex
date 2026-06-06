$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)

function Dump($idx){
  $s = $pres.Slides.Item($idx)
  Write-Output "==================== SLIDE $idx (shapes=$($s.Shapes.Count)) ===================="
  foreach($sh in $s.Shapes){
    $t=""
    try { if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text } } catch {}
    $t = ($t -replace "[\r\n]"," ")
    if ($t.Length -gt 40){ $t=$t.Substring(0,40) }
    $kind = "shp$($sh.Type)"
    if ($sh.Type -eq 13){ $kind="PICTURE" }
    if ($sh.Type -eq 6){ $kind="GROUP" }
    $line = "z={0,3} {1,-9} {2,-16} L={3,4} T={4,4} W={5,4} H={6,4} | '{7}'" -f `
      $sh.ZOrderPosition, $kind, $sh.Name, [int]$sh.Left, [int]$sh.Top, [int]$sh.Width, [int]$sh.Height, $t
    Write-Output $line
  }
}
Dump 2
Dump 5
Dump 7
Dump 16
Dump 30
$pres.Close()
if ($ppt.Presentations.Count -eq 0){ $ppt.Quit() }
Write-Output "PROBE2 DONE"
