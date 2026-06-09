$ErrorActionPreference = "Stop"
$v16 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V16.pptx"
$outdir = "C:\Users\namma\.claude\plm_slide_work\v16_qa_render"
New-Item -ItemType Directory -Force -Path $outdir | Out-Null
$pp=$null
for($a=0;$a -lt 6 -and $pp -eq $null;$a++){ try{$pp=New-Object -ComObject PowerPoint.Application}catch{Start-Sleep -Milliseconds 1000} }
$pres=$pp.Presentations.Open($v16,$true,$false,$false)
foreach($n in @(1,2,3,4,19,32,42)){
  $pres.Slides.Item($n).Export((Join-Path $outdir ("v16_{0:00}.png" -f $n)),"PNG",1600,900)
}
$pres.Close(); $pp.Quit()
Write-Output ("done -> " + $outdir)
