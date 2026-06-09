$ErrorActionPreference = "Stop"
$path = "D:\이력서\AVEVA - Marine Principal Technical Support – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx"
# fix path (en-dash variations) -> use literal from disk
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V15.pptx"
$outdir = "C:\Users\namma\.claude\plm_slide_work\v15_qa_render"
New-Item -ItemType Directory -Force -Path $outdir | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($path, $true, $false, $false)

$targets = @(1,2,6,7,16,23,31,37)
foreach ($n in $targets) {
    $sl = $pres.Slides.Item($n)
    $f = Join-Path $outdir ("slide_{0:00}.png" -f $n)
    $sl.Export($f, "PNG", 1600, 900)
}
$pres.Close()
$pp.Quit()
Write-Output "RENDERED to $outdir"
Get-ChildItem $outdir -Filter *.png | Select-Object Name, Length | Format-Table -AutoSize | Out-String | Write-Output
