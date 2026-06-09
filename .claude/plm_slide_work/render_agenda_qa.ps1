$ErrorActionPreference = "Stop"
$pptx = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Technical_Meeting_Agenda_EN.pptx"
$outdir = "C:\Users\namma\.claude\plm_slide_work\agenda_qa_render"
New-Item -ItemType Directory -Force -Path $outdir | Out-Null
$pp = $null
for ($a=0; $a -lt 6 -and $pp -eq $null; $a++) { try { $pp = New-Object -ComObject PowerPoint.Application } catch { Start-Sleep -Milliseconds 1000 } }
$pres = $pp.Presentations.Open($pptx, $true, $false, $false)
for ($n=1; $n -le $pres.Slides.Count; $n++) {
    $f = Join-Path $outdir ("agenda_{0}.png" -f $n)
    $pres.Slides.Item($n).Export($f, "PNG", 1600, 900)
}
$pres.Close()
$pp.Quit()
Write-Output ("RENDERED " + $pres.Slides.Count + " -> " + $outdir)
