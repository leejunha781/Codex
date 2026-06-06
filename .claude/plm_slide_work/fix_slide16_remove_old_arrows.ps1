$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$renderDir = "C:\Users\namma\.claude\plm_slide_work\v4_slide16_visualpanel_render"
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null

try {
    $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
    $pp = New-Object -ComObject PowerPoint.Application
}
$pp.Visible = $true

$pres = $null
foreach ($p in $pp.Presentations) {
    if ($p.FullName -eq $file) { $pres = $p; break }
}
if (-not $pres) {
    $pres = $pp.Presentations.Open($file, $false, $false, $true)
}

$slide = $pres.Slides.Item(16)
foreach ($name in @("Right Arrow 6", "Right Arrow 7")) {
    try {
        $slide.Shapes.Item($name).Delete()
        Write-Output "DELETED=$name"
    } catch {
        Write-Output "NOTFOUND=$name"
    }
}

$png = Join-Path $renderDir "slide16_visual_panel_final.png"
$pres.Save()
$slide.Export($png, "PNG", 1600, 900)

try {
    $pres.Saved = $true
    $pres.Close()
} catch {}
$pres = $pp.Presentations.Open($file, $false, $false, $true)
$pp.ActiveWindow.View.GotoSlide(16)

Write-Output "SAVED=$file"
Write-Output "PNG=$png"
Write-Output "REOPENED=$($pres.Name) SLIDES=$($pres.Slides.Count) READONLY=$($pres.ReadOnly)"
