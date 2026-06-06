$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$img = "C:\Users\namma\.claude\plm_slide_work\img17_hybrid_boundary.png"
$backupDir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\_backup_20260606_v4_visualsupport"
$renderDir = "C:\Users\namma\.claude\plm_slide_work\v4_visualsupport_render"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null

if (-not (Test-Path -LiteralPath $file)) { throw "Missing PPTX: $file" }
if (-not (Test-Path -LiteralPath $img)) { throw "Missing image: $img" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4_BEFORE_VISUALSUPPORT_{0}.pptx" -f $stamp)
Copy-Item -LiteralPath $file -Destination $backup
Write-Output "BACKUP=$backup"

function RgbInt([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

$pp = $null
$pres = $null

try {
    try {
        $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        $pp = New-Object -ComObject PowerPoint.Application
    }
    $pp.Visible = $true

    foreach ($p in $pp.Presentations) {
        if ($p.FullName -eq $file) {
            $pres = $p
            break
        }
    }
    if (-not $pres) {
        $pres = $pp.Presentations.Open($file, $false, $false, $true)
    }

    $slide = $pres.Slides.Item(16)

    # Use the generated visual as a slide background, then add a dark scrim behind all editable content.
    # This preserves the slide's cards/table as editable PowerPoint objects.
    $slide.FollowMasterBackground = $false
    $slide.Background.Fill.UserPicture($img)

    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $sh = $slide.Shapes.Item($i)
        if ($sh.Name -eq "GeneratedHybridBgScrim") {
            $sh.Delete()
        }
    }

    $scrim = $slide.Shapes.AddShape(1, 0, 0, $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight)
    $scrim.Name = "GeneratedHybridBgScrim"
    $scrim.Fill.Solid()
    $scrim.Fill.ForeColor.RGB = RgbInt 3 10 22
    $scrim.Fill.Transparency = [single]0.36
    $scrim.Line.Visible = 0
    $scrim.ZOrder(1) | Out-Null

    # Slightly strengthen the three top cards so the image reads as context, not noise.
    foreach ($idx in @(3,4,5)) {
        try {
            $card = $slide.Shapes.Item($idx)
            $card.Fill.Transparency = [single]0.05
        } catch {}
    }

    $pres.Save()
    Write-Output "SAVED=$file"

    $slide.Export((Join-Path $renderDir "slide16_visualsupport.png"), "PNG", 1600, 900)

    # Refresh the visible PowerPoint window from disk to avoid stale-window confusion.
    try {
        $pres.Saved = $true
        $pres.Close()
    } catch {}
    $pres = $pp.Presentations.Open($file, $false, $false, $true)
    $pp.ActiveWindow.View.GotoSlide(16)

    Write-Output "REOPENED=$($pres.Name) SLIDES=$($pres.Slides.Count) READONLY=$($pres.ReadOnly)"
    Write-Output "PNG=$(Join-Path $renderDir "slide16_visualsupport.png")"
} finally {
    # Keep PowerPoint open for the user.
}
