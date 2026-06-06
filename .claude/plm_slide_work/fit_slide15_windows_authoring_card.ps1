$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$renderDir = "C:\Users\namma\.claude\plm_slide_work\v4_slide15_windows_card_render"
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null

function RgbInt([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

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

$slide = $pres.Slides.Item(15)

# Reduce only the bottom infrastructure card heights to make room for a legible Windows authoring card.
foreach ($name in @("Rounded Rectangle 31", "Rounded Rectangle 35", "Rounded Rectangle 39", "Rounded Rectangle 43")) {
    try {
        $sh = $slide.Shapes.Item($name)
        $sh.Height = [single]82
    } catch {}
}

try {
    $card = $slide.Shapes.Item("WindowsAuthoringTierCard")
    $card.Left = 36
    $card.Top = 464
    $card.Width = 476
    $card.Height = 48
    $card.TextFrame.TextRange.Text = "Windows Authoring Tier`rGPU/Citrix workstations: AVEVA E3D · Marine · Hull · Draw`rLocal Agent / Plugin / PML.NET Bridge → Linux Control Plane; not hosted on Linux"
    $card.TextFrame.TextRange.Font.Name = "Aptos"
    $card.TextFrame.TextRange.Font.Size = [single]8.2
    $card.TextFrame.TextRange.Font.Color.RGB = RgbInt 221 236 248
    $card.TextFrame.MarginLeft = 34
    $card.TextFrame.MarginRight = 8
    $card.TextFrame.MarginTop = 4
    $card.TextFrame.MarginBottom = 3
    $card.Line.ForeColor.RGB = RgbInt 76 214 156
    $card.Line.Weight = [single]1.35
    try {
        $card.TextFrame.TextRange.Paragraphs(1).Font.Bold = [int]1
        $card.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]9.8
        $card.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RgbInt 76 214 156
    } catch {}
} catch {}

try {
    $icon = $slide.Shapes.Item("WindowsAuthoringTierIcon")
    $icon.Left = 48
    $icon.Top = 481
    $icon.Width = 13
    $icon.Height = 13
} catch {}

$png = Join-Path $renderDir "slide15_windows_authoring_card_final.png"
$pres.Save()
$slide.Export($png, "PNG", 1600, 900)

try {
    $pres.Saved = $true
    $pres.Close()
} catch {}
$pres = $pp.Presentations.Open($file, $false, $false, $true)
$pp.ActiveWindow.View.GotoSlide(15)

Write-Output "SAVED=$file"
Write-Output "PNG=$png"
Write-Output "REOPENED=$($pres.Name) SLIDES=$($pres.Slides.Count) READONLY=$($pres.ReadOnly)"
