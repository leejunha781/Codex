$file = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
$backupDir = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\_backup_20260606_v4_slide15_windows_card"
$renderDir = "C:\Users\namma\.claude\plm_slide_work\v4_slide15_windows_card_render"
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null

function RgbInt([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

if (-not (Test-Path -LiteralPath $file)) { throw "Missing PPTX: $file" }

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

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4_BEFORE_SLIDE15_WINDOWS_CARD_{0}.pptx" -f $stamp)
$pres.SaveCopyAs($backup)
Write-Output "BACKUP=$backup"

$slide = $pres.Slides.Item(15)

# Remove only prior versions of this additive card if the script is re-run.
for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
    $sh = $slide.Shapes.Item($i)
    if ($sh.Name -eq "WindowsAuthoringTierCard" -or $sh.Name -eq "WindowsAuthoringTierIcon") {
        $sh.Delete()
    }
}

# Convert the existing bottom free-text note into a proper visible card by hiding the old
# note text and preserving the same content inside the new card. Other infrastructure
# objects remain untouched.
$oldNote = $null
try { $oldNote = $slide.Shapes.Item("TextBox 47") } catch {}
if ($oldNote) {
    $oldNote.TextFrame.TextRange.Text = ""
}

$cardText = "Windows Authoring Tier`rGPU/Citrix workstations run AVEVA E3D · Marine · Hull · Draw in supported AVEVA environments`rLocal Agent / Plugin / PML.NET Bridge connects to Linux Control Plane — not hosted on Linux"

$card = $slide.Shapes.AddShape(5, 36, 474, 476, 42)
$card.Name = "WindowsAuthoringTierCard"
$card.Fill.Solid()
$card.Fill.ForeColor.RGB = RgbInt 10 42 62
$card.Fill.Transparency = [single]0.03
$card.Line.Visible = -1
$card.Line.ForeColor.RGB = RgbInt 76 214 156
$card.Line.Weight = [single]1.2
$card.TextFrame.TextRange.Text = $cardText
$card.TextFrame.TextRange.Font.Name = "Aptos"
$card.TextFrame.TextRange.Font.Size = [single]7.4
$card.TextFrame.TextRange.Font.Color.RGB = RgbInt 221 236 248
$card.TextFrame.MarginLeft = 34
$card.TextFrame.MarginRight = 8
$card.TextFrame.MarginTop = 3
$card.TextFrame.MarginBottom = 2
try { $card.TextFrame2.AutoSize = 0 } catch {}
try { $card.TextFrame2.WordWrap = -1 } catch {}

try {
    $card.TextFrame.TextRange.Paragraphs(1).Font.Bold = [int]1
    $card.TextFrame.TextRange.Paragraphs(1).Font.Size = [single]9.3
    $card.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RgbInt 76 214 156
} catch {}

$icon = $slide.Shapes.AddShape(9, 48, 487, 12, 12)
$icon.Name = "WindowsAuthoringTierIcon"
$icon.Fill.Solid()
$icon.Fill.ForeColor.RGB = RgbInt 76 214 156
$icon.Line.Visible = 0

# Keep the card above the infrastructure panel but below the global footer line.
$card.ZOrder(0) | Out-Null
$icon.ZOrder(0) | Out-Null
try { $slide.Shapes.Item("TextBox 92").ZOrder(0) | Out-Null } catch {}
try { $slide.Shapes.Item("Connector 30").ZOrder(0) | Out-Null } catch {}

$png = Join-Path $renderDir "slide15_windows_authoring_card.png"
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
