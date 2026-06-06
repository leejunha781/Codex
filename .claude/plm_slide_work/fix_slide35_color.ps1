$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres=$p; break } }
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $false, $false, $true) }
Write-Output ("slides=" + $pres.Slides.Count + "  RO=" + $pres.ReadOnly)

$WHITE = [int]16777215   # 0xFFFFFF
$TEAL  = [int]3394713    # 0x33CF99 - AVEVA accent teal: R=51 G=207 B=153 -> 0x33CF99

$ns = $pres.Slides.Item(35)
Write-Output ("slide35 shapes: " + $ns.Shapes.Count)

foreach ($sh in $ns.Shapes) {
    try {
        if ($sh.HasTextFrame) {
            $name = $sh.Name
            Write-Output ("  fix color: [" + $name + "]")
            # Set all text white first
            $sh.TextFrame.TextRange.Font.Color.RGB = $WHITE
            # Bold labels - make them teal/accent
            if ($name -eq "AILabel1" -or $name -eq "AILabel2") {
                $sh.TextFrame.TextRange.Font.Color.RGB = $TEAL
            }
            # Title - keep white but slightly larger bold
            if ($name -eq "AITitle") {
                $sh.TextFrame.TextRange.Font.Color.RGB = $WHITE
            }
            # Subtitle - slightly dimmer (light grey)
            if ($name -eq "AISubtitle") {
                $sh.TextFrame.TextRange.Font.Color.RGB = [int]0xCCCCCC
            }
        }
    } catch { Write-Output ("  ERR on [" + $sh.Name + "]: " + $_) }
}

# Also bold the section headings inside content text boxes
# Find paragraphs starting with "WHAT IT IS:", "USE CASES:", "ADVANTAGE:", "CONSTRAINT:"
$headings = @("WHAT IT IS:", "USE CASES:", "ADVANTAGE:", "CONSTRAINT:")
foreach ($shName in @("AILeft","AIRight")) {
    $sh = $null
    foreach ($s in $ns.Shapes) { if ($s.Name -eq $shName) { $sh=$s; break } }
    if ($sh) {
        $nPara = $sh.TextFrame.TextRange.Paragraphs().Count
        for ($pi=1; $pi -le $nPara; $pi++) {
            try {
                $pt = $sh.TextFrame.TextRange.Paragraphs($pi).Text.Trim()
                $isHeading = $false
                foreach ($hd in $headings) { if ($pt -like ($hd+"*") -or $pt -eq $hd) { $isHeading=$true; break } }
                if ($isHeading) {
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Bold = [int]1
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Color.RGB = $TEAL
                    Write-Output ("    bold+teal: [" + $shName + "] p$pi '" + $pt + "'")
                }
            } catch {}
        }
    }
}

$pres.Save()
Write-Output "SAVED"

# Re-render slide 35
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
$out = Join-Path $rDir "slide35.png"
$ns.Export($out, "PNG", 1440, 810)
Write-Output ("RENDERED: " + $out)

try { $ppt.ActiveWindow.View.GotoSlide(35) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "DONE"
