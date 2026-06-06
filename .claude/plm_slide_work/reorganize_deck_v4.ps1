$ErrorActionPreference = "Stop"

$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output ("FILE: " + $file)
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# Backup
$bkDir = Join-Path $proposal.FullName "Proposal\_backup_20260606_reorganize"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory $bkDir | Out-Null }
$bk = Join-Path $bkDir "v4_BEFORE_REORGANIZE.pptx"
Copy-Item $file $bk -Force
Write-Output ("BACKUP: " + $bk)

# COM attach
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

# Close any stale in-memory copy
$toClose = @()
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $toClose += $p } }
foreach ($p in $toClose) { try { $p.Saved = $true; $p.Close() } catch {} }

$pres = $ppt.Presentations.Open($file, $false, $false, $true)
Write-Output ("Opened: slides=" + $pres.Slides.Count + "  RO=" + $pres.ReadOnly)

$TEAL      = [int]3394713   # #33CC99 AVEVA teal
$CARDCOLOR = [int]1650015   # RGB(25,45,95) dark navy card

# ===========================
# STEP 1: GRAPHICS — SLIDE 35
# ===========================
Write-Output "=== STEP 1: Graphics slide 35 ==="
$sl = $pres.Slides.Item(35)

# Left column card (behind AILabel1 + AILeft)
$c = $sl.Shapes.AddShape(5, [single]14, [single]78, [single]462, [single]413)
$c.Name = "AICardLeft"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

# Right column card (behind AILabel2 + AIRight)
$c = $sl.Shapes.AddShape(5, [single]482, [single]78, [single]464, [single]413)
$c.Name = "AICardRight"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

# Thin teal accent bar separating title area from column cards
$b = $sl.Shapes.AddShape(1, [single]14, [single]74, [single]932, [single]2)
$b.Name = "AIAccentBar"
$b.Fill.Solid()
$b.Fill.ForeColor.RGB = $TEAL
$b.Line.Visible = [int]0
[void]$b.ZOrder(1)

# Thin teal dashed vertical divider between columns
$dv = $sl.Shapes.AddLine([single]478, [single]80, [single]478, [single]490)
$dv.Name = "AIDivider"
$dv.Line.ForeColor.RGB = $TEAL
$dv.Line.Weight = [single]0.75

Write-Output ("  s35 shapes now: " + $sl.Shapes.Count)

# ===========================
# STEP 2: GRAPHICS — SLIDE 36
# ===========================
Write-Output "=== STEP 2: Graphics slide 36 ==="
$sl = $pres.Slides.Item(36)

# Left column card (behind LWLabel1 + LWLeft)
$c = $sl.Shapes.AddShape(5, [single]14, [single]78, [single]472, [single]413)
$c.Name = "LWCardLeft"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

# Right column card (behind LWLabel2 + LWRight)
$c = $sl.Shapes.AddShape(5, [single]492, [single]78, [single]454, [single]413)
$c.Name = "LWCardRight"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

# Accent bar
$b = $sl.Shapes.AddShape(1, [single]14, [single]74, [single]932, [single]2)
$b.Name = "LWAccentBar"
$b.Fill.Solid()
$b.Fill.ForeColor.RGB = $TEAL
$b.Line.Visible = [int]0
[void]$b.ZOrder(1)

# Column divider
$dv = $sl.Shapes.AddLine([single]488, [single]80, [single]488, [single]490)
$dv.Name = "LWDivider"
$dv.Line.ForeColor.RGB = $TEAL
$dv.Line.Weight = [single]0.75

Write-Output ("  s36 shapes now: " + $sl.Shapes.Count)

# ===========================
# STEP 3: MOVE SLIDES 35-36 → 13-14
# After MoveTo(13): old 35 → pos 13; old 13-34 → pos 14-35; old 36 → pos 36
# After MoveTo(14): old 36 → pos 14; old 13-34 → pos 15-36
# ===========================
Write-Output "=== STEP 3: Move slides ==="
$pres.Slides.Item(35).MoveTo(13)
Write-Output ("  moved s35 -> pos 13  (total: " + $pres.Slides.Count + ")")
$pres.Slides.Item(36).MoveTo(14)
Write-Output ("  moved s36 -> pos 14  (total: " + $pres.Slides.Count + ")")

# ===========================
# STEP 4: UPDATE PAGE NUMBERS
# ===========================
Write-Output "=== STEP 4: Page numbers ==="

# Slide 13 (was s35): find text "35" → "13"
$found = $false
foreach ($sh in $pres.Slides.Item(13).Shapes) {
    try {
        if ($sh.HasTextFrame) {
            if ($sh.TextFrame.TextRange.Text.Trim() -eq "35") {
                $sh.TextFrame.TextRange.Text = "13"
                Write-Output ("  s13: 35->13 [" + $sh.Name + "]")
                $found = $true
                break
            }
        }
    } catch {}
}
if (-not $found) { Write-Output "  s13: MISS '35'" }

# Slide 14 (was s36): find text "36" → "14"
$found = $false
foreach ($sh in $pres.Slides.Item(14).Shapes) {
    try {
        if ($sh.HasTextFrame) {
            if ($sh.TextFrame.TextRange.Text.Trim() -eq "36") {
                $sh.TextFrame.TextRange.Text = "14"
                Write-Output ("  s14: 36->14 [" + $sh.Name + "]")
                $found = $true
                break
            }
        }
    } catch {}
}
if (-not $found) { Write-Output "  s14: MISS '36'" }

# Slides 15-36 (was s13-s34): each page number increments by 2
for ($si = 15; $si -le 36; $si++) {
    $oldPg = [string]($si - 2)
    $newPg = [string]$si
    $found = $false
    foreach ($sh in $pres.Slides.Item($si).Shapes) {
        try {
            if ($sh.HasTextFrame -and -not $sh.HasTable) {
                if ($sh.TextFrame.TextRange.Text.Trim() -eq $oldPg) {
                    $sh.TextFrame.TextRange.Text = $newPg
                    Write-Output ("  s" + $si + ": " + $oldPg + "->" + $newPg + " [" + $sh.Name + "]")
                    $found = $true
                    break
                }
            }
        } catch {}
    }
    if (-not $found) { Write-Output ("  s" + $si + ": MISS '" + $oldPg + "'") }
}

# ===========================
# STEP 5: SECTION LABELS 13-14
# Change "Appendix" footer chip to "AI STRATEGY"
# ===========================
Write-Output "=== STEP 5: Section labels ==="
foreach ($sh in $pres.Slides.Item(13).Shapes) {
    try {
        if ($sh.HasTextFrame -and $sh.Name -eq "AIAppendix") {
            $sh.TextFrame.TextRange.Text = "AI STRATEGY"
            Write-Output "  s13 AIAppendix -> AI STRATEGY"
        }
    } catch {}
}
foreach ($sh in $pres.Slides.Item(14).Shapes) {
    try {
        if ($sh.HasTextFrame -and $sh.Name -eq "LWAppendix") {
            $sh.TextFrame.TextRange.Text = "AI STRATEGY"
            Write-Output "  s14 LWAppendix -> AI STRATEGY"
        }
    } catch {}
}

# ===========================
# STEP 6: SLIDE 2 NAVIGATION
# Original ranges:
#   MeetingNavRange3 p.7-12   → p.7-14  (AI slides now part of Winning Ideas)
#   MeetingNavRange4 p.13-25  → p.15-27 (+2 shift)
#   MeetingNavRange5 p.26     → p.28
#   MeetingNavRange6 p.27-29  → p.29-31
#   MeetingNavRange7 p.30     → p.32
#   GlossaryPointer  p.31-33  → p.34-36
# ===========================
Write-Output "=== STEP 6: Slide 2 navigation ==="
$EN = [char]0x2013   # en-dash
$AR = [char]0x2192   # right arrow

foreach ($sh in $pres.Slides.Item(2).Shapes) {
    try {
        if ($sh.HasTextFrame) {
            $nm = $sh.Name
            if ($nm -eq "MeetingNavRange3") {
                $sh.TextFrame.TextRange.Text = ("p.7" + $EN + "14")
                Write-Output "  MeetingNavRange3 -> p.7-14"
            }
            elseif ($nm -eq "MeetingNavRange4") {
                $sh.TextFrame.TextRange.Text = ("p.15" + $EN + "27")
                Write-Output "  MeetingNavRange4 -> p.15-27"
            }
            elseif ($nm -eq "MeetingNavRange5") {
                $sh.TextFrame.TextRange.Text = "p.28"
                Write-Output "  MeetingNavRange5 -> p.28"
            }
            elseif ($nm -eq "MeetingNavRange6") {
                $sh.TextFrame.TextRange.Text = ("p.29" + $EN + "31")
                Write-Output "  MeetingNavRange6 -> p.29-31"
            }
            elseif ($nm -eq "MeetingNavRange7") {
                $sh.TextFrame.TextRange.Text = "p.32"
                Write-Output "  MeetingNavRange7 -> p.32"
            }
            elseif ($nm -eq "GlossaryPointer") {
                $sh.TextFrame.TextRange.Text = ("Abbreviations " + $AR + " Glossary appendix, p.34" + $EN + "36")
                Write-Output "  GlossaryPointer -> p.34-36"
            }
        }
    } catch {}
}

# ===========================
# SAVE
# ===========================
$pres.Save()
Write-Output "SAVED"
if (Test-Path $file) { Write-Output ("  size: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# ===========================
# RENDER QA — key slides
# ===========================
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
if (-not (Test-Path $rDir)) { New-Item -ItemType Directory $rDir | Out-Null }
foreach ($i in 2, 12, 13, 14, 15, 34, 35, 36) {
    try {
        $out = Join-Path $rDir ("slide" + $i + "_v5.png")
        $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
        Write-Output ("RENDERED s" + $i + " -> " + $out)
    } catch {
        Write-Output ("RENDER FAIL s" + $i + ": " + $_)
    }
}

try { $ppt.ActiveWindow.View.GotoSlide(13) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "ALL DONE"
