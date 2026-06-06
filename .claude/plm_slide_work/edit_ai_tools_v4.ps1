$ErrorActionPreference = "Stop"

# ===== SETUP =====
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output "FILE: $file"
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# Backup
$bkDir = Join-Path $proposal.FullName "Proposal\_backup_20260606_aitools"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory $bkDir | Out-Null }
$bk = Join-Path $bkDir "v4_BEFORE_AITOOLS.pptx"
Copy-Item $file $bk -Force
Write-Output "BACKUP: $bk"

# COM
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

# Close any stale copies WITHOUT saving
$toClose = @()
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $toClose += $p } }
foreach ($p in $toClose) { try { $p.Saved = $true; $p.Close(); Write-Output "closed stale copy" } catch {} }

# Open editable
$pres = $ppt.Presentations.Open($file, $false, $false, $true)
Write-Output "Opened: slides=$($pres.Slides.Count)  RO=$($pres.ReadOnly)"

# ===== HELPERS =====
function Get-Shape { param($slide, $name)
    foreach ($sh in $slide.Shapes) { if ($sh.Name -eq $name) { return $sh } }
    return $null
}
function Set-Text { param([int]$si, $name, $text)
    $sh = Get-Shape $pres.Slides.Item($si) $name
    if ($sh) { $sh.TextFrame.TextRange.Text = $text; Write-Output "  OK  s$si [$name]" }
    else      { Write-Output "  MISS s$si [$name]" }
}
function New-TB { param($slide, $name, [float]$l, [float]$t, [float]$w, [float]$h, $txt, [float]$fs, [bool]$bold=$false)
    $sh = $slide.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
    $sh.Name = $name
    $sh.TextFrame.WordWrap = $true
    $sh.TextFrame.AutoSize = 0  # ppAutoSizeNone
    $sh.TextFrame.TextRange.Text = $txt
    $sh.TextFrame.TextRange.Font.Size = [single]$fs
    if ($bold) { $sh.TextFrame.TextRange.Font.Bold = [int]1 }
    return $sh
}

# ===== SLIDE 6: AVEVA public signals — add both AI tools =====
Write-Output "--- s6 AVEVA public signals ---"
Set-Text 6 "TextBox 20" "CONNECT · Industrial AI Assistant (LLM) · Unified Engineering · Generative / Predictive Design AI · AIM · PI System lifecycle context"
Set-Text 6 "TextBox 21" "Opportunity: secure AI Q&A on ship data, generative design automation and design-to-operations feedback."

# ===== SLIDE 9: AVEVA differentiation — add AI tools to bottom item =====
Write-Output "--- s9 AVEVA differentiation ---"
Set-Text 9 "TextBox 56" "Industrial AI Assistant (CONNECT LLM) + Generative / Predictive Design AI + federated solvers + physics-guarded AI + evidence + operations learning"

# ===== SLIDE 19: AI gate note — mention Generative Design AI =====
Write-Output "--- s19 AI gate note ---"
Set-Text 19 "TextBox 185" "AI gate at every step: error check + anomaly diagnosis  ·  Generative Design AI assists piping routing + clash detection  ·  PASS → auto-promote  ·  FAIL → block & flag"

# ===== SLIDE 33: Glossary table 2 — insert 2 AI-tool rows before Teamcenter =====
Write-Output "--- s33 glossary table2 insert ---"
$s33 = $pres.Slides.Item(33)
$tbls = [System.Collections.ArrayList]@()
foreach ($sh in $s33.Shapes) { try { if ($sh.HasTable) { [void]$tbls.Add($sh) } } catch {} }
Write-Output "  tables on s33: $($tbls.Count)"
if ($tbls.Count -ge 2) {
    $tb = $tbls[1].Table
    Write-Output "  table2 rows=$($tb.Rows.Count) cols=$($tb.Columns.Count)"
    $tcRow = 0
    for ($r=1; $r -le $tb.Rows.Count; $r++) {
        try {
            $c1 = ($tb.Cell($r,1).Shape.TextFrame.TextRange.Text -replace "[\r\n\x0b]","").Trim()
            if ($c1 -like "*Teamcenter*") { $tcRow=$r; break }
        } catch {}
    }
    Write-Output "  Teamcenter at row=$tcRow"
    if ($tcRow -gt 0) {
        # Insert AI Assistant row before Teamcenter
        [void]($tb.Rows.Add($tcRow))
        $tb.Cell($tcRow,1).Shape.TextFrame.TextRange.Text = "AVEVA Industrial AI Asst."
        $tb.Cell($tcRow,2).Shape.TextFrame.TextRange.Text = "Industrial AI Assistant (AVEVA CONNECT)"
        $tb.Cell($tcRow,3).Shape.TextFrame.TextRange.Text = "Secure industrial LLM on CONNECT; drawing Q&A, asset specs, change history, onboarding; no data leaves CONNECT; role-gated access."
        $tb.Cell($tcRow,1).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow,2).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow,3).Shape.TextFrame.TextRange.Font.Size = [single]8
        Write-Output "  added row$tcRow: Industrial AI Assistant"

        # Insert Generative Design AI row before shifted Teamcenter
        [void]($tb.Rows.Add($tcRow+1))
        $tb.Cell($tcRow+1,1).Shape.TextFrame.TextRange.Text = "AVEVA Generative Design AI"
        $tb.Cell($tcRow+1,2).Shape.TextFrame.TextRange.Text = "Generative / Predictive Design AI (Unified Engineering)"
        $tb.Cell($tcRow+1,3).Shape.TextFrame.TextRange.Text = "AI-suggested piping routing, clash resolution, design simulation; requires human validation against class / marine standards."
        $tb.Cell($tcRow+1,1).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow+1,2).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow+1,3).Shape.TextFrame.TextRange.Font.Size = [single]8
        Write-Output "  added row$($tcRow+1): Generative Design AI"
    } else { Write-Output "  WARN: Teamcenter not found — skipping table insert" }
} else { Write-Output "  ERR: tables=$($tbls.Count) need >= 2" }

# ===== NEW SLIDE 35: AVEVA AI Capabilities — full appendix slide =====
Write-Output "--- adding slide 35: AVEVA AI Capabilities ---"
$newIdx = $pres.Slides.Count + 1   # currently 34 slides → new = 35
$ns = $pres.Slides.Add($newIdx, 12)  # 12 = ppLayoutBlank
Write-Output "  slide $newIdx created"

# Title
New-TB $ns "AITitle" 20 14 920 40 "AVEVA AI Capabilities — Industrial AI Assistant + Generative Design AI" 18 $true | Out-Null

# Subtitle
New-TB $ns "AISubtitle" 20 58 920 22 "AVEVA CONNECT LLM for secure ship data Q&A   ·   Unified Engineering generative routing and clash AI   ·   both require human review before class release" 9 | Out-Null

# Section header labels
New-TB $ns "AILabel1" 20  84 450 16 "1.  AVEVA INDUSTRIAL AI ASSISTANT  (AVEVA CONNECT — cloud LLM)" 9 $true | Out-Null
New-TB $ns "AILabel2" 490 84 450 16 "2.  AVEVA GENERATIVE / PREDICTIVE DESIGN AI  (Unified Engineering)" 9 $true | Out-Null

# Left content box: Industrial AI Assistant
$NL = [char]13
$lt = ("WHAT IT IS:") + $NL +
      ("Industry-specific LLM embedded in AVEVA CONNECT cloud platform.") + $NL + $NL +
      ("USE CASES:") + $NL +
      ("  • Drawing review and Q&A") + $NL +
      ("  • Vessel asset specification lookup") + $NL +
      ("  • Design change history tracking") + $NL +
      ("  • Worker onboarding assistance") + $NL + $NL +
      ("ADVANTAGE:") + $NL +
      ("Secure environment — ship / plant data stays inside CONNECT; no external exposure.") + $NL +
      ("Accurate search across confidential engineering and asset data.") + $NL + $NL +
      ("CONSTRAINT:") + $NL +
      ("Initial data connection and indexing required.") + $NL +
      ("Accessible data scope is limited by user role and permission settings.")
New-TB $ns "AILeft" 20 104 450 385 $lt 9 | Out-Null

# Right content box: Generative / Predictive Design AI
$rt = ("WHAT IT IS:") + $NL +
      ("Design automation and predictive AI integrated in AVEVA Unified Engineering.") + $NL + $NL +
      ("USE CASES:") + $NL +
      ("  • Piping layout optimisation") + $NL +
      ("  • 3D model clash detection and resolution") + $NL +
      ("  • AI-suggested optimal routing paths") + $NL +
      ("  • Design simulation acceleration") + $NL + $NL +
      ("ADVANTAGE:") + $NL +
      ("Automatically finds optimised pipe routing in complex ship interiors —") + $NL +
      ("reduces lead time and minimises manual rework.") + $NL + $NL +
      ("CONSTRAINT:") + $NL +
      ("Auto-generated outputs must be manually validated by a marine engineer") + $NL +
      ("against classification / shipbuilding standards before production release.")
New-TB $ns "AIRight" 490 104 450 385 $rt 9 | Out-Null

# Footer elements
New-TB $ns "AINote"     20 491 920 14 "AVEVA product capability basis: official AVEVA product pages and public CONNECT / Unified Engineering documentation. Specific model availability requires vendor confirmation." 7 | Out-Null
New-TB $ns "AIBranding" 20 507 480 20 "Future Industrial PLM" 8 | Out-Null
New-TB $ns "AIAppendix" 400 507 160 20 "Appendix" 8 | Out-Null
New-TB $ns "AIPgNum"    910 507  50 20 "35" 8 | Out-Null
Write-Output "  shapes on slide35: $($ns.Shapes.Count)"

# ===== SAVE =====
$pres.Save()
Write-Output "SAVED"
if (Test-Path $file) { Write-Output ("  new size: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# ===== RENDER QA =====
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
if (-not (Test-Path $rDir)) { New-Item -ItemType Directory $rDir | Out-Null }
foreach ($i in 6,9,19,33,35) {
    $out = Join-Path $rDir "slide$i.png"
    $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
    Write-Output "RENDERED: slide$i → $out"
}

try { $ppt.ActiveWindow.View.GotoSlide(35) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "ALL DONE — v4 AI CAPABILITIES EDIT COMPLETE"
