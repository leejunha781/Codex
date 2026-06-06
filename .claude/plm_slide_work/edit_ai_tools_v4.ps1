$ErrorActionPreference = "Stop"

# ===== SETUP =====
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output "FILE: $file"
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

$bkDir = Join-Path $proposal.FullName "Proposal\_backup_20260606_aitools"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory $bkDir | Out-Null }
$bk = Join-Path $bkDir "v4_BEFORE_AITOOLS.pptx"
Copy-Item $file $bk -Force
Write-Output "BACKUP: $bk"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

$toClose = @()
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $toClose += $p } }
foreach ($p in $toClose) { try { $p.Saved = $true; $p.Close(); Write-Output "closed stale" } catch {} }

$pres = $ppt.Presentations.Open($file, $false, $false, $true)
Write-Output ("Opened: slides=" + $pres.Slides.Count + "  RO=" + $pres.ReadOnly)

function Get-Shape($slide, $name) {
    foreach ($sh in $slide.Shapes) { if ($sh.Name -eq $name) { return $sh } }
    return $null
}
function Set-Text($si, $name, $text) {
    $sh = Get-Shape $pres.Slides.Item($si) $name
    if ($sh) { $sh.TextFrame.TextRange.Text = $text; Write-Output ("  OK  s" + $si + " [" + $name + "]") }
    else      { Write-Output ("  MISS s" + $si + " [" + $name + "]") }
}

# ===== SLIDE 6 =====
Write-Output "--- s6 AVEVA public signals ---"
Set-Text 6 "TextBox 20" "CONNECT - Industrial AI Assistant (LLM) - Unified Engineering - Generative / Predictive Design AI - AIM - PI System lifecycle context"
Set-Text 6 "TextBox 21" "Opportunity: secure AI Q&A on ship data, generative design automation and design-to-operations feedback."

# ===== SLIDE 9 =====
Write-Output "--- s9 AVEVA differentiation ---"
Set-Text 9 "TextBox 56" "Industrial AI Assistant (CONNECT LLM) + Generative / Predictive Design AI + federated solvers + physics-guarded AI + evidence + operations learning"

# ===== SLIDE 19 =====
Write-Output "--- s19 AI gate note ---"
Set-Text 19 "TextBox 185" "AI gate at every step: error check + anomaly diagnosis  -  Generative Design AI assists piping routing + clash detection  -  PASS = auto-promote  /  FAIL = block & flag"

# ===== SLIDE 33: Glossary table 2 — insert AI-tool rows =====
Write-Output "--- s33 glossary table ---"
$s33 = $pres.Slides.Item(33)
$tbls = New-Object System.Collections.ArrayList
foreach ($sh in $s33.Shapes) { try { if ($sh.HasTable) { [void]$tbls.Add($sh) } } catch {} }
Write-Output ("  tables: " + $tbls.Count)
if ($tbls.Count -ge 2) {
    $tb = $tbls[1].Table
    Write-Output ("  table2 rows=" + $tb.Rows.Count + " cols=" + $tb.Columns.Count)
    $tcRow = 0
    for ($r=1; $r -le $tb.Rows.Count; $r++) {
        try {
            $c1 = $tb.Cell($r,1).Shape.TextFrame.TextRange.Text
            $c1 = $c1 -replace "[\r\n\x0b]",""
            $c1 = $c1.Trim()
            if ($c1 -like "*Teamcenter*") { $tcRow=$r; break }
        } catch {}
    }
    Write-Output ("  Teamcenter at row=" + $tcRow)
    if ($tcRow -gt 0) {
        [void]($tb.Rows.Add($tcRow))
        $tb.Cell($tcRow,1).Shape.TextFrame.TextRange.Text = "AVEVA Industrial AI Asst."
        $tb.Cell($tcRow,2).Shape.TextFrame.TextRange.Text = "Industrial AI Assistant (AVEVA CONNECT)"
        $tb.Cell($tcRow,3).Shape.TextFrame.TextRange.Text = "Secure industrial LLM on CONNECT; drawing Q&A, asset specs, change history, onboarding; no data leaves CONNECT; role-gated access."
        $tb.Cell($tcRow,1).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow,2).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow,3).Shape.TextFrame.TextRange.Font.Size = [single]8
        Write-Output ("  added row" + $tcRow + ": Industrial AI Asst.")

        [void]($tb.Rows.Add($tcRow+1))
        $tb.Cell($tcRow+1,1).Shape.TextFrame.TextRange.Text = "AVEVA Generative Design AI"
        $tb.Cell($tcRow+1,2).Shape.TextFrame.TextRange.Text = "Generative / Predictive Design AI (Unified Engineering)"
        $tb.Cell($tcRow+1,3).Shape.TextFrame.TextRange.Text = "AI-suggested piping routing, clash resolution, design simulation; requires human validation against class / marine standards."
        $tb.Cell($tcRow+1,1).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow+1,2).Shape.TextFrame.TextRange.Font.Size = [single]8
        $tb.Cell($tcRow+1,3).Shape.TextFrame.TextRange.Font.Size = [single]8
        Write-Output ("  added row" + ($tcRow+1) + ": Generative Design AI")
    } else { Write-Output "  WARN: Teamcenter not found" }
} else { Write-Output ("  ERR: tables=" + $tbls.Count) }

# ===== NEW SLIDE 35: AVEVA AI Capabilities =====
Write-Output "--- adding slide 35 ---"
$newIdx = $pres.Slides.Count + 1
$ns = $pres.Slides.Add($newIdx, 12)
Write-Output ("  slide " + $newIdx + " created")

function Add-TB($slide, $name, $l, $t, $w, $h, $txt, $fs, $bold) {
    $sh = $slide.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
    $sh.Name = $name
    $sh.TextFrame.WordWrap = $true
    $sh.TextFrame.AutoSize = 0
    $sh.TextFrame.TextRange.Text = $txt
    $sh.TextFrame.TextRange.Font.Size = [single]$fs
    if ($bold) { $sh.TextFrame.TextRange.Font.Bold = [int]1 }
    return $sh
}

$NL = [char]13

Add-TB $ns "AITitle" 20 14 920 40 "AVEVA AI Capabilities - Industrial AI Assistant + Generative Design AI" 18 $true | Out-Null
Add-TB $ns "AISubtitle" 20 58 920 22 "AVEVA CONNECT LLM for secure ship data Q&A   |   Unified Engineering generative routing and clash AI   |   both require human validation before class release" 9 $false | Out-Null
Add-TB $ns "AILabel1" 20  84 450 16 "1.  AVEVA INDUSTRIAL AI ASSISTANT  (AVEVA CONNECT - cloud LLM)" 9 $true | Out-Null
Add-TB $ns "AILabel2" 490 84 450 16 "2.  AVEVA GENERATIVE / PREDICTIVE DESIGN AI  (Unified Engineering)" 9 $true | Out-Null

$lt  = "WHAT IT IS:" + $NL
$lt += "Industry-specific LLM embedded in AVEVA CONNECT cloud platform." + $NL + $NL
$lt += "USE CASES:" + $NL
$lt += "  - Drawing review and Q&A" + $NL
$lt += "  - Vessel asset specification lookup" + $NL
$lt += "  - Design change history tracking" + $NL
$lt += "  - Worker onboarding assistance" + $NL + $NL
$lt += "ADVANTAGE:" + $NL
$lt += "Secure environment - ship / plant data stays inside CONNECT; no external exposure." + $NL
$lt += "Accurate search across confidential engineering and asset data." + $NL + $NL
$lt += "CONSTRAINT:" + $NL
$lt += "Initial data connection and indexing required." + $NL
$lt += "Accessible data scope is limited by user role and permission settings."
Add-TB $ns "AILeft" 20 104 450 385 $lt 9 $false | Out-Null

$rt  = "WHAT IT IS:" + $NL
$rt += "Design automation and predictive AI integrated in AVEVA Unified Engineering." + $NL + $NL
$rt += "USE CASES:" + $NL
$rt += "  - Piping layout optimisation" + $NL
$rt += "  - 3D model clash detection and resolution" + $NL
$rt += "  - AI-suggested optimal routing paths" + $NL
$rt += "  - Design simulation acceleration" + $NL + $NL
$rt += "ADVANTAGE:" + $NL
$rt += "Automatically finds optimised pipe routing in complex ship interiors -" + $NL
$rt += "reduces lead time and minimises manual rework." + $NL + $NL
$rt += "CONSTRAINT:" + $NL
$rt += "Auto-generated outputs must be manually validated by a marine engineer" + $NL
$rt += "against classification / shipbuilding standards before production release."
Add-TB $ns "AIRight" 490 104 450 385 $rt 9 $false | Out-Null

Add-TB $ns "AINote"     20 491 920 14 "AVEVA product capability basis: official AVEVA product pages and public CONNECT / Unified Engineering documentation. Specific model availability requires vendor confirmation." 7 $false | Out-Null
Add-TB $ns "AIBranding" 20 507 480 20 "Future Industrial PLM" 8 $false | Out-Null
Add-TB $ns "AIAppendix" 400 507 160 20 "Appendix" 8 $false | Out-Null
Add-TB $ns "AIPgNum"    910 507  50 20 "35" 8 $false | Out-Null
Write-Output ("  shapes on slide35: " + $ns.Shapes.Count)

# ===== SAVE =====
$pres.Save()
Write-Output "SAVED"
if (Test-Path $file) { Write-Output ("  new size: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# ===== RENDER QA =====
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
if (-not (Test-Path $rDir)) { New-Item -ItemType Directory $rDir | Out-Null }
foreach ($i in 6,9,19,33,35) {
    $out = Join-Path $rDir ("slide" + $i + ".png")
    $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
    Write-Output ("RENDERED: slide" + $i + " -> " + $out)
}

try { $ppt.ActiveWindow.View.GotoSlide(35) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "ALL DONE"
