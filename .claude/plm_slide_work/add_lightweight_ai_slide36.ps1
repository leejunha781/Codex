$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output ("FILE: " + $file)
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# Backup
$bkDir = Join-Path $proposal.FullName "Proposal\_backup_20260606_lightweight_ai"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory $bkDir | Out-Null }
$bk = Join-Path $bkDir "v4_BEFORE_LIGHTWEIGHT_AI.pptx"
Copy-Item $file $bk -Force
Write-Output ("BACKUP: " + $bk)

# COM attach
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres=$p; break } }
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $false, $false, $true) }
Write-Output ("Opened: slides=" + $pres.Slides.Count + "  RO=" + $pres.ReadOnly)

# Colors
$WHITE = [int]16777215
$TEAL  = [int]3394713    # ~#33CC99 AVEVA teal
$AMBER = [int]3846888    # orange/amber accent
$LGREY = [int]13421772   # #CCCCCC light grey subtitle

function Add-TB($slide, $name, $l, $t, $w, $h, $txt, $fs, $bold) {
    $sh = $slide.Shapes.AddTextbox(1, [single]$l, [single]$t, [single]$w, [single]$h)
    $sh.Name = $name
    $sh.TextFrame.WordWrap = $true
    $sh.TextFrame.AutoSize = 0
    $sh.TextFrame.TextRange.Text = $txt
    $sh.TextFrame.TextRange.Font.Size = [single]$fs
    $sh.TextFrame.TextRange.Font.Color.RGB = $WHITE
    if ($bold) { $sh.TextFrame.TextRange.Font.Bold = [int]1 }
    return $sh
}

function Set-ParaColor($slide, $shName, $headings, $color) {
    $sh = $null
    foreach ($s in $slide.Shapes) { if ($s.Name -eq $shName) { $sh=$s; break } }
    if (-not $sh) { return }
    $n = $sh.TextFrame.TextRange.Paragraphs().Count
    for ($pi=1; $pi -le $n; $pi++) {
        try {
            $pt = $sh.TextFrame.TextRange.Paragraphs($pi).Text.Trim()
            foreach ($hd in $headings) {
                if ($pt -like ($hd+"*") -or $pt -eq $hd) {
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Bold = [int]1
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Color.RGB = $color
                    break
                }
            }
        } catch {}
    }
}

# ===== ADD SLIDE 36: Lightweight AI Alternative =====
Write-Output "--- Adding slide 36: Lightweight AI Alternative ---"
$newIdx = $pres.Slides.Count + 1   # 36
$ns = $pres.Slides.Add($newIdx, 12)
Write-Output ("  slide " + $newIdx + " created")

$NL = [char]13

# Title + Subtitle
Add-TB $ns "LWTitle"    20 14 920 40 "Lightweight AI Alternative - Rule-Based Gate  to  On-Prem RAG  to  AVEVA CONNECT" 16 $true | Out-Null
Add-TB $ns "LWSubtitle" 20 58 920 22 "pgvector + Ollama on the existing Linux control-plane  |  zero added license  |  full data sovereignty  |  upgrade to AVEVA CONNECT AI when org is ready" 9 $false | Out-Null

# Section headers
Add-TB $ns "LWLabel1" 20  84 460 16 "LIGHTWEIGHT STACK  (Tier 0 + Tier 1 — on existing infrastructure)" 9 $true | Out-Null
Add-TB $ns "LWLabel2" 500 84 420 16 "WHY LIGHTWEIGHT FIRST  +  PHASED UPGRADE PATH" 9 $true | Out-Null

# ===== LEFT COLUMN: Lightweight Stack =====
$NL = [char]13

$left  = "TIER 0 — RULE-BASED AI GATE  (zero license, immediate)" + $NL
$left += "FastAPI + Pydantic validation rules:" + $NL
$left += "  - Drawing standard: title block, layer names, revision format" + $NL
$left += "  - PLM metadata: naming convention, effectivity scope, required fields" + $NL
$left += "  - VCRM / checklist: coverage completeness check before promote" + $NL
$left += "Deterministic, <10ms response, no GPU, no model dependency." + $NL
$left += "Already in every gate step — just formalise the rule set." + $NL
$left += $NL
$left += "TIER 1 — ON-PREM RAG  (pgvector + Ollama, existing Linux nodes)" + $NL
$left += "pgvector  (PostgreSQL extension, one command: CREATE EXTENSION vector)" + $NL
$left += "  - Embed ship drawings, specs, asset data, change history as vectors" + $NL
$left += "  - Semantic search: 'hull compartments with CoG deviation > 5%'" + $NL
$left += "  - Similarity search over evidence packs without full-text scan" + $NL
$left += $NL
$left += "Ollama  (local LLM server, OpenAI-compatible REST API)" + $NL
$left += "  - Models: Llama 3.2 (3B / 8B)  Phi-3 Mini  Mistral 7B" + $NL
$left += "  - CPU-only viable for 7B; single mid-range GPU handles 13B" + $NL
$left += "  - No data leaves the shipyard network - IP & GDPR compliant" + $NL
$left += "  - FastAPI /ai/query  /ai/embed endpoints: drop-in replacement" + $NL
$left += $NL
$left += "LIGHTWEIGHT PIPING OPTIMISATION  (no Generative Design AI license)" + $NL
$left += "  - Python scipy + networkx constraint solver" + $NL
$left += "  - Naval engineering rules encoded as hard constraints" + $NL
$left += "  - 3D voxel clash grid + shortest-path routing; no ML model needed" + $NL
$left += "  - Validate results against class rules before promote"

Add-TB $ns "LWLeft" 20 104 460 385 $left 8.5 $false | Out-Null

# ===== RIGHT COLUMN: Why + Upgrade Path =====
$right  = "WHY LIGHTWEIGHT FIRST" + $NL
$right += "AVEVA CONNECT AI requires:" + $NL
$right += "  - AVEVA cloud subscription + network dependency" + $NL
$right += "  - Data ingestion pipeline + indexing lead time (weeks)" + $NL
$right += "  - GDPR / IP review before ship data leaves the yard" + $NL
$right += "Generative Design AI requires:" + $NL
$right += "  - AVEVA Unified Engineering full license" + $NL
$right += "  - GPU workstations + model tuning for ship geometry" + $NL
$right += "Tier 0-1 runs on the existing 4-node Linux HA cluster + PostgreSQL." + $NL
$right += $NL
$right += "PHASED UPGRADE PATH" + $NL
$right += $NL
$right += "Tier 0  |  Today  |  No license  |  No GPU" + $NL
$right += "Rule-based AI gate" + $NL
$right += "FastAPI + Pydantic rules — deterministic, instant" + $NL
$right += $NL
$right += "Tier 1  |  Pilot (4-8 wks)  |  OSS free  |  CPU or 1 GPU" + $NL
$right += "On-prem RAG" + $NL
$right += "pgvector + Ollama 7B-13B — ship data Q&A, in-yard" + $NL
$right += $NL
$right += "Tier 2  |  Scale  |  AVEVA subscription  |  Cloud GPU" + $NL
$right += "AVEVA CONNECT AI" + $NL
$right += "Full Industrial LLM on CONNECT — enterprise Q&A + insights" + $NL
$right += $NL
$right += "Tier 3  |  Full design AI  |  Unified Engineering  |  Workstation GPU" + $NL
$right += "AVEVA Generative Design AI" + $NL
$right += "Full automated piping optimisation + clash AI"

Add-TB $ns "LWRight" 500 104 420 385 $right 8.5 $false | Out-Null

# Footer
Add-TB $ns "LWNote"     20 491 920 14 "Tier 0-1 OSS stack: FastAPI (MIT)  |  PostgreSQL (PostgreSQL License)  |  pgvector (MIT)  |  Ollama (MIT)  — all run on the existing Linux HA cluster and PostgreSQL already defined in the reference architecture." 7 $false | Out-Null
Add-TB $ns "LWBranding" 20 507 480 20 "Future Industrial PLM" 8 $false | Out-Null
Add-TB $ns "LWAppendix" 400 507 160 20 "Appendix" 8 $false | Out-Null
Add-TB $ns "LWPgNum"    910 507  50 20 "36" 8 $false | Out-Null
Write-Output ("  shapes on slide36: " + $ns.Shapes.Count)

# Color: section heading paragraphs teal, upgrade tier labels amber
$lwHeadings = @("TIER 0", "TIER 1", "LIGHTWEIGHT PIPING", "WHY LIGHTWEIGHT FIRST", "PHASED UPGRADE PATH", "pgvector", "Ollama")
Set-ParaColor $ns "LWLeft"  $lwHeadings $TEAL
Set-ParaColor $ns "LWRight" @("WHY LIGHTWEIGHT FIRST","PHASED UPGRADE PATH","Tier 0","Tier 1","Tier 2","Tier 3") $TEAL

# Tier labels in right column: make Tier 0/1 teal, Tier 2/3 amber
foreach ($shName in @("LWRight")) {
    $sh = $null
    foreach ($s in $ns.Shapes) { if ($s.Name -eq $shName) { $sh=$s; break } }
    if ($sh) {
        $n = $sh.TextFrame.TextRange.Paragraphs().Count
        for ($pi=1; $pi -le $n; $pi++) {
            try {
                $pt = $sh.TextFrame.TextRange.Paragraphs($pi).Text.Trim()
                if ($pt -like "Tier 2*" -or $pt -like "Tier 3*") {
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Bold = [int]1
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Color.RGB = $AMBER
                }
            } catch {}
        }
    }
}

# Column labels: teal
foreach ($shName in @("LWLabel1","LWLabel2")) {
    $sh = $null
    foreach ($s in $ns.Shapes) { if ($s.Name -eq $shName) { $sh=$s; break } }
    if ($sh) { $sh.TextFrame.TextRange.Font.Color.RGB = $TEAL }
}
# Subtitle: light grey
foreach ($s in $ns.Shapes) {
    if ($s.Name -eq "LWSubtitle") { $s.TextFrame.TextRange.Font.Color.RGB = $LGREY; break }
}

# ===== ALSO: update slide 9 TextBox 52 to hint at lightweight AI option =====
Write-Output "--- s9: add lightweight AI hint ---"
$s9 = $pres.Slides.Item(9)
$sh52 = $null
foreach ($s in $s9.Shapes) { if ($s.Name -eq "TextBox 52") { $sh52=$s; break } }
if ($sh52) {
    $sh52.TextFrame.TextRange.Text = "E3D / Marine + Unified Engineering + AIM + PI + CONNECT (AI Asst.) + Lightweight RAG (pgvector + Ollama)"
    Write-Output "  OK s9 [TextBox 52]"
} else { Write-Output "  MISS s9 [TextBox 52]" }

# ===== SAVE =====
$pres.Save()
Write-Output "SAVED"
if (Test-Path $file) { Write-Output ("  new size: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# ===== RENDER QA =====
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
if (-not (Test-Path $rDir)) { New-Item -ItemType Directory $rDir | Out-Null }
foreach ($i in 9,35,36) {
    $out = Join-Path $rDir ("slide" + $i + "_lw.png")
    $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
    Write-Output ("RENDERED: slide" + $i + " -> " + $out)
}

try { $ppt.ActiveWindow.View.GotoSlide(36) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "ALL DONE — slide36 LIGHTWEIGHT AI ADDED"
