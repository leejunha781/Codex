$ErrorActionPreference = "Stop"

$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output ("FILE: " + $file)
if (Test-Path $file) { Write-Output ("SIZE: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# Backup
$bkDir = Join-Path $proposal.FullName "Proposal\_backup_20260606_piping_routing"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory $bkDir | Out-Null }
$bk = Join-Path $bkDir "v4_BEFORE_PIPING_ROUTING.pptx"
Copy-Item $file $bk -Force
Write-Output ("BACKUP: " + $bk)

# COM attach
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

$toClose = @()
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $toClose += $p } }
foreach ($p in $toClose) { try { $p.Saved = $true; $p.Close() } catch {} }

$pres = $ppt.Presentations.Open($file, $false, $false, $true)
Write-Output ("Opened: slides=" + $pres.Slides.Count + "  RO=" + $pres.ReadOnly)

# Color palette
$WHITE      = [int]16777215
$TEAL       = [int]3394713    # #33CC99
$AMBER      = [int]3846888    # orange-amber
$LGREY      = [int]13421772   # #CCCCCC light grey
$CARDCOLOR  = [int]1650015    # RGB(25,45,95) dark navy card
$FREECOLOR  = [int]1977685    # RGB(30,45,85) voxel free cell
$BLOCKEDCOL = [int]5591120    # RGB(85,80,80) voxel blocked cell
$STARTGREEN = [int]3983460    # RGB(60,200,100) start cell
# AMBER already defined for end cell

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

function Set-HeadingsTeal($slide, $shName, $headings) {
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
                    $sh.TextFrame.TextRange.Paragraphs($pi).Font.Color.RGB = $TEAL
                    break
                }
            }
        } catch {}
    }
}

# ===================================================
# STEP 1: ADD NEW SLIDE 15 — Lightweight Piping Routing
# ===================================================
Write-Output "=== STEP 1: Creating slide 15 ==="
$ns = $pres.Slides.Add(15, 12)  # ppLayoutBlank
Write-Output ("  blank slide created at pos 15  (total: " + $pres.Slides.Count + ")")

$NL = [char]13

# ----- Title + Subtitle -----
Add-TB $ns "LPTitle"    20 14 920 40 "Lightweight Piping Routing  —  scipy + networkx + 3D Voxel Grid" 16 $true  | Out-Null
Add-TB $ns "LPSubtitle" 20 58 920 22 "Constraint-based route proposals without Generative Design AI license  |  deterministic solver  |  human + class validation gate before E3D promote" 9 $false | Out-Null

# Subtitle grey
foreach ($s in $ns.Shapes) {
    if ($s.Name -eq "LPSubtitle") { $s.TextFrame.TextRange.Font.Color.RGB = $LGREY }
}

# ----- Column card backgrounds -----
$c = $ns.Shapes.AddShape(5, [single]14, [single]78, [single]462, [single]413)
$c.Name = "LPCardLeft"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

$c = $ns.Shapes.AddShape(5, [single]482, [single]78, [single]464, [single]413)
$c.Name = "LPCardRight"
$c.Fill.Solid()
$c.Fill.ForeColor.RGB = $CARDCOLOR
$c.Fill.Transparency = [single]0.25
$c.Line.ForeColor.RGB = $TEAL
$c.Line.Weight = [single]1.0
[void]$c.ZOrder(1)

# Accent bar + column divider
$b = $ns.Shapes.AddShape(1, [single]14, [single]74, [single]932, [single]2)
$b.Name = "LPAccentBar"
$b.Fill.Solid()
$b.Fill.ForeColor.RGB = $TEAL
$b.Line.Visible = [int]0
[void]$b.ZOrder(1)

$dv = $ns.Shapes.AddLine([single]478, [single]80, [single]478, [single]490)
$dv.Name = "LPDivider"
$dv.Line.ForeColor.RGB = $TEAL
$dv.Line.Weight = [single]0.75

# ----- Column labels -----
$shL = Add-TB $ns "LPLabel1" 20  84 446 16 "ALGORITHM PIPELINE  (STEP 1 → 4)  +  ENGINEERING CONSTRAINTS" 9 $true
$shL.TextFrame.TextRange.Font.Color.RGB = $TEAL
$shR = Add-TB $ns "LPLabel2" 488 84 450 16 "3D VOXEL ROUTING CONCEPT  (2D cross-section view)" 9 $true
$shR.TextFrame.TextRange.Font.Color.RGB = $TEAL

# ----- LEFT COLUMN: Algorithm pipeline text -----
$lt  = "STEP 1 — INPUT  (E3D/Marine via Local Agent)" + $NL
$lt += "  Compartment boundary, bulkhead positions, keep-out zones" + $NL
$lt += "  Equipment obstacles, clearance volumes, pipe endpoints" + $NL
$lt += "  Export: Local Agent JSON → POST /routing/submit" + $NL
$lt += $NL
$lt += "STEP 2 — BUILD VOXEL GRID" + $NL
$lt += "  Divide ship space into uniform 3D cells (voxels)" + $NL
$lt += "  Mark BLOCKED: bulkheads, equipment, hull keep-outs" + $NL
$lt += "  Assign cost weights: length, bend penalty, zone priority" + $NL
$lt += $NL
$lt += "STEP 3 — SOLVE  (scipy + networkx)" + $NL
$lt += "  Build directed graph: nodes = free voxels, edges = valid moves" + $NL
$lt += "  Hard constraints: max bend radius (class), min support spacing," + $NL
$lt += "    hull/bulkhead clearance margin, zone routing rules" + $NL
$lt += "  Dijkstra / A* → minimum-cost feasible path through free voxels" + $NL
$lt += $NL
$lt += "STEP 4 — OUTPUT + VALIDATE" + $NL
$lt += "  Proposed route JSON → POST /routing/propose" + $NL
$lt += "  Engineer reviews route in E3D spatial context" + $NL
$lt += "  FastAPI class-rule check gate → approve → promote to E3D" + $NL
$lt += $NL
$lt += "WHY THIS APPROACH" + $NL
$lt += "Deterministic — fully explainable; no ML black box; auditable trace." + $NL
$lt += "OSS stack — scipy (BSD) + networkx (BSD); no added license cost." + $NL
$lt += "Runs on existing Linux control-plane; no GPU required." + $NL
$lt += "Human + class gate — no automated E3D promote without approval."

Add-TB $ns "LPLeft" 20 104 454 381 $lt 8.5 $false | Out-Null
Set-HeadingsTeal $ns "LPLeft" @("STEP 1","STEP 2","STEP 3","STEP 4","WHY THIS APPROACH")

# ----- RIGHT COLUMN: Voxel grid cells -----
# Grid: 6 cols x 4 rows, cell 44x32, stride 46x34, origin x=490 y=108
# Cell types: 0=free 1=blocked 2=path 3=start(green) 4=end(amber)
# Path: S(0,0)->(0,1)->(0,2)->(0,3)->(0,4)->(0,5)->(1,5)->(2,5)->(3,5)=E
# Blocked: (1,2),(2,2) = wall section; (1,4),(2,4) = equipment block
$gridData = @(
    @(3,2,2,2,2,2),   # row 0: start + path along top row
    @(0,0,1,0,1,2),   # row 1: blocked at c2(wall) c4(equip); path at c5
    @(0,0,1,0,1,2),   # row 2: blocked at c2(wall) c4(equip); path at c5
    @(0,0,0,0,0,4)    # row 3: end at c5
)

$DARKLINE = [int](5*65536+5*256+5)  # near-black cell border

Write-Output "  Adding voxel grid cells..."
for ($r=0; $r -lt 4; $r++) {
    for ($c=0; $c -lt 6; $c++) {
        $cx = 490 + $c * 46
        $cy = 108 + $r * 34
        $ct = $gridData[$r][$c]
        $cell = $ns.Shapes.AddShape(1, [single]$cx, [single]$cy, [single]44, [single]32)
        $cell.Name = ("VoxCell_" + $r + "_" + $c)
        $cell.Fill.Solid()
        if     ($ct -eq 0) { $cell.Fill.ForeColor.RGB = $FREECOLOR   }
        elseif ($ct -eq 1) { $cell.Fill.ForeColor.RGB = $BLOCKEDCOL  }
        elseif ($ct -eq 2) { $cell.Fill.ForeColor.RGB = $TEAL        }
        elseif ($ct -eq 3) { $cell.Fill.ForeColor.RGB = $STARTGREEN  }
        elseif ($ct -eq 4) { $cell.Fill.ForeColor.RGB = $AMBER       }
        $cell.Line.ForeColor.RGB = $DARKLINE
        $cell.Line.Weight = [single]0.5
        # S / E label on start and end cells
        if ($ct -eq 3 -or $ct -eq 4) {
            $lbl = if ($ct -eq 3) { "S" } else { "E" }
            $cell.TextFrame.TextRange.Text = $lbl
            $cell.TextFrame.TextRange.Font.Size = [single]9
            $cell.TextFrame.TextRange.Font.Bold = [int]1
            $cell.TextFrame.TextRange.Font.Color.RGB = $WHITE
            $cell.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # ppAlignCenter
        }
        # Dim "X" label on blocked cells
        if ($ct -eq 1) {
            $cell.TextFrame.TextRange.Text = "X"
            $cell.TextFrame.TextRange.Font.Size = [single]8
            $cell.TextFrame.TextRange.Font.Color.RGB = $LGREY
            $cell.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }
    }
}
Write-Output "  Voxel grid complete (24 cells)"

# Obstacle labels (to the right of the grid, identifying what the X cells are)
$wallLabel  = Add-TB $ns "LPWallLabel"  737 124 200 16 "<-- WALL (bulkhead)" 7.5 $false
$wallLabel.TextFrame.TextRange.Font.Color.RGB = $LGREY

$equipLabel = Add-TB $ns "LPEquipLabel" 737 162 200 16 "<-- EQUIPMENT block" 7.5 $false
$equipLabel.TextFrame.TextRange.Font.Color.RGB = $LGREY

# Path label
$pathLabel = Add-TB $ns "LPPathLabel" 490 90 280 16 "PATH (Dijkstra/A* optimal route)  →" 7.5 $false
$pathLabel.TextFrame.TextRange.Font.Color.RGB = $TEAL

# ----- RIGHT COLUMN: Legend + solver info (below grid, y=248 onward) -----
$rt  = "LEGEND" + $NL
$rt += "  Teal  = Optimal pipe route  (Dijkstra / A* path)" + $NL
$rt += "  Grey  = Blocked voxel  (bulkhead / equipment / keep-out)" + $NL
$rt += "  Dark  = Free voxel  (valid routing space)" + $NL
$rt += "  S = Route start point    E = Route end point" + $NL
$rt += $NL
$rt += "COST FUNCTION" + $NL
$rt += "  minimize:  pipe_length * w_len + bend_count * w_bend" + $NL
$rt += "  subject to:" + $NL
$rt += "    bend_radius >= class_min  (e.g. 3 x DN for Cu alloy)" + $NL
$rt += "    support_spacing <= max_span  (weight + class rule)" + $NL
$rt += "    hull/bulkhead clearance >= min_margin" + $NL
$rt += $NL
$rt += "SOLVER STACK  (all existing Linux control-plane)" + $NL
$rt += "  scipy      constraint satisfaction backend  (BSD)" + $NL
$rt += "  networkx   directed voxel graph + shortest path  (BSD)" + $NL
$rt += "  FastAPI    /routing/propose endpoint  (MIT)" + $NL
$rt += "  No GPU required — runs on existing 4-node Linux HA cluster"

Add-TB $ns "LPRight" 488 248 450 237 $rt 8.5 $false | Out-Null
Set-HeadingsTeal $ns "LPRight" @("LEGEND","COST FUNCTION","SOLVER STACK")

# ----- Footer -----
Add-TB $ns "LPNote"     20 491 920 14 "Engineering constraints must be parameterised per class society (DNV, LR, NK etc.) and pipe material before use. All route proposals require human review and class-rule validation before promote to E3D/Marine." 7 $false | Out-Null
Add-TB $ns "LPBranding" 20 507 480 20 "Future Industrial PLM" 8 $false | Out-Null
Add-TB $ns "LPSection"  400 507 160 20 "AI STRATEGY" 8 $false | Out-Null
Add-TB $ns "LPPgNum"    910 507  50 20 "15" 8 $false | Out-Null

Write-Output ("  slide 15 shapes: " + $ns.Shapes.Count)

# ===================================================
# STEP 2: UPDATE PAGE NUMBERS — slides 16-37 (was 15-36, +1)
# ===================================================
Write-Output "=== STEP 2: Page numbers ==="
for ($si=16; $si -le 37; $si++) {
    $oldPg = [string]($si - 1)
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

# ===================================================
# STEP 3: UPDATE SLIDE 2 NAVIGATION
# MeetingNavRange3 (WINNING IDEAS):  p.7-14 -> p.7-15
# MeetingNavRange4 (PROPOSAL+DELIVERY): p.15-27 -> p.16-28
# MeetingNavRange5: p.28 -> p.29
# MeetingNavRange6: p.29-31 -> p.30-32
# MeetingNavRange7: p.32 -> p.33
# GlossaryPointer: p.34-36 -> p.35-37
# ===================================================
Write-Output "=== STEP 3: Slide 2 navigation ==="
$EN = [char]0x2013   # en-dash

foreach ($sh in $pres.Slides.Item(2).Shapes) {
    try {
        if ($sh.HasTextFrame) {
            $nm = $sh.Name
            if ($nm -eq "MeetingNavRange3") {
                $sh.TextFrame.TextRange.Text = ("p.7" + $EN + "15")
                Write-Output "  MeetingNavRange3 -> p.7-15"
            }
            elseif ($nm -eq "MeetingNavRange4") {
                $sh.TextFrame.TextRange.Text = ("p.16" + $EN + "28")
                Write-Output "  MeetingNavRange4 -> p.16-28"
            }
            elseif ($nm -eq "MeetingNavRange5") {
                $sh.TextFrame.TextRange.Text = "p.29"
                Write-Output "  MeetingNavRange5 -> p.29"
            }
            elseif ($nm -eq "MeetingNavRange6") {
                $sh.TextFrame.TextRange.Text = ("p.30" + $EN + "32")
                Write-Output "  MeetingNavRange6 -> p.30-32"
            }
            elseif ($nm -eq "MeetingNavRange7") {
                $sh.TextFrame.TextRange.Text = "p.33"
                Write-Output "  MeetingNavRange7 -> p.33"
            }
            elseif ($nm -eq "GlossaryPointer") {
                $AR = [char]0x2192
                $sh.TextFrame.TextRange.Text = ("Abbreviations " + $AR + " Glossary appendix, p.35" + $EN + "37")
                Write-Output "  GlossaryPointer -> p.35-37"
            }
        }
    } catch {}
}

# ===================================================
# SAVE
# ===================================================
$pres.Save()
Write-Output "SAVED"
if (Test-Path $file) { Write-Output ("  size: {0}  LastWrite: {1}" -f (Get-Item $file).Length, (Get-Item $file).LastWriteTime) }

# ===================================================
# RENDER QA
# ===================================================
$rDir = "C:\Users\namma\.claude\plm_slide_work\ai_render"
if (-not (Test-Path $rDir)) { New-Item -ItemType Directory $rDir | Out-Null }
foreach ($i in 2, 14, 15, 16) {
    try {
        $out = Join-Path $rDir ("slide" + $i + "_v6.png")
        $pres.Slides.Item($i).Export($out, "PNG", 1440, 810)
        Write-Output ("RENDERED s" + $i + " -> " + $out)
    } catch { Write-Output ("RENDER FAIL s" + $i + ": " + $_) }
}

try { $ppt.ActiveWindow.View.GotoSlide(15) } catch {}
try { $ppt.Activate() } catch {}
Write-Output "ALL DONE"
