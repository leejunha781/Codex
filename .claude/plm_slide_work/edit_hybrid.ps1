$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$dir = Join-Path $proposal.FullName "Proposal"
$file = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

# Find if already open; else open read-write
$pres = $null; $opened = $false
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $false, $false, $true); $opened = $true }
Write-Output "OPENED-FRESH=$opened  Slides=$($pres.Slides.Count)"

function Get-Shape($slide, $name) {
  foreach ($sh in $slide.Shapes) { if ($sh.Name -eq $name) { return $sh } }
  return $null
}
function SetWhole($slideIdx, $name, $text) {
  $slide = $pres.Slides.Item($slideIdx)
  $sh = Get-Shape $slide $name
  if (-not $sh) { Write-Output "  !! MISSING s$slideIdx [$name]"; return }
  $old = $sh.TextFrame.TextRange.Text
  $sh.TextFrame.TextRange.Text = $text
  Write-Output "  OK s$slideIdx [$name] set ($($text.Length) chars)"
}
function SetPara($slideIdx, $name, $pidx, $text) {
  $slide = $pres.Slides.Item($slideIdx)
  $sh = Get-Shape $slide $name
  if (-not $sh) { Write-Output "  !! MISSING s$slideIdx [$name]"; return }
  $tr = $sh.TextFrame.TextRange
  if ($tr.Paragraphs().Count -lt $pidx) { Write-Output "  !! s$slideIdx [$name] only $($tr.Paragraphs().Count) paras"; return }
  $tr.Paragraphs($pidx).Text = $text
  Write-Output "  OK s$slideIdx [$name] para$pidx set"
}
function SetGeom($slideIdx, $name, $top, $height) {
  $slide = $pres.Slides.Item($slideIdx)
  $sh = Get-Shape $slide $name
  if (-not $sh) { Write-Output "  !! MISSING s$slideIdx [$name]"; return }
  $sh.Top = [single]$top; $sh.Height = [single]$height
  Write-Output "  OK s$slideIdx [$name] geom T=$top H=$height"
}

$CR = [string][char]13
$mid = [char]0x00B7   # middle dot
$harr = [char]0x2194  # left-right arrow
$rarr = [char]0x2192  # right arrow
$dash = [char]0x2014  # em dash

Write-Output "=== SLIDE 15 (Hardware / Hybrid architecture) ==="
SetWhole 15 "TextBox 2"  ("Hybrid Architecture " + $harr + " Windows Authoring " + $harr + " Linux PLM Control Plane (API)    &    HD Hyundai API Example")
SetWhole 15 "TextBox 3"  ("Windows-based AVEVA E3D/Marine authoring + Linux-based Open PLM Control Plane (API " + $mid + " SQL " + $mid + " AI Gate " + $mid + " Evidence " + $mid + " Integration) + Cloud/CONNECT " + $dash + " with a simple step-by-step client HD Hyundai can run and verify.")
SetWhole 15 "TextBox 4"  ("A - Reference Architecture: Windows Authoring " + $harr + " Linux PLM Control Plane " + $harr + " API (hybrid)")
SetWhole 15 "TextBox 15" ("API issues commands  " + $harr + "  Control Plane returns JSON; E3D/Hull Agent executes & calls back")
SetWhole 15 "TextBox 18" ("2.  OPEN PLM CONTROL PLANE (Linux): API " + $mid + " configuration " + $mid + " AI gate " + $mid + " evidence " + $mid + " integration")
SetWhole 15 "Rounded Rectangle 23" "AI Gate / Evidence"
SetWhole 15 "Rounded Rectangle 24" "Integration Gateway"
SetWhole 15 "TextBox 25" ("Linux control plane " + $dash + " objects, BOM/config rules, ECR/impact, approvals, AI gate, evidence and integration adapters, all exposed through the API layer above.")
SetWhole 15 "TextBox 27" ("Control Plane runs on  " + $harr + "  Linux data-center; Windows authoring connects via Adapter")
SetWhole 15 "TextBox 30" ("3.  INFRASTRUCTURE " + $dash + " Linux control-plane cluster + Windows authoring tier (on-prem, HA, multi-yard, 1,000+ users)")
SetWhole 15 "TextBox 33" ("App " + $mid + " API " + $mid + " AI Gate (Linux)")
$tb47 = "Windows authoring tier: GPU workstations running AVEVA E3D " + $mid + " Marine " + $mid + " Draw, linked to the Linux control plane via Local Agent / Adapter (HTTPS callbacks)." + $CR + "Linux: redundant 10/25 GbE backbone " + $mid + " Geo-DR (async replica) + continuous backup " + $mid + " GPU node for AI gate / digital-twin " + $mid + " observability (metrics " + $mid + " logs " + $mid + " traces)."
SetWhole 15 "TextBox 47" $tb47
SetGeom  15 "TextBox 47" 474 42
SetWhole 15 "TextBox 92" ("Hybrid connection: Windows workstations run AVEVA E3D/Marine authoring " + $rarr + " the Linux data-center hosts the Open PLM Control Plane & secured APIs " + $rarr + " HD Hyundai's client calls them; the E3D/Hull Agent executes commands and returns callbacks, every step status-verified.")
SetGeom  15 "TextBox 92" 512 26

Write-Output "=== SLIDE 13 (Target architecture - Integration Layer) ==="
SetPara 13 "Text 8" 2 ("AVEVA E3D/Marine (Windows authoring) " + $mid + " AIM/PI/CONNECT + NAPA/class/CFD/FEM + ERP/MES/PLM adapters")

Write-Output "=== SLIDE 16 (Phase 2 pilot - event capture via agent) ==="
SetWhole 16 "TextBox 18" ("Agent " + $rarr + " POST /events")

Write-Output "=== SLIDE 20 (E3D version control) ==="
SetWhole 20 "TextBox 75" "Snapshots are captured in the background via a queue on the Linux control-plane servers; AVEVA E3D authoring runs on Windows clients, so many concurrent designers never hit a bottleneck."

Write-Output "=== SLIDE 21 (Naval assurance - design event source) ==="
SetWhole 21 "TextBox 13" "E3D/Marine (Windows) + PLM"

$pres.Save()
Write-Output "SAVED. opened-fresh=$opened"
if ($opened) { $pres.Close(); Write-Output "CLOSED (was opened by script)" } else { Write-Output "LEFT OPEN (user session)" }
