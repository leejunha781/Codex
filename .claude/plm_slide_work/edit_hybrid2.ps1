$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$dir = Join-Path $proposal.FullName "Proposal"
$file = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null; $opened = $false
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $false, $false, $true); $opened = $true }

function Get-Shape($slide, $name) { foreach ($sh in $slide.Shapes) { if ($sh.Name -eq $name) { return $sh } }; return $null }
function SetWhole($slideIdx, $name, $text) {
  $sh = Get-Shape $pres.Slides.Item($slideIdx) $name
  if (-not $sh) { Write-Output "  !! MISSING s$slideIdx [$name]"; return }
  $sh.TextFrame.TextRange.Text = $text
  Write-Output "  OK [$name] ($($text.Length))"
}
function SetFont($slideIdx, $name, $size) {
  $sh = Get-Shape $pres.Slides.Item($slideIdx) $name
  if ($sh) { $sh.TextFrame.TextRange.Font.Size = [single]$size; Write-Output "  font [$name]=$size" }
}
function SetGeom($slideIdx, $name, $top, $height) {
  $sh = Get-Shape $pres.Slides.Item($slideIdx) $name
  if ($sh) { $sh.Top=[single]$top; $sh.Height=[single]$height; Write-Output "  geom [$name] T=$top H=$height" }
}

$mid = [char]0x00B7; $harr = [char]0x2194; $rarr = [char]0x2192; $dash = [char]0x2014; $CR = [string][char]13

Write-Output "=== SLIDE 15 fit-up ==="
SetWhole 15 "TextBox 2"  ("Hybrid: Windows Authoring  " + $harr + "  Linux PLM Control Plane (API)")
SetWhole 15 "TextBox 3"  ("Windows-based AVEVA E3D/Marine authoring + Linux PLM Control Plane (API/SQL/AI Gate/Evidence/Integration) + Cloud/CONNECT")
SetFont  15 "TextBox 3"  12
SetWhole 15 "TextBox 4"  ("A - Windows Authoring  " + $harr + "  Linux Control Plane  " + $harr + "  API")
SetWhole 15 "TextBox 15" ("API issues commands  " + $harr + "  Control Plane returns JSON")
SetWhole 15 "TextBox 18" ("2.  OPEN PLM CONTROL PLANE (Linux) " + $dash + " API & services")
SetWhole 15 "TextBox 27" ("Control Plane on Linux  " + $harr + "  authoring via Adapter")
SetWhole 15 "TextBox 30" ("3.  INFRASTRUCTURE " + $dash + " Linux cluster + Windows authoring tier")
$tb47 = "Windows authoring tier: GPU workstations run AVEVA E3D " + $mid + " Marine " + $mid + " Draw, linked to the Linux control plane via Local Agent / Adapter (not hosted on the cluster)." + $CR + "Linux: redundant 10/25 GbE backbone " + $mid + " Geo-DR (async replica) + continuous backup " + $mid + " GPU node for AI gate / digital-twin " + $mid + " observability (metrics " + $mid + " logs " + $mid + " traces)."
SetWhole 15 "TextBox 47" $tb47
SetFont  15 "TextBox 47" 7.5
SetGeom  15 "TextBox 47" 474 42
SetWhole 15 "TextBox 92" ("Hybrid: Windows runs AVEVA E3D/Marine authoring  " + $mid + "  Linux hosts the Open PLM Control Plane & APIs  " + $mid + "  the E3D/Hull Agent executes commands & returns callbacks.")
SetGeom  15 "TextBox 92" 515 22

$pres.Save()
Write-Output "SAVED"
if ($opened) { $pres.Close(); Write-Output "CLOSED" } else { Write-Output "LEFT OPEN" }
