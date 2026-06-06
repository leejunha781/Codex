$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$dir = Join-Path $proposal.FullName "Proposal"
$file = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

# fresh backup before this pass
$bkDir = Join-Path $dir "_backup_20260606_hybridarch"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory -Path $bkDir | Out-Null }
Copy-Item $file (Join-Path $bkDir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BEFORE_PASS3.pptx") -Force

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
$opened = $false
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $false, $false, $true); $opened = $true }

function Get-Shape($slide,$name){ foreach($sh in $slide.Shapes){ if($sh.Name -eq $name){ return $sh } }; return $null }
function SetWhole($s,$n,$t){ $sh=Get-Shape $pres.Slides.Item($s) $n; if(-not $sh){Write-Output "  !! MISSING s$s [$n]";return}; $g=("L={0:N0} T={1:N0} W={2:N0} H={3:N0}" -f $sh.Left,$sh.Top,$sh.Width,$sh.Height); $sz=$sh.TextFrame.TextRange.Font.Size; $sh.TextFrame.TextRange.Text=$t; Write-Output "  OK s$s [$n] ($($t.Length)c) font=$sz $g" }

$harr=[char]0x2194; $rarr=[char]0x2192; $dash=[char]0x2014

# Slide 13 - make the hybrid split explicit on the target-architecture concept subtitle
SetWhole 13 "Text 1" ("The three win-above patterns converge into one open, customer-owned AVEVA control plane " + $dash + " Windows authoring + Linux control plane + Cloud/CONNECT")

# Slide 18 - tag the design tools as Windows authoring and the PLM as the Linux control plane
SetWhole 18 "TextBox 96" ("Design in AVEVA (E3D/Marine + Draw, Windows authoring)  " + $rarr + "  governed & tracked in the Linux API-built PLM (control plane), AI-gated at every step")

$pres.Save()
Write-Output "SAVED"
if ($opened) { $pres.Close(); Write-Output "CLOSED" } else { Write-Output "LEFT OPEN" }
