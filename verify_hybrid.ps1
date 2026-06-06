$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $null
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $pres = $p; break } }
$opened = $false
if (-not $pres) { $pres = $ppt.Presentations.Open($file, $true, $false, $false); $opened = $true }

function Show($s,$n){ $sl=$pres.Slides.Item($s); foreach($sh in $sl.Shapes){ if($sh.Name -eq $n){ $t=($sh.TextFrame.TextRange.Text -replace "[\r\n\x0b]+"," | ").Trim(); Write-Output ("s$s [$n]  ->  $t"); return } }; Write-Output "s$s [$n] MISSING" }

Show 15 "TextBox 2"; Show 15 "TextBox 3"; Show 15 "TextBox 4"; Show 15 "TextBox 15"
Show 15 "TextBox 18"; Show 15 "Rounded Rectangle 23"; Show 15 "Rounded Rectangle 24"
Show 15 "TextBox 25"; Show 15 "TextBox 27"; Show 15 "TextBox 30"; Show 15 "TextBox 33"
Show 15 "TextBox 47"; Show 15 "TextBox 92"
Show 13 "Text 8"
Show 16 "TextBox 17"; Show 16 "TextBox 18"
Show 20 "TextBox 75"
Show 21 "TextBox 13"; Show 21 "TextBox 14"
if ($opened) { $pres.Close() }
Write-Output "DONE"
