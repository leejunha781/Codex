$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true

# Gather any open copies of this file (stale in-memory) and close WITHOUT saving
$toClose = @()
foreach ($p in $ppt.Presentations) { if ($p.FullName -eq $file) { $toClose += $p } }
foreach ($p in $toClose) {
  try { $p.Saved = $true; $p.Close(); Write-Output "CLOSED stale copy (no save)" } catch { Write-Output "close err: $_" }
}

# Reopen fresh from disk
$pres = $ppt.Presentations.Open($file)
try { $ppt.ActiveWindow.View.GotoSlide(15) } catch {}
try { $ppt.Activate() } catch {}

# Confirm the LIVE (now-displayed) presentation shows the new text
$sl = $pres.Slides.Item(15)
function Live($n){ foreach($sh in $sl.Shapes){ if($sh.Name -eq $n){ return ($sh.TextFrame.TextRange.Text -replace "[\r\n\x0b]+"," | ").Trim() } } return "MISSING" }
Write-Output ("LIVE s15 TextBox 2  : " + (Live "TextBox 2"))
Write-Output ("LIVE s15 RoundRect23: " + (Live "Rounded Rectangle 23"))
Write-Output ("LIVE s15 TextBox 18 : " + (Live "TextBox 18"))
Write-Output ("ReadOnly=" + $pres.ReadOnly + "  Slides=" + $pres.Slides.Count)
Write-Output "RELOAD-DONE"
