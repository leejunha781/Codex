$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
# close any stale open copies WITHOUT saving
$toClose=@(); foreach($p in $ppt.Presentations){ if($p.FullName -eq $file){ $toClose+=$p } }
foreach($p in $toClose){ try{ $p.Saved=$true; $p.Close(); Write-Output "closed stale copy (no save)" }catch{ Write-Output "close err $_" } }
# reopen fresh from disk
$pres = $ppt.Presentations.Open($file)
try { $ppt.ActiveWindow.View.GotoSlide(15) } catch {}
try { $ppt.Activate() } catch {}
function Live($s,$n){ $sl=$pres.Slides.Item($s); foreach($sh in $sl.Shapes){ if($sh.Name -eq $n){ return ($sh.TextFrame.TextRange.Text -replace "[\r\n\x0b]+"," | ").Trim() } } return "MISSING" }
Write-Output ("s13 Text 1   : " + (Live 13 "Text 1"))
Write-Output ("s15 TextBox 2: " + (Live 15 "TextBox 2"))
Write-Output ("s18 TextBox96: " + (Live 18 "TextBox 96"))
Write-Output ("ReadOnly=" + $pres.ReadOnly + " Slides=" + $pres.Slides.Count + " LastSaved(disk)=" + (Get-Item $file).LastWriteTime)
Write-Output "RELOAD-FINAL-DONE"
