$ErrorActionPreference = "Stop"
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$dir = Join-Path $proposal.FullName "Proposal"
$file = Join-Path $dir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

# Backup
$bkDir = Join-Path $dir "_backup_20260606_hybridarch"
if (-not (Test-Path $bkDir)) { New-Item -ItemType Directory -Path $bkDir | Out-Null }
$bk = Join-Path $bkDir "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BEFORE_HYBRID.pptx"
Copy-Item $file $bk -Force
Write-Output "BACKUP: $bk  exists=$(Test-Path $bk)"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $ppt.Presentations.Open($file, $true, $false, $false)
$slide = $pres.Slides.Item(15)
Write-Output ("SLIDE15 size W={0} H={1}" -f $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight)
foreach ($sh in $slide.Shapes) {
  $txt = ""
  try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = ($sh.TextFrame.TextRange.Text -replace "[\r\n\x0b]+"," / ").Trim() } } catch {}
  if ($txt.Length -gt 60) { $txt = $txt.Substring(0,60) + "..." }
  Write-Output ("{0,-22} L={1,7:N1} T={2,7:N1} W={3,7:N1} H={4,7:N1}  {5}" -f $sh.Name, $sh.Left, $sh.Top, $sh.Width, $sh.Height, $txt)
}
$pres.Close()
Write-Output "DONE"
