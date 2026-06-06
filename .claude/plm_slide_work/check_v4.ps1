$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant - PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
# Use the actual path with en-dash
$src2 = Get-ChildItem "D:\이력서" -Directory | Where-Object {$_.Name -like "AVEVA*PLM SME*"} | Select-Object -First 1
$file = Join-Path $src2.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Hybrid_Safety_Final_v4.pptx"
Write-Output "FILE: $file"
Write-Output "EXISTS: $(Test-Path $file)"
Write-Output "SIZE: $((Get-Item $file).Length)  MODIFIED: $((Get-Item $file).LastWriteTime)"

# Open via COM to get slide count + any new slides
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $false
try {
  $pres = $ppt.Presentations.Open($file, $true, $false, $false)  # ReadOnly
  Write-Output "SLIDES: $($pres.Slides.Count)"
  # Check last 3 slides for new content
  $n = $pres.Slides.Count
  for ($i = [Math]::Max(1,$n-3); $i -le $n; $i++) {
    $s = $pres.Slides.Item($i)
    $txt = ""
    foreach ($sh in $s.Shapes) {
      try { if ($sh.HasTextFrame -eq -1) { $txt += " | " + $sh.TextFrame.TextRange.Text.Substring(0,[Math]::Min(60,$sh.TextFrame.TextRange.Text.Length)) } } catch {}
    }
    Write-Output "s$i: $($txt.Substring(0,[Math]::Min(200,$txt.Length)))"
  }
  $pres.Close()
} finally {
  if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
}
