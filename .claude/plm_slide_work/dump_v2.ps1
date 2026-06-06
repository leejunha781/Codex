$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support – Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
# the real path uses an en-dash; build it from the known folder via Get-ChildItem to avoid typing it
$proposal = Get-ChildItem "D:\이력서" -Directory | Where-Object { $_.Name -like "AVEVA*PLM SME*" } | Select-Object -First 1
$file = Join-Path $proposal.FullName "Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
Write-Output "FILE: $file"
Write-Output "EXISTS: $(Test-Path $file)"

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { $ppt = New-Object -ComObject PowerPoint.Application }
$ppt.Visible = $true
$pres = $ppt.Presentations.Open($file, $true, $false, $false)  # ReadOnly, untitled=false, withWindow=false

$out = New-Object System.Text.StringBuilder
function Walk($shapes, $sb, $indent) {
  foreach ($sh in $shapes) {
    $name = $sh.Name
    $hasText = $false
    try { $hasText = $sh.HasTextFrame -and $sh.TextFrame.HasText } catch {}
    if ($hasText) {
      $t = $sh.TextFrame.TextRange.Text
      $t = ($t -replace "[\r\n\x0b]+"," / ").Trim()
      if ($t.Length -gt 0) { [void]$sb.AppendLine("$indent[$name] $t") }
    }
    # tables
    try {
      if ($sh.HasTable) {
        $tb = $sh.Table
        for ($r=1; $r -le $tb.Rows.Count; $r++) {
          $rowtxt = @()
          for ($c=1; $c -le $tb.Columns.Count; $c++) {
            $ct = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text
            $rowtxt += (($ct -replace "[\r\n\x0b]+"," ").Trim())
          }
          [void]$sb.AppendLine("$indent  TBLROW: " + ($rowtxt -join " | "))
        }
      }
    } catch {}
    # groups
    try {
      if ($sh.Type -eq 6) { Walk $sh.GroupItems $sb "$indent  " }
    } catch {}
  }
}

for ($i=1; $i -le $pres.Slides.Count; $i++) {
  $slide = $pres.Slides.Item($i)
  [void]$out.AppendLine("")
  [void]$out.AppendLine("========== SLIDE $i ==========")
  Walk $slide.Shapes $out ""
}

$pres.Close()
$outPath = "C:\Users\namma\.claude\plm_slide_work\dump_v2.txt"
[System.IO.File]::WriteAllText($outPath, $out.ToString(), (New-Object System.Text.UTF8Encoding $false))
Write-Output "WROTE: $outPath"
Write-Output "SLIDES: $($pres.Slides.Count)"
