$ErrorActionPreference = 'Stop'
$ppt = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN.pptx'
$out = 'C:\Users\namma\.claude\plm_slide_work\v5_targeted_text_2_6_9_29_34.txt'
$pp = $null; $pres = $null
function ShapeText($sh) { $txt=''; try { if($sh.HasTextFrame -and $sh.TextFrame.HasText){ $txt=$sh.TextFrame.TextRange.Text } } catch {}; return (($txt -replace "[`r`n]+", ' ') -replace '\s+', ' ').Trim() }
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.Visible = -1
  $pres = $pp.Presentations.Open($ppt,$false,$true,$false)
  $lines = New-Object System.Collections.Generic.List[string]
  foreach($i in @(2,6,9,29,30,31,32,33,34)){
    $slide=$pres.Slides.Item($i); $lines.Add("--- SLIDE $i ---")
    foreach($sh in @($slide.Shapes)){
      $txt=ShapeText $sh
      if($txt.Length -gt 0){ $lines.Add("$($sh.Name): $txt") }
    }
  }
  [System.IO.File]::WriteAllLines($out,$lines,[System.Text.Encoding]::UTF8)
  Get-Content -LiteralPath $out
}
finally { if($pres){$pres.Close()}; if($pp){$pp.Quit()} }
