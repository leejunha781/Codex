$ErrorActionPreference = 'Stop'
$ppt = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Final.pptx'
$out = 'C:\Users\namma\.claude\plm_slide_work\final_meeting_ready_text_check.txt'
$pp=$null; $pres=$null
function ShapeText($sh){$txt=''; try{if($sh.HasTextFrame -and $sh.TextFrame.HasText){$txt=$sh.TextFrame.TextRange.Text}}catch{}; return (($txt -replace "[`r`n]+", ' ') -replace '\s+', ' ').Trim()}
try{
  $pp=New-Object -ComObject PowerPoint.Application
  $pp.Visible=-1
  $pres=$pp.Presentations.Open($ppt,$false,$true,$false)
  $lines=New-Object System.Collections.Generic.List[string]
  $lines.Add("SLIDES=$($pres.Slides.Count)")
  foreach($i in @(2,13,14,15,29,31,32,33,34,35,36,37)){
    $slide=$pres.Slides.Item($i); $texts=@(); $pics=0; $tables=0
    foreach($sh in @($slide.Shapes)){ try{if($sh.Type -eq 13){$pics++}}catch{}; try{if($sh.HasTable){$tables++}}catch{}; $t=ShapeText $sh; if($t){$texts += $t} }
    $title=if($texts.Count){$texts[0]}else{'<no text>'}
    $lines.Add(('[{0:00}] pics={1} tables={2} title={3}' -f $i,$pics,$tables,$title))
    foreach($t in $texts){ if($t -match 'AI|routing|route|p\.7|p\.16|p\.35|Glossary|Approve|KPI|Deployment boundary'){ $lines.Add('  '+$t) } }
  }
  [System.IO.File]::WriteAllLines($out,$lines,[System.Text.Encoding]::UTF8)
  Get-Content -LiteralPath $out
}
finally{if($pres){$pres.Close()}; if($pp){$pp.Quit()}}
