$ErrorActionPreference = "Stop"
$prop = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
"==== Proposal .pptx files (recent) ===="
Get-ChildItem -Path $prop -Filter *.pptx -File | Sort-Object LastWriteTime -Descending | Select-Object -First 8 |
  ForEach-Object { "{0,7:N1}MB  {1}  {2}" -f ($_.Length/1MB), $_.LastWriteTime, $_.Name }
"==== recent _backup folders ===="
Get-ChildItem -Path $prop -Directory -Filter "_backup*" | Sort-Object LastWriteTime -Descending | Select-Object -First 6 | ForEach-Object { $_.Name }

$src = Join-Path $prop "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($src, $true, $false, $false)
"==== ACTIVE: Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx ===="
$n = $pres.Slides.Count
"Total slides: $n"
# count pictures + tables across deck
$pics=0;$tables=0
foreach($sl in $pres.Slides){ foreach($sh in $sl.Shapes){ if($sh.Type -eq 13){$pics++}; try{ if($sh.HasTable -eq -1){$tables++} }catch{} } }
"Picture shapes: $pics   Tables: $tables"
# titles of last 4 slides (glossary region)
for($i=[math]::Max(1,$n-3); $i -le $n; $i++){
  $t=""; try{ $t=$pres.Slides.Item($i).Shapes.Item(1).TextFrame.TextRange.Text }catch{}
  if([string]::IsNullOrWhiteSpace($t)){ foreach($sh in $pres.Slides.Item($i).Shapes){ try{ if($sh.TextFrame.HasText -eq -1){ $t=$sh.TextFrame.TextRange.Text; break } }catch{} } }
  "  slide ${i}: $(($t -replace '[\r\n]',' '))"
}
# slide 2: count shapes, look for chips / glossary pointer text
"---- slide 2 text shapes containing 'p.' or 'Glossary' or 'Abbrev' ----"
foreach($sh in $pres.Slides.Item(2).Shapes){
  try{ if($sh.TextFrame.HasText -eq -1){ $tx=$sh.TextFrame.TextRange.Text -replace '[\r\n]',' '
    if($tx -match 'p\.\d|Glossary|Abbrev'){ "  [$($sh.Name)] $tx" } } }catch{}
}
# slide 18: background fill type + picture count + window panel transparency
$s18=$pres.Slides.Item(18)
$p18=0; foreach($sh in $s18.Shapes){ if($sh.Type -eq 13){$p18++} }
$r6=$null; foreach($sh in $s18.Shapes){ if($sh.Name -eq 'Rectangle 6'){ $r6=$sh } }
$r6t = if($r6){ [math]::Round($r6.Fill.Transparency,2) } else { "n/a" }
"---- slide 18: bgFillType=$($s18.Background.Fill.Type)  pictures=$p18  'Rectangle 6' transp=$r6t  shapes=$($s18.Shapes.Count)"
$pres.Close()
if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
