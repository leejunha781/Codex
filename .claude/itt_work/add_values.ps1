$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\values_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$WEAVE = 'These career habits also reflect ITT''s core principles. My defense and naval background is grounded in Impeccable Character - integrity, accurate documentation and verifiable acceptance evidence at every stage. I apply Bold Thinking by taking on unmet technical challenges, reframing field symptoms into root causes and turning customer pain points into practical design-in solutions. And I rely on Collective Know-How - working as the technical bridge across customer engineering, sales, R&D, product management, quality and production to deliver results from concept through production. This is exactly how I would operate as an ITT Cannon FAE, guided by the "We Solve It" purpose.'

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L("opened pages=" + $doc.ComputeStatistics(2))

try {
  # find anchor (section 4 of Self-Intro)
  $anchor=$null
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    if($t -eq '4. Contribution After Joining ITT Cannon'){ $anchor=$p; break }
  }
  if($anchor -ne $null){
    $ir=$doc.Range($anchor.Range.Start,$anchor.Range.Start)
    $ir.InsertBefore($WEAVE+"`r")
    L("Values woven BEFORE '4. Contribution'")
    # format: normal, 10pt, bold on key terms
    foreach($p in $doc.Paragraphs){
      $t=Norm $p
      if($t.StartsWith('These career habits also reflect ITT')){
        $pr=$p.Range; $pr.Font.Name='Arial'; $pr.Font.Size=[single]10; $pr.Font.Bold=[int]0
        $p.SpaceBefore=[single]0; $p.SpaceAfter=[single]3; $p.LineSpacingRule=0
        # bold key terms
        foreach($kw in @('Impeccable Character','Bold Thinking','Collective Know-How','We Solve It')){
          $ix=$p.Range.Text.IndexOf($kw)
          if($ix -ge 0){ $b=$p.Range.Duplicate; $b.Start=$p.Range.Start+$ix; $b.End=$p.Range.Start+$ix+$kw.Length; try{$b.Font.Bold=[int]1}catch{} }
        }
        L("Values para formatted OK")
        break
      }
    }
  } else { L("ANCHOR '4. Contribution' NOT FOUND - searching alternatives") }

  # confirm headings
  L("--- all 14pt bold headings ---")
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
    $bd=$null;try{$bd=$p.Range.Font.Bold}catch{}
    if($sz -eq 14 -and $bd -eq -1 -and $t.Length -gt 0){ L("  H: $t") }
  }

  $doc.Save(); L("saved")
  try{ $doc.ExportAsFixedFormat($pdf,17); L("pdf") }catch{ L("pdf fail: "+$_.Exception.Message) }
  L("FINAL pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))
}
catch{ L("FATAL: "+$_.Exception.Message) }
finally{
  try{$doc.Close([ref]$true)}catch{}
  try{if($word.Documents.Count -eq 0){$word.Quit()}}catch{}
  try{[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
}
Write-Output "VALUES_DONE"
