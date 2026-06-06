$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\final_fix_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L('opened')

try {
  # =========================================================
  # STEP 1: Remove form-feed / hard-page-break characters
  # =========================================================
  $ffRemoved=0
  foreach($p in $doc.Paragraphs){
    if($p.Range.Text -match "`f"){
      $r=$p.Range.Duplicate
      # Replace \f with nothing in this paragraph's range
      $r.Text = ($r.Text -replace "`f","")
      $ffRemoved++
      L("removed \f from para: len=" + (Norm $p).Length)
    }
  }
  L("STEP1 formfeeds removed=$ffRemoved")

  # =========================================================
  # STEP 2: JD strength improvements — append to bullet ends
  # =========================================================
  # Helper: append text to end of a paragraph (before final \r), preserving existing formatting
  function AppendToPara($p, $addText){
    $r=$p.Range.Duplicate
    $r.End=$r.End-1          # exclude paragraph mark
    $r.Collapse(2)           # collapse to end of text content
    $r.InsertAfter($addText)
  }

  $done=@{}

  foreach($p in $doc.Paragraphs){
    $t=Norm $p

    # --- Bullet 1: A&D Background → add market-trend + expansion opportunity language ---
    if(-not $done['B1'] -and $t.StartsWith('Aerospace & Defense / Military Customer Background:') -and $t.Length -gt 50){
      AppendToPara $p ' Positioned to track Korean A&D market trends, naval and aerospace platform modernization programs, and ITT Cannon product expansion opportunities in the defense sector.'
      $done['B1']=$true; L("B1 extended (A&D + market trend)")
    }

    # --- Bullet 5: New Design → add component/assembly design-in language ---
    if(-not $done['B5'] -and $t.StartsWith('New Design / Product Development Support:') -and $t.Length -gt 50){
      AppendToPara $p ' Hands-on hardware design experience — ARM Cortex-M3 CPU/I/O boards, sensor and communication interface circuits, embedded control logic — provides direct insight for customer discussions on new component and assembly design-in.'
      $done['B5']=$true; L("B5 extended (component/assembly design-in)")
    }

    # --- Bullet 6: Technical Presentations → add sales-support + lifecycle language ---
    if(-not $done['B6'] -and $t.StartsWith('Technical Presentations and Solution Proposals:') -and $t.Length -gt 50){
      AppendToPara $p ' Equally ready to deliver sales-oriented product capability presentations and to provide sustained technical customer support through the full design-in and post-design lifecycle.'
      $done['B6']=$true; L("B6 extended (sales support + lifecycle)")
    }

    # --- Bullet 8: Working Style → add organizational skills + customer satisfaction ---
    if(-not $done['B8'] -and $t.StartsWith('Independent, Self-Motivated and Travel-Ready') -and $t.Length -gt 50){
      AppendToPara $p ' Strong organizational skills to anticipate, prioritize and coordinate multiple concurrent activities — with a relentless focus on customer satisfaction at every stage.'
      $done['B8']=$true; L("B8 extended (org skills + customer satisfaction)")
    }

    # --- Verify woven Values paragraph is present ---
    if(-not $done['VW'] -and $t.StartsWith('These working habits also map to ITT')){
      $done['VW']=$true; L("Values paragraph PRESENT (first 50: "+$t.Substring(0,[Math]::Min(50,$t.Length))+")")
    }
  }
  L("STEP2 bullets extended: "+($done.Keys | Where-Object {$_ -ne 'VW'} | Measure-Object).Count+"/4, Values present="+[bool]$done['VW'])

  # =========================================================
  # STEP 3: Save + PDF export + QA
  # =========================================================
  $doc.Save(); L('STEP3 saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('STEP3 pdf') }catch{ L('pdf fail: '+$_.Exception.Message) }
  L('PAGES='+$doc.ComputeStatistics(2)+' WORDS='+$doc.ComputeStatistics(0)+' TABLES='+$doc.Tables.Count)

  # verify no remaining \f
  $ffLeft=0; foreach($p in $doc.Paragraphs){ if($p.Range.Text -match "`f"){$ffLeft++} }
  L("formfeeds remaining=$ffLeft")

  # verify bullets extended
  L('--- bullet verify ---')
  foreach($p in $doc.Paragraphs){
    $t=Norm $p
    foreach($kw in @('market trends, naval and aerospace platform','new component and assembly design-in','sales-oriented product capability','relentless focus on customer satisfaction','These working habits also map')){
      if($t.Contains($kw)){ L('  FOUND: '+$kw.Substring(0,30)) }
    }
  }
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{ $doc.Close([ref]$true) }catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{ [Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }catch{}
}
Write-Output 'FINAL_FIX_DONE'
