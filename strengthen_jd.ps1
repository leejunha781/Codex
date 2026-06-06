$ErrorActionPreference = 'Continue'
$logPath = "C:\Users\namma\.claude\itt_work\jd_log.txt"
$log = New-Object System.Text.StringBuilder
function L($m){ [void]$log.AppendLine($m); Set-Content -Path $logPath -Value $log.ToString() -Encoding UTF8 }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

# Append text before paragraph mark using doc.Range(pos,pos).InsertBefore
function AppendToPara($p, $addText){
  $insertPos = $p.Range.End - 1   # position of paragraph mark \r
  $ir = $doc.Range($insertPos, $insertPos)
  $ir.InsertBefore($addText)
}

$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open($path,$false,$false)
L('opened  pages='+$doc.ComputeStatistics(2))

# verify \f gone
$ffLeft=0; foreach($p in $doc.Paragraphs){ if($p.Range.Text -match "`f"){$ffLeft++} }
L("formfeeds_left=$ffLeft")

try {
  $done=@{}

  # Collect all paragraphs into an ArrayList first to avoid COM collection invalidation
  $allParas = New-Object System.Collections.ArrayList
  foreach($p in $doc.Paragraphs){ [void]$allParas.Add($p) }
  L("total paras collected=" + $allParas.Count)

  foreach($p in $allParas){
    $t=$null; try{$t=Norm $p}catch{continue}
    if($t.Length -lt 20){ continue }

    # B1: A&D Background → market trends + product expansion
    if(-not $done['B1'] -and $t.StartsWith('Aerospace & Defense / Military Customer Background:')){
      try{
        AppendToPara $p ' Positioned to track Korean A&D market trends, naval and aerospace platform modernization programs, and ITT Cannon product expansion opportunities in the defense sector.'
        $done['B1']=$true; L('B1 extended OK')
      }catch{ L('B1 fail: '+$_.Exception.Message) }
    }

    # B5: New Design → component/assembly design-in
    if(-not $done['B5'] -and $t.StartsWith('New Design / Product Development Support:')){
      try{
        AppendToPara $p ' Hands-on hardware design experience — ARM Cortex-M3 CPU/I/O boards, sensor and communication interface circuits, embedded control logic — provides practical insight for customer discussions on new component and assembly design-in.'
        $done['B5']=$true; L('B5 extended OK')
      }catch{ L('B5 fail: '+$_.Exception.Message) }
    }

    # B6: Technical Presentations → sales support + full lifecycle
    if(-not $done['B6'] -and $t.StartsWith('Technical Presentations and Solution Proposals:')){
      try{
        AppendToPara $p ' Equally prepared to deliver sales-oriented product capability presentations and to provide sustained technical customer support through the full design-in and post-design lifecycle.'
        $done['B6']=$true; L('B6 extended OK')
      }catch{ L('B6 fail: '+$_.Exception.Message) }
    }

    # B8: Working Style → organizational skills + customer satisfaction
    if(-not $done['B8'] -and $t.StartsWith('Independent, Self-Motivated and Travel-Ready')){
      try{
        AppendToPara $p ' Strong organizational skills to anticipate, prioritize and coordinate multiple concurrent activities — with a relentless focus on customer satisfaction at every stage.'
        $done['B8']=$true; L('B8 extended OK')
      }catch{ L('B8 fail: '+$_.Exception.Message) }
    }

    # Check Values paragraph presence
    if(-not $done['VW'] -and $t.StartsWith('These working habits also map to ITT')){
      $done['VW']=$true; L('Values woven paragraph CONFIRMED')
    }

    if($done.Count -ge 5){ break }
  }
  L("bullets extended: B1="+[bool]$done['B1']+" B5="+[bool]$done['B5']+" B6="+[bool]$done['B6']+" B8="+[bool]$done['B8']+" Values="+[bool]$done['VW'])

  # Verify keywords landed
  L('--- verify keyword search ---')
  $checks=@('market trends, naval and aerospace','new component and assembly design-in','sales-oriented product capability','relentless focus on customer satisfaction','These working habits also map')
  foreach($p in $doc.Paragraphs){
    $t=$null; try{$t=Norm $p}catch{continue}
    foreach($kw in $checks){ if($t.Contains($kw)){ L("  FOUND: $($kw.Substring(0,35))") } }
  }

  $doc.Save(); L('saved')
  try{ $doc.ExportAsFixedFormat($pdf,17); L('pdf exported') }catch{ L('pdf fail: '+$_.Exception.Message) }
  L('FINAL pages='+$doc.ComputeStatistics(2)+' words='+$doc.ComputeStatistics(0))
}
catch { L('FATAL: '+$_.Exception.Message) }
finally {
  try{$doc.Close([ref]$true)}catch{}
  try{ if($word.Documents.Count -eq 0){$word.Quit()} }catch{}
  try{[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null}catch{}
}
Write-Output 'JD_STRENGTHEN_DONE'
