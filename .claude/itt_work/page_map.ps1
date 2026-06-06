$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open("D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx",$false,$true)
$out = New-Object System.Text.StringBuilder
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a`f]","").Trim() }

$targets=@(
  'Professional Profile and Interest in ITT Cannon FAE Role',
  'Key Strengths Matched to ITT Cannon FAE Role',
  'ITT Cannon Connector Product Knowledge & Target-Market Fit',
  'Education, Certifications, Language and Technical Toolkit',
  'Career Summary',
  'Detailed Professional Experience',
  'Compensation / Salary Information',
  'Self-Introduction and Application Motivation'
)

foreach($p in $doc.Paragraphs){
  $t=Norm $p
  if($targets -contains $t){
    $pg=$null; try{$pg=$p.Range.Information(3)}catch{}  # wdActiveEndPageNumber
    [void]$out.AppendLine("PAGE=$pg | $t")
  }
}
# Also show any paragraph that lands on page 2
[void]$out.AppendLine("--- page 2 content ---")
foreach($p in $doc.Paragraphs){
  $t=Norm $p
  if($t.Length -eq 0){continue}
  $pg=$null; try{$pg=$p.Range.Information(3)}catch{continue}
  if($pg -eq 2){
    [void]$out.AppendLine("P2: [" + $t.Substring(0,[Math]::Min(60,$t.Length)) + "]")
  }
}
Set-Content -Path "C:\Users\namma\.claude\itt_work\page_map.txt" -Value $out.ToString() -Encoding UTF8
$doc.Close([ref]$false)
if($word.Documents.Count -eq 0){$word.Quit()}
[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
Write-Output "MAP_DONE"
