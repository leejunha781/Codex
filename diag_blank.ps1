$ErrorActionPreference = 'Continue'
$out = New-Object System.Text.StringBuilder
function L($m){ [void]$out.AppendLine($m) }
function Norm($p){ return ($p.Range.Text -replace "[`r`n`a]","").Trim() }

$word = New-Object -ComObject Word.Application
$word.Visible = $false; try{$word.DisplayAlerts=0}catch{}
$doc = $word.Documents.Open("D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx",$false,$true)

L("PAGES=" + $doc.ComputeStatistics(2))
$i=0
foreach($p in $doc.Paragraphs){
  $i++
  $t=Norm $p
  $sz=$null;try{$sz=$p.Range.Font.Size}catch{}
  $pb=$false;try{$pb=$p.PageBreakBefore}catch{}
  $spB=0;try{$spB=[math]::Round($p.SpaceBefore,0)}catch{}
  $spA=0;try{$spA=[math]::Round($p.SpaceAfter,0)}catch{}
  $lsr=0;try{$lsr=$p.LineSpacingRule}catch{}
  $ls=0;try{$ls=[math]::Round($p.LineSpacing,0)}catch{}
  # find inline \f (form-feed/page-break) chars in text
  $hasPBchar = $p.Range.Text -match "`f"
  if($i -le 5 -or $t.Length -eq 0 -or $spB -gt 20 -or $spA -gt 20 -or $pb -or $hasPBchar -or $ls -gt 28){
    $disp = if($t.Length -eq 0){"[EMPTY]"} else {$t.Substring(0,[Math]::Min(40,$t.Length))}
    L("P$i sz=$sz spB=$spB spA=$spA lsr=$lsr ls=$ls pb=$pb pfChar=$hasPBchar | $disp")
  }
}
Set-Content -Path "C:\Users\namma\.claude\itt_work\diag.txt" -Value $out.ToString() -Encoding UTF8
$doc.Close([ref]$false)
if($word.Documents.Count -eq 0){$word.Quit()}
[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
Write-Output "DIAG_DONE"
