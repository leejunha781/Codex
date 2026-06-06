$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated.docx"
$doc = $word.Documents.Open($path, $false, $true)

# Full text
Set-Content -Path "C:\Users\namma\.claude\itt_work\final_text.txt" -Value $doc.Content.Text -Encoding UTF8

# Format / structure map
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("PAGES=" + $doc.ComputeStatistics(2) + " WORDS=" + $doc.ComputeStatistics(0) + " PARAS=" + $doc.Paragraphs.Count + " TABLES=" + $doc.Tables.Count + " SECTIONS=" + $doc.Sections.Count)
$ps = $doc.PageSetup
[void]$sb.AppendLine(("PAGESETUP top={0} bottom={1} left={2} right={3} w={4} h={5}" -f $ps.TopMargin,$ps.BottomMargin,$ps.LeftMargin,$ps.RightMargin,$ps.PageWidth,$ps.PageHeight))
$ti=0
foreach ($t in $doc.Tables){ $ti++; [void]$sb.AppendLine(("TABLE #{0}: rows={1} cols={2}" -f $ti,$t.Rows.Count,$t.Columns.Count)) }
[void]$sb.AppendLine("----- PARA MAP: idx | inTbl | style | font | size | bold | align | LSrule | before/after | textlen | text -----")
$i=0
foreach ($p in $doc.Paragraphs){
  $i++
  $r=$p.Range
  $fn='';try{$fn=$r.Font.Name}catch{}
  $fs='';try{$fs=$r.Font.Size}catch{}
  $bd='';try{$bd=$r.Font.Bold}catch{}
  $st='';try{$st=$p.Style.NameLocal}catch{}
  $al='';try{$al=$p.Alignment}catch{}
  $ls='';try{$ls=$p.LineSpacingRule}catch{}
  $sa='';try{$sa=''+[math]::Round($p.SpaceBefore,0)+'/'+[math]::Round($p.SpaceAfter,0)}catch{}
  $intbl='';try{ if($r.Information(12)){$intbl='T'}else{$intbl='.'} }catch{}
  $raw=$r.Text
  $len=0; if($raw){$len=$raw.Length}
  $txt=($raw -replace "[`r`n`a`t]",' ' -replace '\s+',' ').Trim()
  if($txt.Length -gt 85){$txt=$txt.Substring(0,85)}
  [void]$sb.AppendLine(("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10}" -f $i,$intbl,$st,$fn,$fs,$bd,$al,$ls,$sa,$len,$txt))
}
Set-Content -Path "C:\Users\namma\.claude\itt_work\final_format.txt" -Value $sb.ToString() -Encoding UTF8

$doc.Close([ref]$false)
if ($word.Documents.Count -eq 0) { $word.Quit() }
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "FINAL_EXTRACT_DONE"
