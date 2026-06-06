$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$doc = $word.Documents.Open($path, $false, $true)
$sb = New-Object System.Text.StringBuilder
$ti=0
foreach($t in $doc.Tables){
  $ti++
  foreach($c in $t.Range.Cells){
    $bg=$null; try{$bg=$c.Shading.BackgroundPatternColor}catch{}
    if($bg -ne $null -and $bg -ne -16777216){
      $fc=$null; try{$fc=$c.Range.Font.Color}catch{}
      $txt=(($c.Range.Text) -replace "[`r`n`a]","").Trim()
      if($txt.Length -gt 30){$txt=$txt.Substring(0,30)}
      [void]$sb.AppendLine("T$ti bg=$bg font=$fc | $txt")
    }
  }
}
Set-Content -Path "C:\Users\namma\.claude\itt_work\shading.txt" -Value $sb.ToString() -Encoding UTF8
$doc.Close([ref]$false)
if($word.Documents.Count -eq 0){$word.Quit()}
[Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
Write-Output "SHADING_DONE"
