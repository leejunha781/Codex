$ErrorActionPreference = "Stop"
$docx = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Presenter_Facilitation_Guide_EN.docx"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500
foreach($n in @("StartupItems","DocumentRecovery")){ $k="HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\$n"; if(Test-Path $k){ Remove-Item $k -Recurse -Force } }
$null = Start-Process "WINWORD.EXE" -PassThru
Start-Sleep -Seconds 4
$w=$null
for($i=0;$i -lt 6 -and $w -eq $null;$i++){ try{$w=New-Object -ComObject Word.Application}catch{Start-Sleep -Milliseconds 1000} }
$w.Visible=$false; $w.DisplayAlerts=0
$d=$w.Documents.Open($docx,$false,$true)  # readonly
$txt=$d.Content.Text
$pages=$d.ComputeStatistics(2)
$paras=$d.Paragraphs.Count

# artifact checks
$artifacts = @{}
$artifacts["stray_**"]    = ([regex]::Matches($txt,'\*\*')).Count
$artifacts["stray_[[PB]]"]= ([regex]::Matches($txt,'\[\[PB\]\]')).Count
$artifacts["stray_####"]  = ([regex]::Matches($txt,'####')).Count
$artifacts["stray_## "]   = ([regex]::Matches($txt,'## ')).Count

# presence checks
$must = @("Presenter's Facilitation Guide","Know your audience","Chairman","win above the tool silos","executable","KPI thresholds","closing","quick reference")
$present = @{}
foreach($m in $must){ $present[$m] = $txt.Contains($m) }

# count styled headings
$h1=0;$h2=0;$h3=0;$title=0
foreach($p in $d.Paragraphs){
  $sn = ""
  try { $sn = $p.Style.NameLocal } catch {}
  if($sn -match "Title|제목$"){ $title++ }
}
$d.Close([ref]$false); $w.Quit()

Write-Output ("PAGES=" + $pages + "  PARAS=" + $paras)
Write-Output "--- artifacts (want 0) ---"
$artifacts.GetEnumerator() | ForEach-Object { Write-Output ("  " + $_.Key + " = " + $_.Value) }
Write-Output "--- required content present ---"
$present.GetEnumerator() | ForEach-Object { Write-Output ("  [" + $_.Value + "] " + $_.Key) }
