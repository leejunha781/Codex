$ErrorActionPreference = "Stop"

$dir  = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\"
$md   = $dir + "Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.md"
$docx = $dir + "Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.docx"
$pdf  = $dir + "Future_Industrial_PLM_Presenter_Facilitation_Guide_KO.pdf"

# ---- warm Word ----
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 600
foreach($n in @("StartupItems","DocumentRecovery")){ $k="HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\$n"; if(Test-Path $k){ Remove-Item $k -Recurse -Force } }
$null = Start-Process "WINWORD.EXE" -PassThru
Start-Sleep -Seconds 5

$w = $null
for($i=0;$i -lt 6 -and $w -eq $null;$i++){ try{ $w = New-Object -ComObject Word.Application }catch{ Start-Sleep -Milliseconds 1200 } }
if($w -eq $null){ throw "Word COM did not start" }
$w.Visible = $false
$w.DisplayAlerts = 0

$doc = $w.Documents.Add()
$ps = $doc.PageSetup
$ps.TopMargin=[single]56; $ps.BottomMargin=[single]56; $ps.LeftMargin=[single]64; $ps.RightMargin=[single]64

$brand = 6567967   # RGB(31,56,100)
foreach($id in @(-1,-49)){
    try{
        $doc.Styles.Item([int]$id).Font.Name = "Malgun Gothic"
        $doc.Styles.Item([int]$id).Font.NameFarEast = "Malgun Gothic"
        $doc.Styles.Item([int]$id).Font.Size = [single]10.5
    }catch{}
}
foreach($id in @(-63,-2,-3,-4)){ try{ $doc.Styles.Item([int]$id).Font.Color = $brand }catch{} }
foreach($id in @(-63,-2,-3,-4)){
    try{
        $doc.Styles.Item([int]$id).Font.Name = "Malgun Gothic"
        $doc.Styles.Item([int]$id).Font.NameFarEast = "Malgun Gothic"
        $doc.Styles.Item([int]$id).Font.Bold = $true
    }catch{}
}

$sel = $w.Selection
function Set-Style($id){ $sel.Style = $doc.Styles.Item([int]$id) }
function Emit-Inline($text){
    $parts = $text -split '\*\*'
    for($p=0;$p -lt $parts.Count;$p++){
        if($parts[$p].Length -gt 0){
            if($p % 2 -eq 1){ $sel.Font.Bold = $true; $sel.TypeText($parts[$p]); $sel.Font.Bold = $false }
            else { $sel.TypeText($parts[$p]) }
        }
    }
}
function Render($text,$styleId,$italic,$indent){
    Set-Style $styleId
    $sel.ParagraphFormat.LeftIndent = [single]$indent
    if($italic){ $sel.Font.Italic = $true }
    Emit-Inline $text
    if($italic){ $sel.Font.Italic = $false }
    $sel.TypeParagraph()
}

$lines = Get-Content $md -Encoding UTF8
foreach($raw in $lines){
    $line = $raw.TrimEnd()
    if($line -eq "[[PB]]"){ $sel.InsertBreak(7); continue }
    if($line.Trim() -eq ""){ continue }
    if($line.StartsWith("#### ")){ Render $line.Substring(5) -4 $false 0 }
    elseif($line.StartsWith("### ")){ Render $line.Substring(4) -3 $false 0 }
    elseif($line.StartsWith("## ")){ Render $line.Substring(3) -2 $false 0 }
    elseif($line.StartsWith("# ")){ Render $line.Substring(2) -63 $false 0 }
    elseif($line.StartsWith("> ")){ Render $line.Substring(2) -1 $true 18 }
    elseif($line.StartsWith("- ")){ Render $line.Substring(2) -49 $false 0 }
    else { Render $line -1 $false 0 }
}

$sel.Style = $doc.Styles.Item([int]-1)
$sel.ParagraphFormat.LeftIndent = [single]0
try {
    $lastPara = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $lastText = (($lastPara.Text -replace "[\r\n\a\x07]", "") -replace "\s+", "").Trim()
    if($lastText.Length -eq 0){ $lastPara.Delete() | Out-Null }
} catch {}

try {
    $ftr = $doc.Sections.Item(1).Footers.Item(1).Range
    $ftr.Text = "발표 진행 가이드  |  Future Industrial PLM - AVEVA Marine  |  "
    $ftr.Collapse(0) | Out-Null
    $doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add() | Out-Null
} catch {}

$pages = 0
try { $pages = $doc.ComputeStatistics(2) } catch {}

$doc.SaveAs([ref]$docx, [ref]16)
$doc.ExportAsFixedFormat($pdf, 17)
$doc.Close([ref]$false)
$w.Quit()

Write-Output ("PAGES=" + $pages)
Write-Output ("docx=" + (Test-Path $docx) + " pdf=" + (Test-Path $pdf))
