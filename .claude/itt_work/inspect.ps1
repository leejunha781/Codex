$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Standard_Format.docx"
$doc = $word.Documents.Open($path, $false, $true)

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("PAGES=" + $doc.ComputeStatistics(2))
[void]$sb.AppendLine("WORDS=" + $doc.ComputeStatistics(0))
[void]$sb.AppendLine("PARAS=" + $doc.Paragraphs.Count)
[void]$sb.AppendLine("TABLES=" + $doc.Tables.Count)
[void]$sb.AppendLine("SECTIONS=" + $doc.Sections.Count)

# Page setup
$ps = $doc.PageSetup
[void]$sb.AppendLine(("PAGESETUP top={0} bottom={1} left={2} right={3} pagew={4} pageh={5}" -f $ps.TopMargin,$ps.BottomMargin,$ps.LeftMargin,$ps.RightMargin,$ps.PageWidth,$ps.PageHeight))

# Tables overview
$ti = 0
foreach ($t in $doc.Tables) {
    $ti++
    [void]$sb.AppendLine(("TABLE #{0}: rows={1} cols={2}" -f $ti,$t.Rows.Count,$t.Columns.Count))
}

[void]$sb.AppendLine("----- PARAGRAPH MAP (idx | style | font | size | bold | lineSpacingRule | spaceBefore/After | text) -----")
$i = 0
foreach ($p in $doc.Paragraphs) {
    $i++
    $r = $p.Range
    $fn = ""
    try { $fn = $r.Font.Name } catch {}
    $fs = ""
    try { $fs = $r.Font.Size } catch {}
    $bold = ""
    try { $bold = $r.Font.Bold } catch {}
    $sty = ""
    try { $sty = $p.Style.NameLocal } catch {}
    $lsr = ""
    try { $lsr = $p.LineSpacingRule } catch {}
    $sb4 = ""
    try { $sb4 = "" + [math]::Round($p.SpaceBefore,0) + "/" + [math]::Round($p.SpaceAfter,0) } catch {}
    $txt = ($r.Text -replace '[\r\n\a\t]',' ' -replace '\s+',' ').Trim()
    if ($txt.Length -gt 70) { $txt = $txt.Substring(0,70) }
    [void]$sb.AppendLine(("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}" -f $i,$sty,$fn,$fs,$bold,$lsr,$sb4,$txt))
}

Set-Content -Path "C:\Users\namma\.claude\itt_work\format_report.txt" -Value $sb.ToString() -Encoding UTF8
$doc.Close([ref]$false)
if ($word.Documents.Count -eq 0) { $word.Quit() }
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "INSPECT_DONE"
