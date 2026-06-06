$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\review_png"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
foreach($p in $ppt.Presentations){ if($p.Name -eq $name){ $pres=$p; break } }
$openedByUs = $false
if($pres -eq $null){ $pres = $ppt.Presentations.Open($v2, -1, 0, 0); $openedByUs = $true }  # ReadOnly, no window

Write-Output ("SLIDE_COUNT=" + $pres.Slides.Count)
Write-Output ("SIZE=" + $pres.PageSetup.SlideWidth + "x" + $pres.PageSetup.SlideHeight)

function DumpShape($sh, $indent){
    $pre = (" " * $indent)
    try {
        if($sh.Type -eq 6){  # msoGroup
            Write-Output ($pre + "[group " + $sh.Name + "]")
            foreach($c in $sh.GroupItems){ DumpShape $c ($indent+2) }
            return
        }
        if($sh.HasTable -eq -1){
            $tb = $sh.Table
            Write-Output ($pre + "[TABLE " + $tb.Rows.Count + "x" + $tb.Columns.Count + "]")
            for($r=1; $r -le $tb.Rows.Count; $r++){
                $cells=@()
                for($c=1; $c -le $tb.Columns.Count; $c++){
                    $tx = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text -replace "[\r\n]"," "
                    $cells += $tx
                }
                Write-Output ($pre + "  R" + $r + ": " + ($cells -join " | "))
            }
            return
        }
        if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){
            $t = $sh.TextFrame.TextRange.Text -replace "\r","  /  " -replace "\n"," "
            Write-Output ($pre + "- " + $t)
        }
    } catch {}
}

for($i=1; $i -le $pres.Slides.Count; $i++){
    $s = $pres.Slides.Item($i)
    Write-Output ""
    Write-Output ("===================== SLIDE " + $i + " =====================")
    foreach($sh in $s.Shapes){ DumpShape $sh 0 }
    $num = if($pres.Slides.Count -ge 10){ "{0:D2}" -f $i } else { "$i" }
    $s.Export("$outDir\slide-$num.png","PNG",1280,720)
}

if($openedByUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "EXTRACT_DONE"
