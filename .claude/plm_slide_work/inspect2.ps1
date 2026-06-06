$ErrorActionPreference = "Stop"
$v2 = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$name = "Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
foreach($p in $ppt.Presentations){ if($p.Name -eq $name){ $pres=$p; break } }
$openedByUs = $false
if($pres -eq $null){ $pres = $ppt.Presentations.Open($v2, -1, 0, 0); $openedByUs = $true }

Write-Output ("OPEN_BY_USER=" + (-not $openedByUs))
Write-Output ("SAVED(clean)=" + $pres.Saved)
Write-Output ("COUNT=" + $pres.Slides.Count)
Write-Output ("PATH=" + $pres.FullName)

foreach($idx in @(10,13,14)){
    $s = $pres.Slides.Item($idx)
    Write-Output ""
    Write-Output ("########## SLIDE $idx  shapes=" + $s.Shapes.Count + " ##########")
    $n=0
    foreach($sh in $s.Shapes){
        $n++
        $txt=""
        try { if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){ $txt = ($sh.TextFrame.TextRange.Text -replace "[\r\n]"," ") } } catch {}
        if($txt.Length -gt 60){ $txt = $txt.Substring(0,60)+"..." }
        $line = "[{0}] name='{1}' type={2} L={3:N0} T={4:N0} W={5:N0} H={6:N0} :: {7}" -f $n,$sh.Name,$sh.Type,$sh.Left,$sh.Top,$sh.Width,$sh.Height,$txt
        Write-Output $line
    }
}

if($openedByUs){ $pres.Close() }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Output "INSPECT2_DONE"
