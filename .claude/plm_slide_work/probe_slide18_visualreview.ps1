$ErrorActionPreference = "Stop"
$deck = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_VisualReviewed.pptx"
$out = "C:\Users\namma\.claude\plm_slide_work\visualreview_current\slide18_probe.txt"
$ppt=$null;$pres=$null
$sb=New-Object System.Text.StringBuilder
try {
    $ppt=New-Object -ComObject PowerPoint.Application
    $pres=$ppt.Presentations.Open($deck,$true,$false,$false)
    $sl=$pres.Slides.Item(18)
    foreach($sh in $sl.Shapes){
        $txt=""
        try{if($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1){$txt=($sh.TextFrame.TextRange.Text -replace "[\r\n\v]"," ").Trim()}}catch{}
        [void]$sb.AppendLine(("z={0,3} type={1,2} {2,-24} L={3,6:N1} T={4,6:N1} W={5,6:N1} H={6,6:N1} :: {7}" -f $sh.ZOrderPosition,$sh.Type,$sh.Name,$sh.Left,$sh.Top,$sh.Width,$sh.Height,$txt))
    }
} finally {
    if($pres){try{$pres.Close()}catch{}}
    if($ppt){try{if($ppt.Presentations.Count-eq 0){$ppt.Quit()}}catch{}}
}
[IO.File]::WriteAllText($out,$sb.ToString(),[Text.Encoding]::UTF8)
Write-Output ("OUT="+$out)
