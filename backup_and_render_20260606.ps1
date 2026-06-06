$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
$propDir = Split-Path $src
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$bkDir = Join-Path $propDir "_backup_$stamp"
New-Item -ItemType Directory -Force -Path $bkDir | Out-Null
$bk = Join-Path $bkDir ("Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2_BACKUP_$stamp.pptx")
Copy-Item -LiteralPath $src -Destination $bk -Force
Write-Output ("BACKUP -> " + $bk)

$renderDir = "C:\Users\namma\.claude\plm_slide_work\review_20260606"
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null
Get-ChildItem $renderDir -Filter *.png -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $doc = $ppt.Presentations.Open($src, $true, $false, $false)
    foreach ($slide in $doc.Slides) {
        $i = "{0:D2}" -f $slide.SlideIndex
        $f = Join-Path $renderDir ("slide-$i.png")
        $slide.Export($f, "PNG", 1400, 788)
    }
    Write-Output ("RENDERED " + $doc.Slides.Count + " slides -> " + $renderDir)
    $doc.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
