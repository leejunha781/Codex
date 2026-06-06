$ErrorActionPreference = "Stop"
# force-close any orphan PowerPoint first
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 600

$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

# (find, replace) pairs - all unique substrings, formatting preserved via run-level set
$pairs = @(
  @("Operation Sweets Spot","Operations Sweet Spot"),
  @("Exposee","exposes"),
  @("Stepps","Steps")
)

$changes = New-Object System.Collections.Generic.List[string]

function Fix-Shape($shape, $slideIdx) {
    try {
        if ($shape.Type -eq 6) { foreach ($s in $shape.GroupItems) { Fix-Shape $s $slideIdx }; return }
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            $tr = $shape.TextFrame.TextRange
            $runs = $tr.Runs()
            for ($i=1; $i -le $runs.Count; $i++) {
                $run = $tr.Runs($i)
                $t = $run.Text
                foreach ($p in $pairs) {
                    if ($t -like ("*" + $p[0] + "*")) {
                        $new = $t -replace [regex]::Escape($p[0]), $p[1]
                        $run.Text = $new
                        $changes.Add("slide $slideIdx : '" + $p[0] + "' -> '" + $p[1] + "'")
                        $t = $new
                    }
                }
            }
        }
    } catch {}
}

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $doc = $ppt.Presentations.Open($src, $false, $false, $false)
    foreach ($slide in $doc.Slides) {
        foreach ($shape in $slide.Shapes) { Fix-Shape $shape $slide.SlideIndex }
    }
    $doc.Save()
    $doc.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Write-Output ("CHANGES APPLIED: " + $changes.Count)
$changes | ForEach-Object { Write-Output ("  " + $_) }
