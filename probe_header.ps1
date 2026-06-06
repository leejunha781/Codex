$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"
try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch { $ppt = New-Object -ComObject PowerPoint.Application }
$pres = $ppt.Presentations.Open($src, $true, $false, $false)

function Dump($sl, $idx) {
  $s = $pres.Slides.Item($idx)
  Write-Output "=== SLIDE $idx ==="
  foreach ($sh in $s.Shapes) {
    $t = ""
    try { if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) { $t = $sh.TextFrame.TextRange.Text } } catch {}
    $t = ($t -replace "[\r\n]"," ")
    if ($t.Length -gt 34) { $t = $t.Substring(0,34) }
    $fn=""; $fs=""; $fb=""; $fc=""
    try {
      $tr = $sh.TextFrame.TextRange
      $fn = $tr.Font.Name; $fs = $tr.Font.Size; $fb = $tr.Font.Bold; $fc = $tr.Font.Color.RGB
    } catch {}
    $line = "{0,-14} L={1,4} T={2,4} W={3,4} H={4,3} | font={5} sz={6} b={7} colorBGR={8} | '{9}'" -f `
      $sh.Name, [int]$sh.Left, [int]$sh.Top, [int]$sh.Width, [int]$sh.Height, $fn, $fs, $fb, $fc, $t
    Write-Output $line
  }
}
Dump $pres 8
Write-Output ""
# master background color
try {
  $bg = $pres.SlideMaster.Background.Fill.ForeColor.RGB
  Write-Output ("MASTER BG BGR = " + $bg)
} catch { Write-Output "master bg n/a" }
try {
  $lbg = $pres.Slides.Item(8).Background.Fill.ForeColor.RGB
  Write-Output ("SLIDE8 BG BGR = " + $lbg)
} catch { Write-Output "slide8 bg n/a" }
# layout name of slide 8
try { Write-Output ("SLIDE8 layout = " + $pres.Slides.Item(8).CustomLayout.Name) } catch {}

# Sample pixel colors from renders for navy + accents
Add-Type -AssemblyName System.Drawing
function Px($file,$x,$y){
  $bmp = New-Object System.Drawing.Bitmap($file)
  $c = $bmp.GetPixel($x,$y)
  $bmp.Dispose()
  return ("R={0} G={1} B={2}" -f $c.R,$c.G,$c.B)
}
$r = "C:\Users\namma\.claude\plm_slide_work\render_v2"
Write-Output ("navy corner slide08 = " + (Px "$r\slide08.png" 6 700))
Write-Output ("cyan dot slide04   = " + (Px "$r\slide04.png" 480 226))
Write-Output ("orange num slide02 = " + (Px "$r\slide02.png" 98 444))
Write-Output ("title white slide08= " + (Px "$r\slide08.png" 60 42))

$pres.Close()
Write-Output "PROBE DONE"
