$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt.pptx"
$outDir = "C:\Users\namma\.claude\plm_slide_work\inspect_png"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
} catch {
    # some versions: assign 1
    try { $ppt.Visible = 1 } catch {}
}
try { $ppt.WindowState = 2 } catch {}  # minimized

# Open read-only, no window, untitled false
$pres = $ppt.Presentations.Open($src, $true, $false, $false)  # ReadOnly, Untitled, WithWindow

Write-Output ("SLIDE_W_PTS=" + $pres.PageSetup.SlideWidth)
Write-Output ("SLIDE_H_PTS=" + $pres.PageSetup.SlideHeight)
Write-Output ("SLIDE_COUNT=" + $pres.Slides.Count)

# Theme colors from slide master
try {
    $scheme = $pres.SlideMaster.Theme.ThemeColorScheme
    for ($i = 1; $i -le 12; $i++) {
        try {
            $rgb = $scheme.Colors($i).RGB
            $r = $rgb -band 0xFF
            $g = ($rgb -shr 8) -band 0xFF
            $b = ($rgb -shr 16) -band 0xFF
            $hex = "{0:X2}{1:X2}{2:X2}" -f $r,$g,$b
            Write-Output ("THEME_COLOR_$i=#$hex")
        } catch {}
    }
} catch { Write-Output "THEME_COLOR_ERR" }

# Major/minor fonts
try {
    Write-Output ("FONT_MAJOR=" + $pres.SlideMaster.Theme.ThemeFontScheme.MajorFont.Item(1).Name)
    Write-Output ("FONT_MINOR=" + $pres.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(1).Name)
} catch { Write-Output "FONT_ERR" }

function Dump-Slide($idx) {
    $s = $pres.Slides.Item($idx)
    Write-Output ("===== SLIDE $idx (layout=" + $s.Layout + ") =====")
    foreach ($sh in $s.Shapes) {
        if ($sh.HasTextFrame -eq -1) {
            if ($sh.TextFrame.HasText -eq -1) {
                $t = $sh.TextFrame.TextRange.Text
                $t = $t -replace "\r"," | "
                if ($t.Length -gt 200) { $t = $t.Substring(0,200) + "..." }
                Write-Output ("  [" + $sh.Name + "] " + $t)
            }
        }
    }
}

Dump-Slide 1
Dump-Slide 2
Dump-Slide ($pres.Slides.Count)

# Export a few slides to PNG for visual style reference
$pres.Slides.Item(1).Export("$outDir\s1.png","PNG",1280,720)
$pres.Slides.Item(2).Export("$outDir\s2.png","PNG",1280,720)
$pres.Slides.Item([math]::Min(3,$pres.Slides.Count)).Export("$outDir\s3.png","PNG",1280,720)
$pres.Slides.Item($pres.Slides.Count).Export("$outDir\slast.png","PNG",1280,720)

$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
[GC]::Collect()
Write-Output "DONE"
