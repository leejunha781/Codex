$ErrorActionPreference = "Stop"
$src = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

function New-PPT {
    for ($i=1; $i -le 6; $i++) {
        try { return (New-Object -ComObject PowerPoint.Application) }
        catch { Start-Sleep -Seconds 4 }
    }
    throw "PowerPoint COM failed to start after retries"
}

function ToHex($bgr) {
    $r = $bgr -band 0xFF; $g = ($bgr -shr 8) -band 0xFF; $b = ($bgr -shr 16) -band 0xFF
    return ("#{0:X2}{1:X2}{2:X2}  (R={3} G={4} B={5})" -f $r,$g,$b,$r,$g,$b)
}

$ppt = New-PPT
try {
    $doc = $ppt.Presentations.Open($src, $true, $false, $false)
    Write-Output ("SlideWidth x Height = " + $doc.PageSetup.SlideWidth + " x " + $doc.PageSetup.SlideHeight)
    foreach ($idx in @(3,8,24)) {
        $sl = $doc.Slides.Item($idx)
        Write-Output "==== SLIDE $idx ===="
        try { Write-Output ("  BG fill type=" + $sl.Background.Fill.Type + "  fore=" + (ToHex $sl.Background.Fill.ForeColor.RGB)) } catch { Write-Output ("  BG: " + $_.Exception.Message) }
        foreach ($sh in $sl.Shapes) {
            $nm = $sh.Name
            $hasText = $false
            try { $hasText = ($sh.HasTextFrame -and $sh.TextFrame.HasText) } catch {}
            $info = "  shape '$nm' type=" + $sh.Type + " L=" + [int]$sh.Left + " T=" + [int]$sh.Top + " W=" + [int]$sh.Width + " H=" + [int]$sh.Height
            if ($hasText) {
                $tr = $sh.TextFrame.TextRange
                $snippet = ($tr.Text -replace "[\r\n\v]"," ").Trim()
                if ($snippet.Length -gt 28) { $snippet = $snippet.Substring(0,28) }
                $font = $tr.Font
                $info += " | TEXT='" + $snippet + "' font='" + $font.Name + "' size=" + $font.Size + " color=" + (ToHex $font.Color.RGB) + " bold=" + $font.Bold
            }
            try { if ($sh.Fill.Visible -eq -1 -and $sh.Type -ne 6) { $info += " | fill=" + (ToHex $sh.Fill.ForeColor.RGB) + " transp=" + [math]::Round($sh.Fill.Transparency,2) } } catch {}
            Write-Output $info
        }
    }
    $doc.Close()
} finally {
    $ppt.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
