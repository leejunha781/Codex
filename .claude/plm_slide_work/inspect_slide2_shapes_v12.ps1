$ErrorActionPreference = 'Stop'

$deck = 'D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_V12.pptx'
$out = 'C:\Users\namma\.claude\plm_slide_work\slide2_shapes_v12.txt'

function Get-PowerPointApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        try {
            return (New-Object -ComObject PowerPoint.Application)
        } catch {
            $exe = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'
            Start-Process -FilePath $exe -WindowStyle Hidden
            Start-Sleep -Seconds 8
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        }
    }
}

$ppt = $null
$pres = $null
try {
    $ppt = Get-PowerPointApplication
    $pres = $ppt.Presentations.Open($deck, $true, $false, $false)
    $slide = $pres.Slides.Item(2)
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("SlideWidth=$($pres.PageSetup.SlideWidth) SlideHeight=$($pres.PageSetup.SlideHeight)")
    foreach ($shape in $slide.Shapes) {
        try {
            $txt = ''
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                $txt = $shape.TextFrame.TextRange.Text
                $txt = ($txt -replace "`r", " | " -replace "`v", " | " -replace "\s+", " ").Trim()
            }
            if ($txt.Length -gt 0 -or $shape.Name -match 'Glossary|Pointer|Nav|Meeting|Line|Rule') {
                [void]$lines.Add(("{0}`tType={1}`tL={2:N1}`tT={3:N1}`tW={4:N1}`tH={5:N1}`tTEXT={6}" -f $shape.Name, $shape.Type, $shape.Left, $shape.Top, $shape.Width, $shape.Height, $txt))
            }
        } catch {
            [void]$lines.Add("ERROR $($shape.Name): $($_.Exception.Message)")
        }
    }
    [System.IO.File]::WriteAllLines($out, $lines.ToArray(), [System.Text.Encoding]::UTF8)
    Write-Host ([string]::Join("`n", $lines.ToArray()))
} finally {
    if ($pres -ne $null) { try { $pres.Close() } catch {} }
    if ($ppt -ne $null) {
        try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}
    }
    if ($pres -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null 2>$null }
    if ($ppt -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null 2>$null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
