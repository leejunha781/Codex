$ErrorActionPreference = "Stop"
$path = $args[0]
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
    $pres = $ppt.Presentations.Open($path, 0, 0, 0)  # ReadOnly, Untitled, WithWindow=msoFalse
    Write-Output "=== DECK: $($pres.Name) ==="
    Write-Output "=== SLIDE COUNT: $($pres.Slides.Count) ==="
    Write-Output "=== SLIDE SIZE: $($pres.PageSetup.SlideWidth) x $($pres.PageSetup.SlideHeight) ==="
    Write-Output ""

    function Dump-Shape($shp, $slideNum, $indent) {
        $prefix = "  " * $indent
        $type = $shp.Type
        $name = $shp.Name
        # Group
        if ($shp.Type -eq 6) {
            Write-Output "$prefix[GROUP: $name]"
            foreach ($sub in $shp.GroupItems) { Dump-Shape $sub $slideNum ($indent+1) }
            return
        }
        # Picture
        if ($shp.Type -eq 13) {
            Write-Output "$prefix[PICTURE: $name  W=$([math]::Round($shp.Width))pt H=$([math]::Round($shp.Height))pt  L=$([math]::Round($shp.Left)) T=$([math]::Round($shp.Top))]"
            return
        }
        # Table
        if ($shp.HasTable) {
            Write-Output "$prefix[TABLE: $name]"
            $tb = $shp.Table
            for ($r=1; $r -le $tb.Rows.Count; $r++) {
                $rowtext = @()
                for ($c=1; $c -le $tb.Columns.Count; $c++) {
                    $t = $tb.Cell($r,$c).Shape.TextFrame.TextRange.Text -replace "[\r\n\v]"," / "
                    $rowtext += $t
                }
                Write-Output "$prefix  ROW${r}: $($rowtext -join '  |  ')"
            }
            return
        }
        # Text
        if ($shp.HasTextFrame -and $shp.TextFrame.HasText) {
            $t = $shp.TextFrame.TextRange.Text
            $t = $t -replace "\r","`n$prefix    "
            Write-Output "$prefix[TXT: $name]"
            Write-Output "$prefix    $t"
        }
    }

    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        Write-Output "##################################################"
        Write-Output "########## SLIDE $i ##########"
        Write-Output "##################################################"
        foreach ($shp in $slide.Shapes) {
            Dump-Shape $shp $i 0
        }
        # Speaker notes
        try {
            $notes = $slide.NotesPage.Shapes | Where-Object { $_.HasTextFrame -and $_.TextFrame.HasText }
            foreach ($n in $notes) {
                $nt = $n.TextFrame.TextRange.Text
                if ($nt.Trim().Length -gt 0 -and $n.PlaceholderFormat.Type -ne 12) {
                    Write-Output "  [NOTES]: $($nt -replace "`r"," ")"
                }
            }
        } catch {}
        Write-Output ""
    }
}
finally {
    if ($pres -ne $null) { $pres.Close() }
    if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
}
