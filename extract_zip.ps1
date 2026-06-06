$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$path = "D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal\Future_Industrial_PLM_Meeting_Deck_EN_Rebuilt_v2.pptx"

$zip = [System.IO.Compression.ZipFile]::OpenRead($path)
try {
    function Read-Entry($name) {
        $e = $zip.Entries | Where-Object { $_.FullName -eq $name } | Select-Object -First 1
        if ($null -eq $e) { return $null }
        $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
        $txt = $sr.ReadToEnd(); $sr.Close()
        return $txt
    }

    # presentation order
    $presXml = [xml](Read-Entry "ppt/presentation.xml")
    $relsXml = [xml](Read-Entry "ppt/_rels/presentation.xml.rels")
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($presXml.NameTable)
    $nsMgr.AddNamespace("p","http://schemas.openxmlformats.org/presentationml/2006/main")
    $nsMgr.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $sldIds = $presXml.SelectNodes("//p:sldIdLst/p:sldId", $nsMgr)

    $relMap = @{}
    foreach ($rel in $relsXml.Relationships.Relationship) { $relMap[$rel.Id] = $rel.Target }

    $sb = New-Object System.Text.StringBuilder
    $order = 0
    foreach ($sid in $sldIds) {
        $order++
        $rid = $sid.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $target = $relMap[$rid]   # e.g. slides/slide5.xml
        $slidePath = "ppt/" + $target
        $slideFile = Split-Path $target -Leaf

        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("################## SLIDE $order  ($slideFile) ##################")

        $sx = Read-Entry $slidePath
        if ($null -eq $sx) { [void]$sb.AppendLine("  (missing $slidePath)"); continue }
        $xml = [xml]$sx
        $sm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $sm.AddNamespace("a","http://schemas.openxmlformats.org/drawingml/2006/main")
        $sm.AddNamespace("p","http://schemas.openxmlformats.org/presentationml/2006/main")

        # Walk shapes in document order. For each sp/pic, collect its paragraphs.
        $spTree = $xml.SelectSingleNode("//p:cSld/p:spTree", $sm)
        function Walk($node, $sm, $sb, $depth) {
            foreach ($child in $node.ChildNodes) {
                $ln = $child.LocalName
                if ($ln -eq "sp" -or $ln -eq "pic" -or $ln -eq "graphicFrame") {
                    # name
                    $nm = ""
                    $nv = $child.SelectSingleNode(".//p:nvSpPr/p:cNvPr | .//p:nvPicPr/p:cNvPr | .//p:nvGraphicFramePr/p:cNvPr", $sm)
                    if ($nv) { $nm = $nv.GetAttribute("name") }
                    $pad = "  " * $depth
                    if ($ln -eq "pic") {
                        [void]$sb.AppendLine("$pad<PICTURE> name='$nm'")
                    } elseif ($ln -eq "graphicFrame") {
                        # could be table
                        $tbl = $child.SelectSingleNode(".//a:tbl", $sm)
                        if ($tbl) {
                            [void]$sb.AppendLine("$pad<TABLE> name='$nm'")
                            foreach ($tr in $tbl.SelectNodes("a:tr", $sm)) {
                                $cells = @()
                                foreach ($tc in $tr.SelectNodes("a:tc", $sm)) {
                                    $ctext = ($tc.SelectNodes(".//a:t", $sm) | ForEach-Object { $_.InnerText }) -join ""
                                    $cells += $ctext
                                }
                                [void]$sb.AppendLine("$pad   ROW | " + ($cells -join " || "))
                            }
                        } else {
                            [void]$sb.AppendLine("$pad<GFX> name='$nm'")
                        }
                    } else {
                        # sp: gather paragraphs
                        $paras = $child.SelectNodes(".//p:txBody/a:p", $sm)
                        $txt = @()
                        foreach ($p in $paras) {
                            $runtext = ($p.SelectNodes("a:r/a:t", $sm) | ForEach-Object { $_.InnerText }) -join ""
                            $txt += $runtext
                        }
                        $joined = ($txt | Where-Object { $_.Trim().Length -gt 0 })
                        if ($joined.Count -gt 0) {
                            [void]$sb.AppendLine("$pad[sp] name='$nm'")
                            foreach ($t in $joined) { [void]$sb.AppendLine("$pad   | $t") }
                        } else {
                            [void]$sb.AppendLine("$pad[sp empty] name='$nm'")
                        }
                    }
                } elseif ($ln -eq "grpSp") {
                    $nv = $child.SelectSingleNode("./p:nvGrpSpPr/p:cNvPr", $sm)
                    $nm = ""; if ($nv) { $nm = $nv.GetAttribute("name") }
                    $pad = "  " * $depth
                    [void]$sb.AppendLine("$pad<GROUP> name='$nm'")
                    Walk $child $sm $sb ($depth+1)
                }
            }
        }
        Walk $spTree $sm $sb 1

        # slide rels -> media images
        $slideRelsPath = "ppt/slides/_rels/" + $slideFile + ".rels"
        $sr = Read-Entry $slideRelsPath
        if ($sr) {
            $rx = [xml]$sr
            $imgs = $rx.Relationships.Relationship | Where-Object { $_.Type -like "*image*" }
            if ($imgs) {
                foreach ($im in $imgs) { [void]$sb.AppendLine("   >> media: $($im.Target)") }
            }
        }

        # notes
        $notesRel = "ppt/slides/_rels/" + $slideFile + ".rels"
        $nr = Read-Entry $notesRel
        if ($nr) {
            $nrx = [xml]$nr
            $noteSlide = $nrx.Relationships.Relationship | Where-Object { $_.Type -like "*notesSlide*" } | Select-Object -First 1
            if ($noteSlide) {
                $ntarget = ($noteSlide.Target -replace "\.\./","ppt/")
                $nx = Read-Entry $ntarget
                if ($nx) {
                    $nxml = [xml]$nx
                    $nm2 = New-Object System.Xml.XmlNamespaceManager($nxml.NameTable)
                    $nm2.AddNamespace("a","http://schemas.openxmlformats.org/drawingml/2006/main")
                    $alltext = ($nxml.SelectNodes("//a:t", $nm2) | ForEach-Object { $_.InnerText }) -join " "
                    if ($alltext.Trim().Length -gt 0) {
                        [void]$sb.AppendLine("   --- NOTES --- " + $alltext.Trim())
                    }
                }
            }
        }
    }

    $out = "C:\Users\namma\.claude\plm_slide_work\v2_full_dump.txt"
    [System.IO.File]::WriteAllText($out, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Output "WROTE: $out  (slides=$order)"
} finally {
    $zip.Dispose()
}
