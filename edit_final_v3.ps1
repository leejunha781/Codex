$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\edit_v3_log.txt"
$L = @()

# enums
$wdCollapseStart = 1
$wdStyleNormal   = -1
$wdStyleListBullet = -49
$wdBorderBottom  = -3
$wdLineStyleNone = 0

# ---- content (ASCII only) ----
$H  = "Alignment with ITT Purpose, Values & DNA"
$I  = "My motivation for this role is reinforced by a strong cultural fit with ITT's purpose (`"We Solve It`"), its three DNA principles, and the ten practices of ITT's Higher Performance Culture (`"Who We Are`"). The way I have worked throughout my career maps directly to how ITT describes how its people think, act and win:"
$B1 = "We Solve It - turning customer pain points into solutions: Across naval, defense, aerospace and satellite programs I have repeatedly taken unclear field symptoms and unmet technical challenges and driven them to a verified, customer-accepted resolution - the same mindset ITT brings to its customers' most critical applications, with the goal of making an enduring impact."
$B2 = "Impeccable Character (Work Ethic & Integrity, Honesty & Transparency): I bring dignity and ownership to my work, communicate status candidly even when things go wrong, and back every claim with accurate documentation and verifiable acceptance evidence - FAT/HAT/SAT records, defect logs and corrective-action / re-test data."
$B3 = "Bold Thinking (Granularity & Data-based Decision Making, Passion for Renewal): I go deep into logs, measurements, defect history and environmental data to reach fact-based root causes rather than surface fixes, and I keep the curiosity to challenge the status quo and renew both the product and the process."
$B4 = "Collective Know-How (Service Leadership, Meritocracy): I work as the technical bridge across sales, R&D, product management, quality, production and operations - helping others succeed and supporting the best solution regardless of where it comes from, so the whole team wins the design."
$B5 = "Speed & Simplicity, Proud & Never Satisfied: I keep solutions simple and act decisively under schedule pressure, yet I am never satisfied with good enough - for example, structuring issue analysis to cut monthly defect rate by about 30% and customer-issue lead time by about 25%."
$B6 = "Accountability, Authenticity & Humility: I own outcomes end-to-end, every day; I listen to customers and colleagues, stay open about what I do not yet know, and learn fast - the working posture I would bring as an ITT Cannon FAE."

$bulletTexts = @($B1,$B2,$B3,$B4,$B5,$B6)
$block = (@($H,$I,$B1,$B2,$B3,$B4,$B5,$B6) -join "`r") + "`r"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # ---------- 1) Insert new section before the Education heading ----------
    $anchor = $null
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
        if ($t.StartsWith("Education, Certifications, Language")) { $anchor = $p; break }
    }
    if ($anchor -eq $null) { throw "Education anchor not found" }
    $rng = $anchor.Range.Duplicate
    $rng.Collapse($wdCollapseStart)
    $rng.InsertBefore($block)
    $L += "Inserted block before Education heading."

    # ---------- 2) Restyle inserted intro + bullets ----------
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
        if ($t.StartsWith("My motivation for this role is reinforced")) {
            $p.Style = $doc.Styles.Item($wdStyleNormal)
            $p.Range.Font.Size = [single]10
            $p.Range.Font.Bold = [int]0
            $p.Range.ParagraphFormat.Borders.Item($wdBorderBottom).LineStyle = $wdLineStyleNone
            $L += "Styled intro paragraph."
        }
        foreach ($bt in $bulletTexts) {
            $btHead = $bt.Substring(0, 25)
            if ($t.StartsWith($btHead)) {
                $p.Style = $doc.Styles.Item($wdStyleListBullet)
                $p.Range.Font.Size = [single]10
                $p.Range.Font.Bold = [int]0
                $p.Range.ParagraphFormat.Borders.Item($wdBorderBottom).LineStyle = $wdLineStyleNone
                # bold lead-in up to first colon
                $colon = $p.Range.Text.IndexOf(":")
                if ($colon -gt 0) {
                    $lead = $p.Range.Duplicate
                    $lead.End = $lead.Start + $colon + 1
                    $lead.Font.Bold = [int]1
                }
                $L += ("Styled bullet: {0}..." -f $btHead)
            }
        }
    }

    # ---------- 3) Ensure heading formatting (inherited; enforce) ----------
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
        if ($t -eq $H) {
            $p.Style = $doc.Styles.Item($wdStyleNormal)
            $p.Range.Font.Size = [single]14
            $p.Range.Font.Bold = [int]1
            $L += "Enforced new section heading size/bold (color+border inherited)."
        }
    }

    # ---------- 4) Fix per-company summary boxes (Table 4 = Intellian, Table 5 = Genohco) ----------
    $t4 = $doc.Tables.Item(4)
    $t4.Cell(2,2).Range.Text = "LEO satellite (Eutelsat OneWeb Ku-band) terminal system integration, RF/antenna and network interface validation, software-log-based field issue analysis, production validation and global customer technical support."
    $t4.Cell(3,2).Range.Text = "Acted as the customer-facing technical owner for terminal field issues - capturing symptoms, reproducing conditions, isolating interface boundaries, coordinating internal R&D, quality and production owners, and delivering verified evidence. This mirrors the ITT Cannon FAE loop of solving application and product issues with R&D support and converting field constraints into design-in and design-win opportunities."
    foreach ($rc in @(@(2,2),@(3,2))) { $cc=$t4.Cell($rc[0],$rc[1]).Range; $cc.Font.Size=[single]10; $cc.Font.Bold=[int]0 }
    $L += "Fixed Table 4 (Intellian) Role/FAE Fit."

    $t5 = $doc.Tables.Item(5)
    $t5.Cell(2,2).Range.Text = "Defense/aerospace electronics PL/PM support, RFP analysis, technical proposal preparation, design review, external test coordination and delivery documentation."
    $t5.Cell(3,2).Range.Text = "Converted aerospace and defense customer requirements, RFP items and technical constraints into practical technical proposals, execution plans, interface requirements, test procedures and customer deliverables. This is closely matched to ITT Cannon FAEs who must identify design requirements, understand customer pain points, communicate technical solutions and coordinate with engineering, product management and operations."
    foreach ($rc in @(@(2,2),@(3,2))) { $cc=$t5.Cell($rc[0],$rc[1]).Range; $cc.Font.Size=[single]10; $cc.Font.Bold=[int]0 }
    $L += "Fixed Table 5 (Genohco) Role/FAE Fit."

    # ---------- 5) Tighten materials bullet to echo JD phrasing ----------
    foreach ($p in $doc.Paragraphs) {
        $t = ($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
        if ($t.StartsWith("Materials and Manufacturing Awareness:")) {
            $newText = "Materials and Manufacturing Awareness: Understand that connector selection depends on the use of high-performance thermoplastics, aluminum and copper alloys, contact reliability, sealing, shielding, manufacturability, environmental robustness and the production requirements needed to utilize these materials. My ECO/BOM, pilot-build, defective-part review, EMC/RoHS and production-readiness experience supports this view."
            $r = $p.Range.Duplicate
            $r.End = $r.End - 1   # keep trailing paragraph mark
            $r.Text = $newText
            # re-bold lead-in
            $colon = $r.Text.IndexOf(":")
            if ($colon -gt 0) { $lead=$r.Duplicate; $lead.End=$lead.Start+$colon+1; $lead.Font.Bold=[int]1 }
            $L += "Updated Materials & Manufacturing bullet to JD phrasing."
            break
        }
    }

    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    $L += ("RESULT pages=$pages words=$words")

    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "EDIT DONE"
