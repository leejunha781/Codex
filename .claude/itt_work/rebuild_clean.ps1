$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\rebuild_log.txt"
$L = @()

$wdCollapseStart   = 1
$wdStyleNormal     = -1
$wdStyleListBullet = -49
$wdBorderBottom    = -3
$wdLineStyleNone   = 0

function CleanText($p) { return ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }
function BoldPhrase($p, $phrase) {
    $idx = $p.Range.Text.IndexOf($phrase)
    if ($idx -ge 0) {
        $base = $p.Range.Start
        $r = $p.Range.Duplicate
        $r.Start = $base + $idx
        $r.End   = $base + $idx + $phrase.Length
        $r.Font.Bold = [int]1
    }
}
function BoldLeadIn($p) {
    $colon = $p.Range.Text.IndexOf(":")
    if ($colon -gt 0) { $lead = $p.Range.Duplicate; $lead.End = $lead.Start + $colon + 1; $lead.Font.Bold = [int]1 }
}

# ---------- content ----------
$t4_role = "LEO satellite (Eutelsat OneWeb Ku-band) terminal system integration, RF/antenna and network interface validation, software-log-based field issue analysis, production validation and global customer technical support."
$t4_fit  = "Acted as the customer-facing technical owner for terminal field issues - capturing symptoms, reproducing conditions, isolating interface boundaries, coordinating internal R&D, quality and production owners, and delivering verified evidence. This mirrors the ITT Cannon FAE loop of solving application and product issues with R&D support and converting field constraints into design-in and design-win opportunities."
$t5_role = "Defense/aerospace electronics PL/PM support, RFP analysis, technical proposal preparation, design review, external test coordination and delivery documentation."
$t5_fit  = "Converted aerospace and defense customer requirements, RFP items and technical constraints into practical technical proposals, execution plans, interface requirements, test procedures and customer deliverables. This is closely matched to ITT Cannon FAEs who must identify design requirements, understand customer pain points, communicate technical solutions and coordinate with engineering, product management and operations."

$newHarsh = "Harsh-Environment Reliability and Manufacturing Mindset: Proven background in defense, naval, shipboard and satellite systems where reliability, EMI/EMC awareness, environmental robustness, interface stability, mechanical durability, formal documentation and acceptance evidence were essential. Also familiar with ECO/BOM coordination, pilot-build findings, defective-part review, EMC/RoHS support and manufacturability considerations, including the use of high-performance thermoplastics, aluminum and copper alloys and the production requirements needed to utilize them."

$newConn = "ITT Cannon Connector Product and Application Knowledge: Familiar with the harsh-environment connector families relevant to ITT Cannon's markets - D38999-style circular connectors and KJ/KJA/KJB families, miniature circular, Micro-D/Nano-D, and filtered/hermetic types for Aerospace & Defense; rectangular interconnects, CA Bayonet/MIL-DTL-5015-style, high-power, mixed power-signal and EV/customized cable solutions for Transportation, Industrial, Heavy-Vehicle and Energy; and fiber-optic, RF/coaxial, Quadrax-style and high-power-contact options for high-speed and high-power applications. I treat product knowledge as a way to match the customer's real application conditions - environment, cable routing, power/signal/data needs, RF paths and installation constraints - to the right connector, cabling and interconnect solution, rather than catalog recall."

$paraA = "This experience also fits well with ITT's purpose, values and DNA. My naval and defense background is grounded in Impeccable Character - integrity, accurate documentation and verifiable acceptance evidence at every stage. I apply Bold Thinking by going deep into logs, measurements and field data to reach fact-based root causes and by reframing customer pain points into practical, design-in solutions. And I rely on Collective Know-How, working as the technical bridge across sales, R&D, product management, quality and production so the whole team wins the design."

$paraB = "In day-to-day work this shows up as the behaviors ITT's Higher Performance Culture describes: service leadership and accountability in owning outcomes end-to-end; granular, data-based decision-making; speed and simplicity under schedule pressure; honesty and transparency in status reporting; meritocracy and humility in backing the best solution and staying open to learning; and a proud, never-satisfied drive for continuous improvement - guided throughout by ITT's `"We Solve It`" mindset and a relentless focus on customer satisfaction."

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # ---------- 1) Fix off-by-one summary boxes ----------
    $t4 = $doc.Tables.Item(4)
    $t4.Cell(2,2).Range.Text = $t4_role; $t4.Cell(3,2).Range.Text = $t4_fit
    foreach ($rc in @(@(2,2),@(3,2))) { $cc=$t4.Cell($rc[0],$rc[1]).Range; $cc.Font.Size=[single]10; $cc.Font.Bold=[int]0 }
    $t5 = $doc.Tables.Item(5)
    $t5.Cell(2,2).Range.Text = $t5_role; $t5.Cell(3,2).Range.Text = $t5_fit
    foreach ($rc in @(@(2,2),@(3,2))) { $cc=$t5.Cell($rc[0],$rc[1]).Range; $cc.Font.Size=[single]10; $cc.Font.Bold=[int]0 }
    $L += "1) Tables 4 (Intellian) & 5 (Genohco) corrected."

    # ---------- 2) Key Strengths: update Harsh bullet (+materials) ----------
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Harsh-Environment Reliability and Manufacturing")) {
            $r=$p.Range.Duplicate; $r.End=$r.End-1; $r.Text=$newHarsh; BoldLeadIn $p
            $L += "2) Harsh-Environment bullet updated."; break
        }
    }

    # ---------- 3) Key Strengths: insert connector-knowledge bullet ----------
    $nextB=$null
    foreach ($p in $doc.Paragraphs) { if ((CleanText $p).StartsWith("Customer Engineering Engagement")) { $nextB=$p; break } }
    if ($nextB -eq $null) { throw "anchor 'Customer Engineering Engagement' not found" }
    $r=$nextB.Range.Duplicate; $r.Collapse($wdCollapseStart); $r.InsertBefore($newConn + "`r")
    foreach ($p in $doc.Paragraphs) {
        if ((CleanText $p).StartsWith("ITT Cannon Connector Product and Application Knowledge")) {
            $p.Style=$doc.Styles.Item($wdStyleListBullet); $p.Range.Font.Size=[single]10; $p.Range.Font.Bold=[int]0
            $p.Range.ParagraphFormat.Borders.Item($wdBorderBottom).LineStyle=$wdLineStyleNone
            BoldLeadIn $p; $L += "3) Connector-knowledge bullet inserted."; break
        }
    }

    # ---------- 4) Profile: weave values/DNA into 2 paragraphs ----------
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.Contains("also fits well with ITT") -and $t.Contains("culture and DNA")) {
            $r=$p.Range.Duplicate; $r.End=$r.End-1; $r.Text=($paraA + "`r" + $paraB)
            $L += "4) Profile values paragraph rewoven into 2 paragraphs."; break
        }
    }
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("This experience also fits well with ITT's purpose")) {
            BoldPhrase $p "Impeccable Character"; BoldPhrase $p "Bold Thinking"; BoldPhrase $p "Collective Know-How"
        }
        if ($t.StartsWith("In day-to-day work this shows up")) { BoldPhrase $p "We Solve It" }
    }

    # ---------- 5) Delete standalone Product/Application Knowledge section ----------
    $collect=@(); $inSec=$false
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Education, Certifications, Language")) { break }
        if ($t.StartsWith("Product/Application Knowledge")) { $inSec=$true }
        if ($inSec) { $collect += $p }
    }
    $L += ("5) Collected {0} paragraphs of Product section to delete." -f $collect.Count)
    for ($k=$collect.Count-1; $k -ge 0; $k--) { $collect[$k].Range.Delete() }

    # ---------- verify headings intact ----------
    $expect = @("Professional Profile","Key Strengths Matched to ITT Cannon","Education, Certifications, Language","Career Summary","Detailed Professional Experience","Compensation / Salary Information","Self-Introduction and Application Motivation")
    foreach ($e in $expect) {
        $ok=$false
        foreach ($p in $doc.Paragraphs) { if ((CleanText $p).StartsWith($e)) { $ok=$true; break } }
        $L += ("  HEADING ok={0} : {1}" -f $ok,$e)
    }
    $stillProd=$false
    foreach ($p in $doc.Paragraphs) { if ((CleanText $p).StartsWith("Product/Application Knowledge")) { $stillProd=$true } }
    $L += ("  Product section removed = {0}" -f (-not $stillProd))

    $L += ("RESULT pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))
    $doc.Save(); $doc.ExportAsFixedFormat($pdf, 17); $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "REBUILD DONE"
