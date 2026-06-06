$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\restructure_C_log.txt"
$L = @()

$wdCollapseStart = 1
$wdStyleNormal   = -1
$wdStyleListBullet = -49
$wdBorderBottom  = -3
$wdLineStyleNone = 0

function CleanText($p) { return ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }
function BoldPhrase($p, $phrase) {
    $idx = $p.Range.Text.IndexOf($phrase)
    if ($idx -ge 0) {
        $base = $p.Range.Start
        $r = $p.Range.Duplicate
        $r.Start = $base + $idx
        $r.End   = $base + $idx + $phrase.Length
        $r.Font.Bold = [int]1
        return $true
    }
    return $false
}
function BoldLeadIn($p) {
    $colon = $p.Range.Text.IndexOf(":")
    if ($colon -gt 0) {
        $lead = $p.Range.Duplicate
        $lead.End = $lead.Start + $colon + 1
        $lead.Font.Bold = [int]1
    }
}

# ---- new content (ASCII) ----
$newHarsh = "Harsh-Environment Reliability and Manufacturing Mindset: Proven background in defense, naval, shipboard and satellite systems where reliability, EMI/EMC awareness, environmental robustness, interface stability, mechanical durability, formal documentation and acceptance evidence were essential. Also familiar with ECO/BOM coordination, pilot-build findings, defective-part review, EMC/RoHS support and manufacturability considerations, including the use of high-performance thermoplastics, aluminum and copper alloys and the production requirements needed to utilize them."

$newConn = "ITT Cannon Connector Product and Application Knowledge: Familiar with the harsh-environment connector families relevant to ITT Cannon's markets - D38999-style circular connectors and KJ/KJA/KJB families, miniature circular, Micro-D/Nano-D, and filtered/hermetic types for Aerospace & Defense; rectangular interconnects, CA Bayonet/MIL-DTL-5015-style, high-power, mixed power-signal and EV/customized cable solutions for Transportation, Industrial, Heavy-Vehicle and Energy; and fiber-optic, RF/coaxial, Quadrax-style and high-power-contact options for high-speed and high-power applications. I treat product knowledge as a way to match the customer's real application conditions - environment, cable routing, power/signal/data needs, RF paths and installation constraints - to the right connector, cabling and interconnect solution, rather than catalog recall."

$paraA = "This experience also fits well with ITT's purpose, values and DNA. My naval and defense background is grounded in Impeccable Character - integrity, accurate documentation and verifiable acceptance evidence at every stage. I apply Bold Thinking by going deep into logs, measurements and field data to reach fact-based root causes and by reframing customer pain points into practical, design-in solutions. And I rely on Collective Know-How, working as the technical bridge across sales, R&D, product management, quality and production so the whole team wins the design."

$paraB = "In day-to-day work this shows up as the behaviors ITT's Higher Performance Culture describes: service leadership and accountability in owning outcomes end-to-end; granular, data-based decision-making; speed and simplicity under schedule pressure; honesty and transparency in status reporting; meritocracy and humility in backing the best solution and staying open to learning; and a proud, never-satisfied drive for continuous improvement - guided throughout by ITT's `"We Solve It`" mindset and a relentless focus on customer satisfaction."

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)

    # ---------- A) enhance Harsh-Environment Key Strengths bullet (add materials) ----------
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Harsh-Environment Reliability and Manufacturing")) {
            $r = $p.Range.Duplicate; $r.End = $r.End - 1; $r.Text = $newHarsh
            BoldLeadIn $p
            $L += "A) Harsh-Environment bullet updated (materials folded in)."
            break
        }
    }

    # ---------- B) insert new connector-knowledge bullet before 'Customer Engineering Engagement' ----------
    $nextB = $null
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Customer Engineering Engagement")) { $nextB = $p; break }
    }
    if ($nextB -eq $null) { throw "Customer Engineering Engagement anchor not found" }
    $r = $nextB.Range.Duplicate; $r.Collapse($wdCollapseStart)
    $r.InsertBefore($newConn + "`r")
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("ITT Cannon Connector Product and Application Knowledge")) {
            $p.Style = $doc.Styles.Item($wdStyleListBullet)
            $p.Range.Font.Size = [single]10
            $p.Range.Font.Bold = [int]0
            $p.Range.ParagraphFormat.Borders.Item($wdBorderBottom).LineStyle = $wdLineStyleNone
            BoldLeadIn $p
            $L += "B) Connector product-knowledge bullet inserted into Key Strengths."
            break
        }
    }

    # ---------- C) replace Profile values paragraph with 2 woven paragraphs ----------
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.Contains("also fits well with ITT") -and $t.Contains("culture and DNA")) {
            $r = $p.Range.Duplicate; $r.End = $r.End - 1
            $r.Text = ($paraA + "`r" + $paraB)
            $L += "C) Profile values paragraph replaced with 2 woven paragraphs."
            break
        }
    }
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("This experience also fits well with ITT's purpose")) {
            BoldPhrase $p "Impeccable Character" | Out-Null
            BoldPhrase $p "Bold Thinking" | Out-Null
            BoldPhrase $p "Collective Know-How" | Out-Null
        }
        if ($t.StartsWith("In day-to-day work this shows up")) {
            BoldPhrase $p "We Solve It" | Out-Null
        }
    }

    # ---------- D) delete the two standalone sections (Product/App + Values) ----------
    $startPos = $null; $endPos = $null
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.Contains("Knowledge Integrated with ITT Cannon")) { $startPos = $p.Range.Start }
        if ($t.StartsWith("Education, Certifications, Language")) { $endPos = $p.Range.Start }
    }
    if ($startPos -eq $null -or $endPos -eq $null) { throw "delete bounds not found (start=$startPos end=$endPos)" }
    if ($startPos -ge $endPos) { throw "bad delete range start=$startPos end=$endPos" }
    $del = $doc.Range($startPos, $endPos)
    $del.Delete()
    $L += "D) Deleted Product/Application + Values sections (range $startPos..$endPos)."

    # verify key headings survived
    $hasEdu=$false; $hasCareer=$false; $hasKS=$false
    foreach ($p in $doc.Paragraphs) {
        $t = CleanText $p
        if ($t.StartsWith("Education, Certifications, Language")) { $hasEdu=$true }
        if ($t -eq "Career Summary") { $hasCareer=$true }
        if ($t.StartsWith("Key Strengths Matched to ITT Cannon")) { $hasKS=$true }
    }
    $L += "VERIFY headings: Education=$hasEdu Career=$hasCareer KeyStrengths=$hasKS"

    $pages = $doc.ComputeStatistics(2); $words = $doc.ComputeStatistics(0)
    $L += "RESULT pages=$pages words=$words"

    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "RESTRUCTURE DONE"
