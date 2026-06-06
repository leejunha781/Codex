$ErrorActionPreference = "Stop"
$path = "D:\이력서\ITT Cannon\Joonha_Lee_ITT_Cannon_FAE_Resume_Final_Integrated_v2.docx"
$pdf  = "C:\Users\namma\.claude\itt_work\v2_current.pdf"
$log  = "C:\Users\namma\.claude\itt_work\shorten_ks_log.txt"
$L = @()
function CleanText($p) { return ($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }
function BoldLeadIn($p) {
    $colon = $p.Range.Text.IndexOf(":")
    if ($colon -gt 0) { $lead=$p.Range.Duplicate; $lead.End=$lead.Start+$colon+1; $lead.Font.Bold=[int]1 }
}

$pairs = @(
  @("Aerospace & Defense / Military Customer Background",
    "Aerospace & Defense / Military Customer Background: 15+ years supporting ROK Navy ship and submarine communication systems, plus defense/aerospace electronics and LEO satellite terminals - directly matching ITT Cannon's preferred Korea Aerospace & Defense experience, with insight into Korean A&D market trends and platform-modernization programs."),
  @("Connector, Cable and System Interface Understanding",
    "Connector, Cable and System Interface Understanding: Hands-on with RF/coaxial cabling, antenna interfaces, shipboard wiring, Ethernet/high-speed data, power/control/signal interfaces, board-to-system interconnects and ICDs - fluent in circular and multi-pin connectors, high-density interfaces, cable assemblies, signal integrity, shielding and durability."),
  @("ITT Cannon Connector Product and Application Knowledge",
    "ITT Cannon Connector Product and Application Knowledge: Familiar with ITT Cannon's harsh-environment families - D38999 / KJ-KJA-KJB, Micro-D/Nano-D and filtered/hermetic (Aerospace & Defense); CA Bayonet/MIL-DTL-5015, high-power and EV (Transportation, Industrial, Heavy-Vehicle, Energy); fiber-optic, RF/coaxial and Quadrax (high-speed/high-power). I match the customer's real application to the right interconnect, not the catalog."),
  @("Customer Engineering Engagement / Voice of Customer",
    "Customer Engineering Engagement / Voice of Customer: Proven technical bridge between customer engineering and internal sales, R&D, product management, quality and operations - capturing requirements, clarifying pain points and driving issues transparently to closure."),
  @("Application Issue Solving and Root-Cause Analysis",
    "Application Issue Solving and Root-Cause Analysis: Evidence-based troubleshooting from symptoms, logs, measurements and defect history - isolating causes across hardware, cabling, interface, network, RF and power, with corrective-action and re-test proof."),
  @("New Design / Product Development and Design-In Support",
    "New Design and Design-In Support: Translate customer requirements into architecture, HW/SW/interface specs, test plans, ECO/BOM and production-readiness decisions; hands-on ARM Cortex-M3 board design adds practical insight for new component and assembly design-in."),
  @("Technical Presentations and Solution Proposals",
    "Technical Presentations and Solution Proposals: Skilled at technical proposals, WBS and cost estimates, test procedures and customer-facing reports - equally ready to deliver sales-oriented capability presentations and sustained technical support across the full design-in lifecycle."),
  @("Harsh-Environment Reliability and Manufacturing Mindset",
    "Harsh-Environment Reliability and Manufacturing Mindset: Defense, naval and satellite background in reliability, EMI/EMC, environmental robustness and acceptance evidence - plus ECO/BOM, defective-part review, EMC/RoHS and the use of high-performance thermoplastics, aluminum and copper alloys and their production requirements."),
  @("Independent, Self-Motivated and Travel-Ready Working Style",
    "Independent, Self-Motivated and Travel-Ready: Comfortable with shipyard/onboard field conditions, schedule pressure and cross-functional ownership; strong organizational skills, willing to travel up to 50%, with a relentless focus on customer satisfaction.")
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($path, $false, $false)
    foreach ($pair in $pairs) {
        $prefix = $pair[0]; $newText = $pair[1]; $done = $false
        foreach ($p in $doc.Paragraphs) {
            if ((CleanText $p).StartsWith($prefix)) {
                $r = $p.Range.Duplicate; $r.End = $r.End - 1; $r.Text = $newText
                $p.Range.Font.Bold = [int]0; $p.Range.Font.Size = [single]10
                BoldLeadIn $p
                $done = $true; break
            }
        }
        $L += ("{0} : {1}" -f $(if ($done) { 'OK ' } else { 'MISS' }), $prefix)
    }
    $L += ("RESULT pages=" + $doc.ComputeStatistics(2) + " words=" + $doc.ComputeStatistics(0))
    $doc.Save(); $doc.ExportAsFixedFormat($pdf, 17); $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}
$L | Out-File -FilePath $log -Encoding utf8
$L | ForEach-Object { Write-Output $_ }
Write-Output "SHORTEN DONE"
