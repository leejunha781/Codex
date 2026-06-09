# Build 최종_v2: targeted JD-alignment + defect fixes on the MoaComm resume
$ErrorActionPreference = "Stop"
$src = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종.docx"
$dst = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.docx"
$pdf = "D:\이력서\모아컴코리아 - HW 개발\이준하_모아컴코리아_HW설계_최종_v2.pdf"

function NormText([string]$t) { ($t -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim() }

# Edits: Kind = Partial (regex replace inside), Full (replace whole paragraph text), Delete (remove text)
$edits = @(
    @{ Id="E1-career-date";   Match="엑시큐어하이트론";                          Match2="2004년 01월"; Kind="Partial"; From="2004년 01월"; To="2006년 01월" },
    @{ Id="E2-trip-date";     Match="호주 OneWeb 평판안테나";                     Kind="Partial"; From="2025년 08월"; To="2024년 08월" },
    @{ Id="E4-leadership";    Match="소규모 고기술 조직";                         Kind="Partial"; From="소규모 고기술 조직에서 요구되는"; To="소수 정예 기술조직에 적합한" },
    @{ Id="E5-daeyang-kw";    Match="대양전기공업 약 15년";                       Kind="Partial"; From="CPU/I/O Board 설계"; To="CPU/I/O Board·Flash/Memory 회로 설계" },
    @{ Id="E6a-intellian-kw"; Match="인텔리안: LEO ESA 안테나의 제어전자부";       Kind="Partial"; From="제어전자부, DFE/ADC-DAC, 네트워크"; To="제어전자부, FPGA/DFE·ADC/DAC, 네트워크" },
    @{ Id="E6b-intellian-kw"; Match="인텔리안: LEO ESA 안테나의 제어전자부";       Kind="Partial"; From="field issue 및 양산 안정화"; To="field issue 분석 및 양산 안정화" },
    @{ Id="E7-int-role";      Match="담당 분야 : Ku-band LEO";                    Kind="Full";
       To="담당 분야 : Ku-band LEO 위성통신 단말/ESA 안테나 제어·신호처리 모듈의 통합검증 및 Digital H/W 설계 분석·디버깅, 양산 전환 개발관리, 글로벌 고객 대응" },
    @{ Id="E8-int-leave";     Match="방산·5G RF/MW Digital H/W 설계 직무에서 더 직접적으로"; Kind="Full";
       To="인텔리안테크놀로지스에서는 Eutelsat OneWeb 대상 LEO 위성통신 단말 및 ESA 안테나의 안테나 제어·신호처리 모듈 통합검증, field issue 분석, 양산검증, 글로벌 고객 대응을 수행하며 RF 장비 Digital H/W 분석 역량을 축적하였습니다. 이후에는 ESA 안테나 제어, Link/Power Budget 해석, FPGA/DFE·CPU 제어부 분석, HW-RF 통합검증으로 다져 온 역량을 방산·5G RF 장비의 Digital H/W 설계와 개발관리 직무에서 보다 직접적으로 발휘하고자 퇴직을 결정하였습니다." },
    @{ Id="E9-dy-role";       Match="담당 분야: 국방 통신·C4I 시스템 PM/PL";       Kind="Full";
       To="담당 분야: 국방 통신·C4I 시스템 PM/PL, CPU/I/O Board·Flash/Memory 주변회로·통신 인터페이스 회로 및 PCB 설계, RF 통신장비 시험 및 시스템 통합 인수 총괄" },
    @{ Id="E10-dy-c4i";       Match="MOSCOS 연동 및 LINK-11";                     Kind="Full";
       To="MOSCOS 연동 및 LINK-11 기반 C4I 전술데이터링크 검토를 수행하며 데이터 무결성, 지연, 통신 인터페이스, 네트워크 안정성, 현장 운용성을 검증하였습니다. 안테나, 통신장비, 네트워크가 결합된 체계에서 신호처리와 인터페이스를 검증한 이 경험은 Digital&RF Module, Repeater/NMS, 안테나 제어·신호처리 모듈의 시스템 통합검증에 바로 적용 가능합니다." },
    @{ Id="E11-dy-mgmt";      Match="설계 부장급 직무에서 요구하는";               Kind="Full";
       To="프로젝트 진행회의, 품질 이슈 회의, 고객/조선소/협력사 기술회의를 운영하며 변경관리, defect log, risk register, 일정·품질·납기 조율을 수행하였습니다. 이는 모아컴코리아 HW 설계 직무가 요구하는 Digital H/W 설계 분석·디버깅 및 개발관리 역량과 직접 맞닿아 있는 경험입니다." },
    @{ Id="E12-salary";       Match="회사 보상 밴드 및 직무·역량 수준에 따라 협의"; Kind="Full"; To="회사 내규에 따라 협의" }
)
$applied = @{}
$edits | ForEach-Object { $applied[$_.Id] = 0 }
$dotsDeleted = 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($src, $false, $true)   # ReadOnly (user has original open)
    $doc.SaveAs2($dst)                                  # detach to new path, becomes writable
    Write-Output "SavedAs: $dst"

    foreach ($p in $doc.Paragraphs) {
        $raw = $p.Range.Text
        $norm = NormText $raw

        # stray lone period cleanup
        if ($norm -eq ".") {
            $r = $p.Range.Duplicate
            $r.End = $r.End - 1
            if ($r.End -gt $r.Start) { $r.Delete() | Out-Null; $dotsDeleted++ }
            continue
        }

        foreach ($e in $edits) {
            if ($norm -notlike ("*" + $e.Match + "*")) { continue }
            if ($e.ContainsKey("Match2") -and ($norm -notlike ("*" + $e.Match2 + "*"))) { continue }

            $r = $p.Range.Duplicate
            # last paragraph in a table cell ends with CR + Chr(7): exclude both
            if ($raw -match "\x07$") { $r.End = $r.End - 2 } else { $r.End = $r.End - 1 }

            if ($e.Kind -eq "Partial") {
                $cur = $r.Text
                if ($cur -notlike ("*" + $e.From + "*")) { continue }
                $r.Text = $cur.Replace($e.From, $e.To)
            } else {
                $r.Text = $e.To
            }
            $applied[$e.Id]++
            Write-Output ("APPLIED {0}" -f $e.Id)
        }
    }

    Write-Output "DotsDeleted: $dotsDeleted"
    $edits | ForEach-Object { if ($applied[$_.Id] -ne 1) { Write-Output ("WARN {0} applied {1} times" -f $_.Id, $applied[$_.Id]) } }

    # global conformance check: 맑은 고딕 + single spacing
    $badFont = 0; $badLsr = 0
    foreach ($p in $doc.Paragraphs) {
        $t = (NormText $p.Range.Text)
        if ($t.Length -gt 0) {
            $fn = $p.Range.Font.Name
            if ($fn -ne "맑은 고딕" -and $fn -ne "") { $badFont++; Write-Output ("FONT? [{0}] {1}" -f $fn, $t.Substring(0,[Math]::Min(40,$t.Length))) }
        }
        if ($p.Format.LineSpacingRule -ne 0) { $badLsr++; Write-Output ("LSR? [{0}] {1}" -f $p.Format.LineSpacingRule, $t.Substring(0,[Math]::Min(40,$t.Length))) }
    }
    Write-Output "BadFont:$badFont BadLsr:$badLsr"

    # verify all section headings survived
    $expected = @("기본사항","학력 사항","경력 사항","병역 사항","상세 경력 사항","주요성과","특허 / 논문 및 수상내역","자격증 / 교육사항 / OA / 외국어","해외연수 및 출장경력","주요 활동사항","연봉 정보","자기소개서")
    $headings = @()
    foreach ($p in $doc.Paragraphs) {
        if ($p.Range.Font.Size -eq 14) { $h = NormText $p.Range.Text; if ($h.Length -gt 0) { $headings += $h } }
    }
    foreach ($x in $expected) {
        $hit = $headings | Where-Object { $_ -like ("*" + $x + "*") }
        if (-not $hit) { Write-Output ("HEADING MISSING: {0}" -f $x) }
    }
    Write-Output ("Headings: {0}" -f ($headings -join " | "))

    $doc.Save()
    Write-Output ("PAGES: {0}  WORDS: {1}" -f $doc.ComputeStatistics(2), $doc.ComputeStatistics(0))
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output "PDF: $pdf"
    $doc.Close([ref]$false)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
