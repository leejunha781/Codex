$ErrorActionPreference = "Stop"

$SourcePath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정_영업지원보강.docx"
$OutPath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정_영업지원보강_키워드보강.docx"
$PdfPath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정_영업지원보강_키워드보강.pdf"

Copy-Item -LiteralPath $SourcePath -Destination $OutPath -Force

$replacements = New-Object System.Collections.Generic.List[object]
function Add-Replacement([string]$Old, [string]$New) {
    $replacements.Add([pscustomobject]@{ Old = $Old; New = $New }) | Out-Null
}

Add-Replacement @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 21년 이상 제품개발, 시스템 통합, 품질 안정화, 고객 인수, 양산이관을 수행해 온 무선통신 시스템 R&D 리더입니다. 직접 최대 7명과 8명 규모 TFT를 리딩했고, RF/HW/SW/기구/품질/생산/협력사를 연결해 고객 요구를 제품개발 우선순위와 검증근거로 전환해 왔습니다. 특히 대양전기공업 재직 시 영업본부와 함께 매 신규 특수선 영업 단계마다 기술구성안, 개발비·공수·자재 기반 원가 산출, 견적 및 제안자료를 지원하며 R&D 성과를 수주 경쟁력과 사업성과로 연결했습니다. 통신장비 연구소장으로서 개발 Gate, 품질·신뢰성 개선, 양산준비, 원가·수주 관점의 사업기여를 하나의 운영체계로 묶어 실행할 수 있습니다.
'@ @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 21년 이상 제품개발, 시스템 통합, 품질 안정화, 고객 인수, 양산이관을 수행해 온 무선통신 시스템 R&D 리더입니다. 직접 최대 7명과 8명 규모 TFT를 리딩했고, 복잡한 기술·품질·사업 문제를 정리하여 RF/HW/SW/기구/품질/생산/영업/고객 사이의 중간 번역자 역할을 수행했습니다. 고객이 명확히 표현하지 못한 잠재 기술 요구사항을 끌어내 제품개발 우선순위와 검증근거로 전환했고, 장애 원인 분석과 재발 방지 구조를 통해 품질 신뢰성을 높였습니다. 특히 대양전기공업 재직 시 영업본부와 함께 매 신규 특수선 영업 단계마다 기술구성안, 개발비·공수·자재 기반 원가 산출, 견적 및 제안자료를 지원하며 기술을 사업 기회와 수주 경쟁력으로 연결했습니다.
'@

Add-Replacement @'
R&D총괄 제품개발 기술전략 제품로드맵 무선통신장비 방산통신 위성통신 LEO통신 ESA안테나 RF/MW Digital H/W FPGA/DFE CPU Board Flash/Memory ADC/DAC Ethernet 시스템아키텍처 요구사항관리 ICD 설계검토 개발Gate FAT/HAT/SAT 고객인수 양산이관 ECO/BOM 품질개선 신뢰성개선 Field Issue Root Cause Analysis Corrective Action 협력사관리 외주개발 개발비산정 원가검증 견적산정 특수선영업지원 영업본부협업 수주지원 방위사업청원가검증 WBS PMBOK 조직리딩 임원보고 글로벌고객대응 Technical Review
'@ @'
R&D총괄 제품개발 기술전략 제품로드맵 무선통신장비 방산통신 위성통신 LEO통신 ESA안테나 RF/MW Digital H/W FPGA/DFE CPU Board Flash/Memory ADC/DAC Ethernet 시스템아키텍처 요구사항관리 잠재요구도출 복잡한문제정리 기술요구사항번역 중간번역자 ICD 설계검토 개발Gate FAT/HAT/SAT 고객인수 양산이관 ECO/BOM 장애원인분석 재발방지구조 품질개선 신뢰성개선 Field Issue Root Cause Analysis Corrective Action 협력사관리 외주개발 개발비산정 원가검증 견적산정 특수선영업지원 영업본부협업 수주지원 기술사업화 방위사업청원가검증 WBS PMBOK 조직리딩 임원보고 글로벌고객대응 Technical Review
'@

Add-Replacement @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 제품개발, 시스템 통합, 검증, 품질개선, 고객 인수, 양산이관을 수행했습니다. 연구소장 직무에 필요한 기술 판단력은 현장 이슈를 제품개선 과제로 전환해 본 경험에서 나오며, 조직 운영 역량은 HW/FW/SW/RF/기구/품질/생산/협력사를 하나의 목표와 일정으로 정렬해 온 경험에서 나옵니다.
'@ @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 제품개발, 시스템 통합, 검증, 품질개선, 고객 인수, 양산이관을 수행했습니다. 연구소장 직무에 필요한 기술 판단력은 복잡한 문제를 구조화하고, 고객이 명확히 표현하지 못하는 기술 요구사항을 끌어내며, 현장 이슈를 제품개선과 사업기회로 전환해 본 경험에서 나옵니다. 조직 운영 역량은 개발·품질·생산·영업·고객 사이의 중간 번역자로서 각 조직의 언어를 하나의 목표와 일정, 검증 기준으로 정렬해 온 경험에서 나옵니다.
'@

Add-Replacement @'
• FFX Batch-II, KSS-III, 인도네시아 잠수함 프로젝트에서 통신장비 요구사항을 시스템 아키텍처, ICD, 시험계획, 고객 인수 기준으로 전환
'@ @'
• FFX Batch-II, KSS-III, 인도네시아 잠수함 프로젝트에서 고객이 명확히 말하지 못한 운용·설치·검수 요구사항까지 끌어내 통신장비 요구사항을 시스템 아키텍처, ICD, 시험계획, 고객 인수 기준으로 전환
'@

Add-Replacement @'
• LEO 단말/ESA 안테나에서 field issue, defect log, SW log, 시험데이터, 환경조건을 통합 분석하여 원인과 개선 우선순위 정의
'@ @'
• LEO 단말/ESA 안테나에서 field issue, defect log, SW log, 시험데이터, 환경조건을 통합 분석하여 장애 원인을 규명하고 개선 우선순위와 재발 방지 구조 정의
'@

Add-Replacement @'
• 대양전기공업 재직 기간 중 영업본부와 협업하여 신규 특수선 영업이 발생할 때마다 고객 요구사항, 함정 운용조건, 장비 구성, 개발 범위, 시험·인수 조건을 검토하고 기술제안 및 사업성 판단 자료로 정리
'@ @'
• 대양전기공업 재직 기간 중 영업본부와 협업하여 신규 특수선 영업이 발생할 때마다 고객 요구사항, 함정 운용조건, 장비 구성, 개발 범위, 시험·인수 조건을 끌어내고 복잡한 기술 이슈를 기술제안 및 사업성 판단 자료로 정리
'@

Add-Replacement @'
• 수주 이후에는 프로젝트 착수 원가와 실제 개발·제작·시험·현지지원 범위를 비교하며 원가검증 대응과 향후 유사 사업 견적 기준 개선에 기여
'@ @'
• 수주 이후에는 프로젝트 착수 원가와 실제 개발·제작·시험·현지지원 범위를 비교하며 원가검증 대응, 장애·변경 원인 분석, 향후 유사 사업 견적 기준과 재발 방지 기준 개선에 기여
'@

Add-Replacement @'
통신장비 연구소장 직무는 단순히 개발 일정을 관리하거나 기술 검토를 승인하는 자리가 아니라고 생각합니다. 고객 요구를 제품 사양으로 정리하고, 제품 사양을 설계와 검증 기준으로 바꾸며, 개발 과정에서 발생하는 문제를 품질·양산·고객 신뢰까지 이어지는 실행 체계로 관리해야 하는 자리입니다. 특히 무선통신 장비는 RF, 안테나, 디지털 하드웨어, 임베디드 제어, 네트워크, 기구, 품질, 생산이 서로 맞물려야 완성도가 나옵니다. 어느 한 부서의 판단만으로는 제품 경쟁력을 만들기 어렵고, 연구소장이 기술의 우선순위와 조직의 실행 기준을 분명하게 잡아야 합니다.
'@ @'
통신장비 연구소장 직무는 단순히 개발 일정을 관리하거나 기술 검토를 승인하는 자리가 아니라고 생각합니다. 고객 요구를 제품 사양으로 정리하고, 제품 사양을 설계와 검증 기준으로 바꾸며, 개발 과정에서 발생하는 문제를 품질·양산·고객 신뢰와 사업기회까지 이어지는 실행 체계로 관리해야 하는 자리입니다. 특히 무선통신 장비는 RF, 안테나, 디지털 하드웨어, 임베디드 제어, 네트워크, 기구, 품질, 생산, 영업이 서로 맞물려야 완성도가 나옵니다. 저는 연구소장이 복잡한 문제를 정리하고, 고객이 말로 다 표현하지 못한 기술 요구사항을 끌어내며, 개발·품질·생산·영업·고객 사이의 중간 번역자로서 기술의 우선순위와 실행 기준을 분명하게 잡아야 한다고 생각합니다.
'@

Add-Replacement @'
대양전기공업에서는 약 15년간 국방·조선용 통신/C4I 시스템을 담당했습니다. FFX Batch-II 함내 무선통신·방송체계, KSS-III 통합통신체계, 인도네시아 잠수함 오락방송장치 등 해군 통신장비 프로젝트에서 PM/PL 역할을 수행했습니다. 요구사항을 장비 구성, CPU/I/O Board, 통신 인터페이스, ICD, 시험계획, 고객 인수 기준으로 전환했고, 직접 최대 7명과 8명 규모 TFT를 조정했습니다. 또한 영업본부와 협업하여 신규 특수선 영업이 진행될 때마다 기술대안 비교, 장비 구성안, 개발 범위, 공수·자재·외주비 기반 원가 산출, 견적 검토, 제안 발표자료를 지원했습니다. HW, SW, RF, 기구, 품질, 생산 담당자가 각자의 업무만 보는 구조에서는 일정과 품질을 맞추기 어렵습니다. 그래서 저는 프로젝트별로 책임 범위, 일정, 시험 기준, 결함 처리 기준을 명확히 하고, 고객·조선소·협력사·영업본부와의 기술회의를 통해 기술적 타당성과 사업성을 함께 정리하는 방식으로 일했습니다.
'@ @'
대양전기공업에서는 약 15년간 국방·조선용 통신/C4I 시스템을 담당했습니다. FFX Batch-II 함내 무선통신·방송체계, KSS-III 통합통신체계, 인도네시아 잠수함 오락방송장치 등 해군 통신장비 프로젝트에서 PM/PL 역할을 수행했습니다. 요구사항을 장비 구성, CPU/I/O Board, 통신 인터페이스, ICD, 시험계획, 고객 인수 기준으로 전환했고, 직접 최대 7명과 8명 규모 TFT를 조정했습니다. 또한 영업본부와 협업하여 신규 특수선 영업이 진행될 때마다 고객이 명확히 표현하지 못한 운용·설치·검수 요구까지 질문으로 끌어내고, 기술대안 비교, 장비 구성안, 개발 범위, 공수·자재·외주비 기반 원가 산출, 견적 검토, 제안 발표자료를 지원했습니다. HW, SW, RF, 기구, 품질, 생산, 영업, 고객이 각자의 언어로 문제를 보는 구조에서는 일정과 품질을 맞추기 어렵습니다. 그래서 저는 중간 번역자로서 책임 범위, 일정, 시험 기준, 결함 처리 기준을 명확히 하고, 고객·조선소·협력사·영업본부와의 기술회의를 통해 복잡한 문제를 기술적 타당성, 원가, 납기, 사업성 관점으로 함께 정리하는 방식으로 일했습니다.
'@

Add-Replacement @'
둘째, 기술 이슈를 조직의 실행 과제로 바꾸는 능력입니다. 문제가 발생했을 때 단순히 “원인을 찾아보자”로 끝내면 조직은 움직이지 않습니다. 재현 조건, 담당 부서, 조치 기한, 검증 방법, 고객 보고 기준, release 판단 기준이 있어야 합니다. 저는 field issue를 defect log, corrective action, ECO/BOM, 재검증 결과, 고객 보고자료, 양산 출하 판단 기준으로 연결해 왔습니다. 400여 항목 이상의 설계검증 체크리스트를 활용해 시험 누락을 줄였고, 반복 이슈는 제품 신뢰성 개선 과제로 전환했습니다.
'@ @'
둘째, 기술 이슈를 조직의 실행 과제로 바꾸는 능력입니다. 문제가 발생했을 때 단순히 “원인을 찾아보자”로 끝내면 조직은 움직이지 않습니다. 재현 조건, 담당 부서, 조치 기한, 검증 방법, 고객 보고 기준, release 판단 기준이 있어야 합니다. 저는 field issue를 defect log, 장애 원인 분석, corrective action, ECO/BOM, 재검증 결과, 고객 보고자료, 양산 출하 판단 기준으로 연결해 왔습니다. 400여 항목 이상의 설계검증 체크리스트를 활용해 시험 누락을 줄였고, 반복 이슈는 제품 신뢰성 개선과 재발 방지 구조로 전환했습니다.
'@

Add-Replacement @'
또한 인도네시아 잠수함 오락방송장치 프로젝트에서는 무선 네트워크 아키텍처를 제안하고 고객·조선소를 설득하여 약 12억 원 규모 사업 수주에 기여했습니다. 이 프로젝트뿐 아니라 대양전기공업 재직 중 신규 특수선 영업이 발생할 때마다 영업본부와 함께 고객 요구, 기술구성, 개발 범위, 원가, 납기, 시험·인수 조건을 검토했습니다. 착수 단계에서는 개발비와 견적을 산정하고, 수행 중에는 실제 공수·자재·외주·현지지원 범위를 확인했으며, 종료 후에는 방위사업청 원가검증 대응에도 참여했습니다. 이를 통해 R&D가 기술 구현만 담당하는 조직이 아니라, 수주·원가·품질·납기·고객 신뢰까지 함께 책임져야 한다는 점을 실무로 배웠습니다. 연구소장은 개발비와 견적, 원가검증, 협력사, 생산성, 수주 경쟁력까지 이해해야 합니다. 저는 그 연결고리를 실제 프로젝트에서 경험했습니다.
'@ @'
또한 인도네시아 잠수함 오락방송장치 프로젝트에서는 무선 네트워크 아키텍처를 제안하고 고객·조선소를 설득하여 약 12억 원 규모 사업 수주에 기여했습니다. 이 프로젝트뿐 아니라 대양전기공업 재직 중 신규 특수선 영업이 발생할 때마다 영업본부와 함께 고객 요구, 기술구성, 개발 범위, 원가, 납기, 시험·인수 조건을 검토했습니다. 착수 단계에서는 고객이 말하지 못한 현장 요구를 끌어내 개발비와 견적을 산정하고, 수행 중에는 실제 공수·자재·외주·현지지원 범위와 장애 원인을 분석했으며, 종료 후에는 방위사업청 원가검증과 재발 방지 기준 정리에 참여했습니다. 이를 통해 R&D가 기술 구현만 담당하는 조직이 아니라, 기술을 사업 기회로 바꾸고 수주·원가·품질·납기·고객 신뢰까지 함께 책임져야 한다는 점을 실무로 배웠습니다.
'@

Add-Replacement @'
연구소 운영에서 가장 먼저 정비해야 할 부분은 요구사항과 검증 기준이라고 생각합니다. 요구사항이 명확하지 않으면 설계가 흔들리고, 검증 기준이 약하면 양산과 고객 인수에서 문제가 반복됩니다. 저는 고객 요구를 시스템 아키텍처, ICD, 설계검토 항목, 시험계획, 성적서, release readiness 기준으로 연결하겠습니다. 또한 HW/SW/RF/기구/품질/생산 조직이 각자의 언어로만 일하지 않도록, 공통 이슈 대장과 기술회의 체계를 운영하겠습니다.
'@ @'
연구소 운영에서 가장 먼저 정비해야 할 부분은 요구사항과 검증 기준이라고 생각합니다. 요구사항이 명확하지 않으면 설계가 흔들리고, 검증 기준이 약하면 양산과 고객 인수에서 문제가 반복됩니다. 저는 고객 요구와 잠재 요구를 시스템 아키텍처, ICD, 설계검토 항목, 시험계획, 성적서, release readiness 기준으로 연결하겠습니다. 또한 HW/SW/RF/기구/품질/생산/영업 조직이 각자의 언어로만 일하지 않도록, 공통 이슈 대장과 기술회의 체계를 운영해 복잡한 문제를 한 장의 실행 계획으로 정리하겠습니다.
'@

function Normalize-Text([string]$Text) {
    return (($Text -replace "[\r\n\a\x07]", "") -replace "\s+", " ").Trim()
}

function Set-RangeTextKeepMark($Range, [string]$Text) {
    $r = $Range.Duplicate
    if ($r.End -gt $r.Start) {
        $r.End = $r.End - 1
    }
    $r.Text = $Text
}

function New-WordApplication {
    try {
        return (New-Object -ComObject Word.Application)
    }
    catch {
        Start-Process -FilePath "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -ArgumentList "/automation" -WindowStyle Hidden
        Start-Sleep -Seconds 5
        return (New-Object -ComObject Word.Application)
    }
}

$word = New-WordApplication
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
try { $word.EventsEnabled = $false } catch {}
try { $word.AutomationSecurity = 3 } catch {}
$doc = $null

try {
    $confirmConversions = $false
    $readOnly = $false
    $addToRecentFiles = $false
    $doc = $word.Documents.Open([ref]$OutPath, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)

    $hits = 0
    $misses = New-Object System.Collections.Generic.List[string]
    foreach ($rep in $replacements) {
        $target = Normalize-Text $rep.Old
        $found = $false
        foreach ($p in $doc.Paragraphs) {
            $current = Normalize-Text $p.Range.Text
            if ($current -eq $target) {
                Set-RangeTextKeepMark $p.Range $rep.New
                $hits++
                $found = $true
                break
            }
        }
        if (-not $found) {
            $misses.Add($target) | Out-Null
        }
    }

    $doc.Save()
    $doc.ExportAsFixedFormat($PdfPath, 17)
    Write-Host "Saved DOCX: $OutPath"
    Write-Host "Saved PDF : $PdfPath"
    Write-Host "Pages     : $($doc.ComputeStatistics(2))"
    Write-Host "Words     : $($doc.ComputeStatistics(0))"
    Write-Host "Hits      : $hits"
    Write-Host "Misses    : $($misses.Count)"
    foreach ($m in $misses) {
        Write-Host "MISS: $m"
    }
    $doc.Close([ref]$true)
    $doc = $null
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close([ref]$false) } catch {}
    }
    if ($word -ne $null -and $word.Documents.Count -eq 0) {
        $word.Quit()
    }
}
