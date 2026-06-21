$ErrorActionPreference = "Stop"

$SourcePath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정.docx"
$OutPath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정_영업지원보강.docx"
$PdfPath = "D:\이력서\에이치알브로스 - 무선통신 전문기업 통신장비 연구소장\이준하_통신장비_연구소장_사람인양식_이력서_자기소개서정정_영업지원보강.pdf"

Copy-Item -LiteralPath $SourcePath -Destination $OutPath -Force

$replacements = New-Object System.Collections.Generic.List[object]
function Add-Replacement([string]$Old, [string]$New) {
    $replacements.Add([pscustomobject]@{ Old = $Old; New = $New }) | Out-Null
}

Add-Replacement @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 21년 이상 제품개발, 시스템 통합, 품질 안정화, 고객 인수, 양산이관을 수행해 온 무선통신 시스템 R&D 리더입니다. 직접 최대 7명과 8명 규모 TFT를 리딩했고, RF/HW/SW/기구/품질/생산/협력사를 연결해 고객 요구를 제품개발 우선순위와 검증근거로 전환해 왔습니다. 통신장비 연구소장으로서 개발 Gate, 품질·신뢰성 개선, 양산준비, 사업기여를 하나의 운영체계로 묶어 실행할 수 있습니다.
'@ @'
방산·해군 통신체계, 위성·항공 전자장비, LEO 위성통신 단말/ESA 안테나 분야에서 21년 이상 제품개발, 시스템 통합, 품질 안정화, 고객 인수, 양산이관을 수행해 온 무선통신 시스템 R&D 리더입니다. 직접 최대 7명과 8명 규모 TFT를 리딩했고, RF/HW/SW/기구/품질/생산/협력사를 연결해 고객 요구를 제품개발 우선순위와 검증근거로 전환해 왔습니다. 특히 대양전기공업 재직 시 영업본부와 함께 매 신규 특수선 영업 단계마다 기술구성안, 개발비·공수·자재 기반 원가 산출, 견적 및 제안자료를 지원하며 R&D 성과를 수주 경쟁력과 사업성과로 연결했습니다. 통신장비 연구소장으로서 개발 Gate, 품질·신뢰성 개선, 양산준비, 원가·수주 관점의 사업기여를 하나의 운영체계로 묶어 실행할 수 있습니다.
'@

Add-Replacement @'
R&D총괄 제품개발 기술전략 제품로드맵 무선통신장비 방산통신 위성통신 LEO통신 ESA안테나 RF/MW Digital H/W FPGA/DFE CPU Board Flash/Memory ADC/DAC Ethernet 시스템아키텍처 요구사항관리 ICD 설계검토 개발Gate FAT/HAT/SAT 고객인수 양산이관 ECO/BOM 품질개선 신뢰성개선 Field Issue Root Cause Analysis Corrective Action 협력사관리 외주개발 개발비산정 원가검증 WBS PMBOK 조직리딩 임원보고 글로벌고객대응 Technical Review
'@ @'
R&D총괄 제품개발 기술전략 제품로드맵 무선통신장비 방산통신 위성통신 LEO통신 ESA안테나 RF/MW Digital H/W FPGA/DFE CPU Board Flash/Memory ADC/DAC Ethernet 시스템아키텍처 요구사항관리 ICD 설계검토 개발Gate FAT/HAT/SAT 고객인수 양산이관 ECO/BOM 품질개선 신뢰성개선 Field Issue Root Cause Analysis Corrective Action 협력사관리 외주개발 개발비산정 원가검증 견적산정 특수선영업지원 영업본부협업 수주지원 방위사업청원가검증 WBS PMBOK 조직리딩 임원보고 글로벌고객대응 Technical Review
'@

Add-Replacement @'
• 인도네시아 잠수함 오락방송장치에 무선 네트워크 아키텍처를 제안하고 고객·조선소 설득, 원가산정, 수주까지 연결하여 12억 원 규모 사업에 기여
'@ @'
• 영업본부와 함께 매 신규 특수선 영업 단계마다 고객 요구 분석, 기술구성안 검토, 개발비·공수·자재 기반 원가 산출, 견적 및 제안자료 작성 지원을 수행
• 인도네시아 잠수함 오락방송장치에 무선 네트워크 아키텍처를 제안하고 고객·조선소 설득, 원가산정, 방위사업청 원가검증 대응, 수주까지 연결하여 약 12억 원 규모 사업에 기여
'@

Add-Replacement @'
5. 사업기여 및 대외 커뮤니케이션
'@ @'
5. 사업기여, 영업지원 및 대외 커뮤니케이션
'@

Add-Replacement @'
• 인도네시아 잠수함 오락방송장치의 무선 네트워크 아키텍처 제안, 고객·조선소 설득, 원가산정, 12억 원 규모 수주 기여
'@ @'
• 대양전기공업 재직 기간 중 영업본부와 협업하여 신규 특수선 영업이 발생할 때마다 고객 요구사항, 함정 운용조건, 장비 구성, 개발 범위, 시험·인수 조건을 검토하고 기술제안 및 사업성 판단 자료로 정리
• 개발비, 공수, 자재비, 외주·협력사 범위, 시험·현지지원 비용을 반영한 원가 산출과 견적 검토를 지원하여 영업 단계의 가격·범위·납기 리스크를 사전에 조율
• 인도네시아 잠수함 오락방송장치의 무선 네트워크 아키텍처 제안, 고객·조선소 설득, 원가산정, 방위사업청 원가검증 대응, 약 12억 원 규모 수주 기여
'@

Add-Replacement @'
• 프로젝트 착수 단계에서 영업팀과 개발비·견적 산정, 종료 후 방위사업청 원가검증 대응
'@ @'
• 수주 이후에는 프로젝트 착수 원가와 실제 개발·제작·시험·현지지원 범위를 비교하며 원가검증 대응과 향후 유사 사업 견적 기준 개선에 기여
'@

Add-Replacement @'
이번 연구소장 포지션은 경력 15년 이상, 석사 이상, 임원급 수준의 기술 의사결정과 조직 운영 능력을 동시에 요구하는 자리로 이해하고 있습니다. 저는 전자전기공학 석사 기반의 기술 전문성과 더불어, 현장에서 일정·품질·고객 요구·사업성을 함께 관리해 온 경험을 갖고 있습니다. 연구소가 좋은 기술을 보유하는 것에 그치지 않고, 그 기술이 제품 경쟁력과 고객 신뢰, 납기 준수, 수주 기여로 이어지도록 만드는 역할을 수행하고자 합니다.
'@ @'
이번 연구소장 포지션은 경력 15년 이상, 석사 이상, 임원급 수준의 기술 의사결정과 조직 운영 능력을 동시에 요구하는 자리로 이해하고 있습니다. 저는 전자전기공학 석사 기반의 기술 전문성과 더불어, 현장에서 일정·품질·고객 요구·사업성을 함께 관리해 온 경험을 갖고 있습니다. 대양전기공업에서는 연구개발 조직에 있으면서도 영업본부와 신규 특수선 영업을 함께 검토하며 원가 산출, 견적 검토, 제안 기술자료, 수주 후 원가검증까지 지원했습니다. 연구소가 좋은 기술을 보유하는 것에 그치지 않고, 그 기술이 제품 경쟁력과 고객 신뢰, 납기 준수, 원가 경쟁력, 수주 기여로 이어지도록 만드는 역할을 수행하고자 합니다.
'@

Add-Replacement @'
대양전기공업에서는 약 15년간 국방·조선용 통신/C4I 시스템을 담당했습니다. FFX Batch-II 함내 무선통신·방송체계, KSS-III 통합통신체계, 인도네시아 잠수함 오락방송장치 등 해군 통신장비 프로젝트에서 PM/PL 역할을 수행했습니다. 요구사항을 장비 구성, CPU/I/O Board, 통신 인터페이스, ICD, 시험계획, 고객 인수 기준으로 전환했고, 직접 최대 7명과 8명 규모 TFT를 조정했습니다. HW, SW, RF, 기구, 품질, 생산 담당자가 각자의 업무만 보는 구조에서는 일정과 품질을 맞추기 어렵습니다. 그래서 저는 프로젝트별로 책임 범위, 일정, 시험 기준, 결함 처리 기준을 명확히 하고, 고객·조선소·협력사와의 기술회의를 통해 문제를 끝까지 종결하는 방식으로 일했습니다.
'@ @'
대양전기공업에서는 약 15년간 국방·조선용 통신/C4I 시스템을 담당했습니다. FFX Batch-II 함내 무선통신·방송체계, KSS-III 통합통신체계, 인도네시아 잠수함 오락방송장치 등 해군 통신장비 프로젝트에서 PM/PL 역할을 수행했습니다. 요구사항을 장비 구성, CPU/I/O Board, 통신 인터페이스, ICD, 시험계획, 고객 인수 기준으로 전환했고, 직접 최대 7명과 8명 규모 TFT를 조정했습니다. 또한 영업본부와 협업하여 신규 특수선 영업이 진행될 때마다 기술대안 비교, 장비 구성안, 개발 범위, 공수·자재·외주비 기반 원가 산출, 견적 검토, 제안 발표자료를 지원했습니다. HW, SW, RF, 기구, 품질, 생산 담당자가 각자의 업무만 보는 구조에서는 일정과 품질을 맞추기 어렵습니다. 그래서 저는 프로젝트별로 책임 범위, 일정, 시험 기준, 결함 처리 기준을 명확히 하고, 고객·조선소·협력사·영업본부와의 기술회의를 통해 기술적 타당성과 사업성을 함께 정리하는 방식으로 일했습니다.
'@

Add-Replacement @'
또한 인도네시아 잠수함 오락방송장치 프로젝트에서는 무선 네트워크 아키텍처를 제안하고 고객·조선소를 설득하여 약 12억 원 규모 사업 수주에 기여했습니다. 이 프로젝트를 통해 R&D가 기술 구현만 담당하는 조직이 아니라, 수주·원가·품질·납기·고객 신뢰까지 함께 책임져야 한다는 점을 실무로 배웠습니다. 연구소장은 개발비와 견적, 원가검증, 협력사, 생산성까지 이해해야 합니다. 저는 그 연결고리를 실제 프로젝트에서 경험했습니다.
'@ @'
또한 인도네시아 잠수함 오락방송장치 프로젝트에서는 무선 네트워크 아키텍처를 제안하고 고객·조선소를 설득하여 약 12억 원 규모 사업 수주에 기여했습니다. 이 프로젝트뿐 아니라 대양전기공업 재직 중 신규 특수선 영업이 발생할 때마다 영업본부와 함께 고객 요구, 기술구성, 개발 범위, 원가, 납기, 시험·인수 조건을 검토했습니다. 착수 단계에서는 개발비와 견적을 산정하고, 수행 중에는 실제 공수·자재·외주·현지지원 범위를 확인했으며, 종료 후에는 방위사업청 원가검증 대응에도 참여했습니다. 이를 통해 R&D가 기술 구현만 담당하는 조직이 아니라, 수주·원가·품질·납기·고객 신뢰까지 함께 책임져야 한다는 점을 실무로 배웠습니다. 연구소장은 개발비와 견적, 원가검증, 협력사, 생산성, 수주 경쟁력까지 이해해야 합니다. 저는 그 연결고리를 실제 프로젝트에서 경험했습니다.
'@

Add-Replacement @'
1년 안에는 연구소가 기술·품질·사업을 함께 보는 조직으로 자리 잡게 만들겠습니다. 제품 로드맵과 개발 우선순위를 정리하고, 고객 요구와 시장 방향을 반영한 기술 과제를 선별하겠습니다. 개발 결과가 고객 인수, 양산 안정화, 원가 개선, 수주 경쟁력으로 이어지도록 사업부·품질·생산·영업과의 협업 구조를 강화하겠습니다.
'@ @'
1년 안에는 연구소가 기술·품질·사업을 함께 보는 조직으로 자리 잡게 만들겠습니다. 제품 로드맵과 개발 우선순위를 정리하고, 고객 요구와 시장 방향을 반영한 기술 과제를 선별하겠습니다. 개발 초기부터 영업·사업부와 함께 목표 원가, 개발비, 견적 기준, 양산성과 서비스성을 검토해 제품개발 결과가 고객 인수, 양산 안정화, 원가 개선, 수주 경쟁력으로 이어지도록 협업 구조를 강화하겠습니다.
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
