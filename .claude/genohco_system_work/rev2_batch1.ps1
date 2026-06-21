$ErrorActionPreference='Stop'
trap { Write-Output ("TRAP @"+$_.InvocationInfo.ScriptLineNumber+": "+$_.Exception.Message); break }
$rev = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.docx"
$revPdf = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.pdf"
$tmp = "$env:TEMP\rev2_b1.docx"; $tmpPdf="$env:TEMP\rev2_b1.pdf"
Copy-Item $rev $tmp -Force
$VT = [string][char]11
function TrailCtl([string]$t){$c=0;for($k=$t.Length-1;$k -ge 0;$k--){if([int][char]$t[$k] -lt 32){$c++}else{break}};return $c}

$careerDesc = @(
 '■ 프로젝트 1   Eutelsat OneWeb LEO 위성통신 단말 및 Ku-band ESA 안테나 통합검증 · 양산검증',
 '   소속/기간 · ㈜인텔리안테크놀로지스 · 판교연구소 SIT팀 · 2023.03~2025.07 · 부장(General Manager)/Senior Engineer',
 '   주요 업무 · LEO 단말·Ku-band ESA 안테나 통합검증, field issue 분석, Link/Power Budget 성능 검토, 양산 release readiness 관리',
 '   기술 스택 · Ku-band ESA · Phased Array · beam steering/calibration · EIRP/G/T/NF/C/N0 · FPGA/DFE · ADC/DAC · Ethernet · Wireshark · iperf · EMI/EMC · 진동시험',
 '   상세 · 단품 보드·개별 RF 성능 확인이 아니라, 실제 고객 환경의 수신품질 저하·throughput 변동·blockage recovery·handover·antenna pointing·네트워크 불안정을 시스템 관점에서 분석. 400여 항목 설계검증 체크리스트로 반복 검증 체계 운영 → 월간 불량률 약 30% 감소, 고객 이슈 리드타임 약 25% 단축, 현장 품질 이슈 유형 8→2개 축소.',
 '',
 '■ 프로젝트 2   KSS-III 잠수함 통합통신체계 개발 · 통합 · 고객 인수',
 '   소속/기간 · 대양전기공업㈜ · R&BD연구소 · PM/PL',
 '   주요 업무 · 잠수함 통합통신체계 시스템 아키텍처, 통합 네트워크, ICD, FAT/HAT/SAT, 항해시험, 고객 인수',
 '   상세 · 백본망과 유·무선 Ethernet을 결합한 통합 네트워크를 구축. 장비 간 ICD·RF 통신장비 성능시험·FAT/HAT/SAT·시운전·항해시험·고객 인수시험 수행. 운용환경·보안성·장비 연동·인수 기준이 엄격한 방산 통신체계에서 인터페이스·시험 기준을 명확히 잡고 결함을 종결. 함정 3척 개발·인수 완료.',
 '',
 '■ 프로젝트 3   FFX Batch-II 함내 무선통신 · 방송체계 개발 및 인수시험',
 '   소속/기간 · 대양전기공업㈜ · PM/PL',
 '   주요 업무 · 수상함 함내 무선통신·방송체계 개발, RF 성능검증, 시스템 통합, 시험계획, 고객 인수',
 '   상세 · VHF/UHF/HF 통신장비의 출력·수신감도·주파수 정렬·채널 안정성·안테나/케이블 영향도를 계측 기반으로 확인하고, 시험 결과를 고객 제출자료·인수 기준에 반영. 수상함 3척 통신체계 개발.',
 '',
 '■ 프로젝트 4   인도네시아 잠수함 오락방송장치 무선 네트워크 아키텍처 제안 · 수주 · 현지 지원',
 '   소속/기간 · 대양전기공업㈜ · 제안/수주 PL',
 '   주요 업무 · 무선 네트워크 구조 제안, 개발비·견적 산정, 고객·조선소 기술 설득, 현지 지원, 원가검증 대응',
 '   상세 · 기술대안 비교와 무선 네트워크 아키텍처 제안으로 약 12억 원 규모 사업 수주에 기여. R&D가 수주·원가·품질·납기·고객 신뢰까지 함께 책임지는 사업기여 경험을 확보.',
 '',
 '■ 프로젝트 5   C4I / MOSCOS / LINK-11 전술데이터링크 연동 검토 및 통신체계 안정화',
 '   소속/기간 · 대양전기공업㈜',
 '   주요 업무 · 장비 간 데이터 흐름, 통신 프로토콜, latency, 데이터 무결성, 네트워크 안정성, 운용 시나리오 검토',
 '   상세 · 단일 장비 중심이 아니라 상위 체계와 현장 운용 시나리오를 고려해 요구사항·ICD·시험절차·검증 결과를 일관되게 관리.',
 '',
 '■ 프로젝트 6   방산 정비장비(K2 조준경 · 무인기) 시스템 · PCB 개발',
 '   소속/기간 · ㈜제노코 · 시스템기술연구소 · Senior System Engineer / PL · 2022.08~2023.03',
 '   주요 업무 · K2 전차 조준경(Sight) 정비장비 시스템 개발 PL, 무인기(UAV) 정비장비 시스템·PCB 개발 PL',
 '   상세 · 방산 정비장비(EGSE 성격) 2개 과제를 PL로 수행하며 요구사항 정의·시스템 설계·PCB 설계검토·시험·검사기관 대응을 수행. 제노코의 정비장비·EGSE 개발 프로세스와 방산 산출물 기준을 직접 체득.'
)
$careerBold = @('■ 프로젝트 1','■ 프로젝트 2','■ 프로젝트 3','■ 프로젝트 4','■ 프로젝트 5','■ 프로젝트 6','소속/기간','주요 업무','기술 스택','약 12억 원','약 30% 감소','약 25% 단축')

$skillLines = @(
 '• 체계종합/통합검증 : 인텔리안 LEO·ESA 단말, 대양전기공업 함정·잠수함 통신체계, 제노코 방산 정비장비에서 조립·통합 절차, FAT/HAT/SAT, field issue 분석, release readiness 관리 경험.',
 '• 환경시험/품질 : EMI/EMC·진동·방수 등 환경시험 대응, 우주환경시험(열진공 등) 요구사항 검토, 시험기관·검사기관 질의 및 산출물 대응.'
)
$skillBold = @('체계종합/통합검증','환경시험/품질')

$entries = @(
 @{ m='1) 프로젝트명: Eutelsat OneWeb'; lines=$careerDesc; bold=$careerBold },
 @{ m='• 체계종합/통합검증'; lines=$skillLines; bold=$skillBold }
)

$word=New-Object -ComObject Word.Application; $word.Visible=$false; $word.DisplayAlerts=0
$doc=$word.Documents.Open($tmp,$false,$false)
foreach($e in $entries){
  $done=$false
  foreach($p in $doc.Paragraphs){
    $t=($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
    if($t -like ($e.m+'*')){
      $newText = ($e.lines -join $VT)
      $trail = TrailCtl $p.Range.Text
      $r=$p.Range.Duplicate; $r.End=$r.End-$trail; $r.Text=$newText
      $r.Font.Bold=[int]0
      $base=$p.Range.Start; $full=$p.Range.Text
      foreach($bl in $e.bold){
        $start=0
        while(($idx=$full.IndexOf($bl,$start)) -ge 0){
          $doc.Range($base+$idx,$base+$idx+$bl.Length).Font.Bold=[int]1
          $start=$idx+$bl.Length
        }
      }
      Write-Output ("OK: "+$e.m+" -> "+$newText.Length+" chars")
      $done=$true; break
    }
  }
  if(-not $done){ Write-Output ("NOT FOUND: "+$e.m) }
}
$doc.Save()
$doc.ExportAsFixedFormat($tmpPdf,17)
$doc.Close([ref]$true)
$word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
Copy-Item $tmp $rev -Force; Copy-Item $tmpPdf $revPdf -Force
Write-Output "BATCH1 (경력기술서 + 스킬) applied to rev2"
