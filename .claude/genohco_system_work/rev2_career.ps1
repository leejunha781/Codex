$ErrorActionPreference='Stop'
trap { Write-Output ("TRAP @"+$_.InvocationInfo.ScriptLineNumber+": "+$_.Exception.Message); break }
$rev = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.docx"
$revPdf = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.pdf"
$tmp = "$env:TEMP\rev2_career.docx"; $tmpPdf="$env:TEMP\rev2_career.pdf"
Copy-Item $rev $tmp -Force
$VT = [string][char]11
function TrailCtl([string]$t){$c=0;for($k=$t.Length-1;$k -ge 0;$k--){if([int][char]$t[$k] -lt 32){$c++}else{break}};return $c}

$intLines = @(
 '담당 분야  |  Eutelsat OneWeb LEO 위성통신 단말 및 Ku-band ESA(전자식 위상배열) 안테나 시스템 통합검증 · 품질 · 양산 release readiness',
 '직책  |  부장(General Manager) / Senior Engineer · SI개발 SIT팀',
 '',
 '【핵심 역할】',
 '  • RF/안테나 · DFE/ADC·DAC 인터페이스 · 제어전자부 · 전원 · 네트워크 · SW log · 환경시험 결과를 통합 분석해 제품 수준의 원인분석 체계 운영',
 '  • EIRP · G/T · NF · C/N0 · scan loss · pointing error · beam steering · calibration 등 ESA 성능지표를 수신품질 · throughput · blockage recovery · handover 이슈와 연결',
 '  • 400여 항목 설계검증 체크리스트로 재현조건 · 원인 · corrective action · ECO/BOM · 재검증 · release 판단자료를 추적관리',
 '  • Eutelsat OneWeb · Marlink/STW 등 글로벌 고객과 영어 Technical Review 수행 — 품질 이슈 · 개선현황 · 잔여 리스크 설명',
 '',
 '【대표 성과】',
 '  • 반복불량 · 현장 이슈를 신뢰성 개선 과제로 전환 → 월간 불량률 약 30% 감소, 고객 이슈 처리 리드타임 약 25% 단축',
 '  • 현장 품질 이슈 유형을 8개 → 2개 수준으로 축소, 비치명 잔여 이슈는 고객과 단계적 개선계획 합의',
 '  • Tx/Rx 경로 · BFIC · LO/clock · DFE · gain plan · noise figure · linearity · phase noise · EMI/EMC · power integrity 리스크를 HW/RF 통합 관점에서 검토',
 '',
 '【이직 사유】  LEO · ESA 상용제품의 통합검증 · 품질 안정화 경험을 바탕으로, 제노코 시스템 체계 직무에서 위성 · 방산 체계 역량을 본격 발휘하고자 함'
)
$intBold = @('담당 분야','직책','【핵심 역할】','【대표 성과】','【이직 사유】','약 30% 감소','약 25% 단축','8개 → 2개')

$genLines = @(
 '담당 분야  |  방산 정비장비(Defense Maintenance Equipment) 프로젝트 리더(PL) · 시스템 / PCB 개발',
 '직책  |  Senior System Engineer / 수석연구원 · 시스템기술연구소',
 '',
 '【핵심 역할】',
 '  • K2 전차 조준경(Sight) 정비장비 시스템 개발 프로젝트 리더 — 요구사항 정의 · 시스템 설계 · 개발 일정/품질 관리',
 '  • 무인기(UAV) 정비장비 개발 프로젝트 리더 — 시스템 개발 및 PCB 설계 · 개발 총괄',
 '  • 정비장비 시스템 아키텍처 · HW/전원·신호 인터페이스 · 회로/PCB 설계검토 및 시험 · 검증 수행',
 '  • 국책 · 방산형 과제 산출물(요구사항 · 설계 · 시험 문서) 체계화 및 검사기관 대응',
 '',
 '【대표 성과】',
 '  • K2 조준경 · UAV 정비장비 2개 과제를 PL로 수행하며 방산 정비장비(EGSE 성격) 개발 전 주기를 경험',
 '  • 제노코의 정비장비 · EGSE 개발 프로세스와 방산 산출물 기준을 직접 체득',
 '',
 '【이직 사유】  방산 정비장비 PL 경험을 LEO 위성통신 단말 · ESA 안테나 상용제품의 통합검증 · 글로벌 고객 대응으로 확장하기 위해 인텔리안테크놀로지스로 이직'
)
$genBold = @('담당 분야','직책','【핵심 역할】','【대표 성과】','【이직 사유】','K2 전차 조준경(Sight) 정비장비','무인기(UAV) 정비장비')

$dyLines = @(
 '담당 분야  |  국방 통신 · C4I 체계 PM/PL — 함정 · 잠수함 통합통신체계 개발, RF 통신장비 성능검증, 시스템 통합시험, 고객 인수 및 양산이관',
 '직책  |  R&BD연구소 · SI개발 (재직 14년 11개월)',
 '',
 '【핵심 역할】',
 '  • 직접 최대 7명(SW 3 · HW 2 · 기구 2) 및 8명 규모 TFT, RFA/RF/SW/HW/기구/품질/생산 담당자를 조율하며 일정 · 인터페이스 · 이슈 종결 주도',
 '  • KSS-III 잠수함 통합통신체계 3척, FFX Batch-II 함내 무선통신 · 방송체계 3척, 인도네시아 잠수함 오락방송장치 3척의 개발 · 통합 · 고객 인수 수행',
 '  • 고객 요구사항을 시스템 아키텍처 · HW/SW 사양 · ICD · 시험계획 · FAT/HAT/SAT · 고객 인수 기준으로 구조화',
 '  • 프로젝트별 양산이관, 외주개발 범위 정의, 공급업체 선정 및 기술 · 품질 · 납기 조정',
 '',
 '【대표 성과】',
 '  • 인도네시아 잠수함 오락방송장치에 무선 네트워크 아키텍처를 제안 · 고객사/조선소 설득 → 약 12억 원 규모 수주 기여',
 '  • KSS-III에서 백본망과 유 · 무선 Ethernet 통합 네트워크를 구축해 함내 통신 인프라의 연동성 · 확장성 · 운용 안정성 확보',
 '  • VHF/UHF/HF 통신장비의 송수신 출력 · 수신감도 · 주파수 정렬 · 링크 품질 · 채널 안정성을 계측 기반으로 검증',
 '  • FAT/HAT/SAT · 시운전 · 항해시험 · 고객 인수시험을 수행하고 시험계획서 · 절차서 · 성적서 · 기술보고서로 결함 종결과 인수완료 주도',
 '',
 '【이직 사유】  약 15년간 축적한 방산 통신 · C4I 체계 PM/PL 역량을 위성 · 항공 전장품과 고신뢰성 통신장비로 확장하기 위해 이직'
)
$dyBold = @('담당 분야','직책','【핵심 역할】','【대표 성과】','【이직 사유】','약 12억 원')

$entries = @(
 @{ m='담당 분야: Eutelsat OneWeb'; lines=$intLines; bold=$intBold },
 @{ m='담당 분야: 위성·항공 전자장비'; lines=$genLines; bold=$genBold },
 @{ m='담당 분야: 국방 통신'; lines=$dyLines; bold=$dyBold }
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
Write-Output "CAREER edits applied to rev2"
