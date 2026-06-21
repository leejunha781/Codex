$ErrorActionPreference='Stop'
trap { Write-Output ("TRAP @"+$_.InvocationInfo.ScriptLineNumber+": "+$_.Exception.Message); break }
$rev = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.docx"
$revPdf = "D:\이력서\제노코\이준하_제노코_시스템체계_rev2.pdf"
$tmp = "$env:TEMP\rev2_jaso.docx"; $tmpPdf="$env:TEMP\rev2_jaso.pdf"
Copy-Item $rev $tmp -Force
function TrailCtl([string]$t){$c=0;for($k=$t.Length-1;$k -ge 0;$k--){if([int][char]$t[$k] -lt 32){$c++}else{break}};return $c}

$ji1 = '제노코 시스템 체계 직무는 위성체 시스템 설계와 체계종합, 시스템 조립·통합 절차, 우주환경시험, 아키텍처 및 링크 버짓 분석을 하나의 실행 흐름으로 연결해야 하는 자리입니다. 저는 이 흐름을 방산과 위성통신 양쪽 현장에서 직접 수행해 온 시스템 체계 엔지니어입니다. 대양전기공업에서 약 15년간 KSS-III 잠수함과 FFX Batch-II 함정 통합통신체계의 PM/PL로 고객 요구사항을 시스템 아키텍처·ICD·FAT/HAT/SAT·고객 인수 기준으로 구조화하며 체계통합의 전 주기를 책임졌고, 인텔리안테크놀로지스에서는 Eutelsat OneWeb LEO 위성통신 단말과 Ku-band ESA 안테나의 통합검증, field issue 분석, Link/Power Budget 성능 검토, 양산 release readiness를 수행하며 상용 위성통신 제품의 시스템 검증 역량을 완성했습니다. 또한 제노코 시스템기술연구소에서 K2 전차 조준경과 무인기(UAV) 정비장비 개발 PL로 방산 정비장비의 시스템·PCB 개발과 검사기관 대응을 수행하며, 제노코의 정비장비·EGSE 개발 방식과 방산 산출물 기준을 이미 체득했습니다. 제가 제노코에 지원하는 이유는 분명합니다. 제노코의 위성탑재체·위성지상국·EGSE/정비장비·항공전자·방산 핵심부품 포트폴리오가 요구하는 시스템 체계 역량과, 제가 축적한 방산 체계통합·위성통신 단말 검증·방산 정비장비 개발 경험이 정확히 맞닿아 있기 때문입니다. 조직과 개발 방식을 이미 이해하고 있는 만큼 적응 시간을 최소화하고, 입사 초기부터 위성체 시스템 설계와 체계종합 실무에 실질적으로 기여하겠습니다.'

$ji2 = '제 경력은 회로와 보드에서 시작해 시스템 체계와 통합검증으로 확장되어 왔습니다. 초기에는 RF/MW 회로, ARM 기반 제어보드, 전원·신호 인터페이스, board bring-up, 기능시험을 수행하며 하드웨어 기본기를 쌓았습니다. 이후 대양전기공업에서 약 15년간 KSS-III 잠수함 통합통신체계, FFX Batch-II 함내 무선통신·방송체계, C4I/MOSCOS/LINK-11 연동 검증을 수행하며 고객 요구사항을 시스템 구성·HW/SW 사양·ICD·FAT/HAT/SAT·고객 인수 기준으로 구조화하고, 직접 최대 7명과 8명 규모 TFT를 이끌며 일정·인터페이스·이슈 종결을 주도했습니다. 제노코 시스템기술연구소에서는 K2 전차 조준경과 무인기(UAV) 정비장비 개발 PL로 방산 정비장비의 시스템 설계·PCB 개발·시험 및 검사기관 대응을 수행했습니다. 인텔리안테크놀로지스에서는 LEO 위성통신 단말과 Ku-band ESA 안테나에서 RF/안테나, DFE/ADC-DAC, 제어전자부, 전원, 네트워크, SW log, 환경시험 결과를 통합 분석하며 상용제품의 통합검증·품질·양산 readiness를 책임졌습니다. 이 흐름은 단순한 이직 이력이 아니라, 방산 체계통합에서 방산 정비장비 개발과 상용 위성통신 단말 검증으로 시스템 체계 역량을 단계적으로 축적해 온 과정입니다.'

$ji4 = '방산·위성통신 시스템은 설계가 완료되었다는 사실보다 검증 근거가 남아 있는지가 더 중요합니다. 대양전기공업에서는 FAT/HAT/SAT, 시운전, 항해시험, 고객 인수시험을 수행하고 시험계획서·절차서·성적서·기술보고서로 결함을 종결하며, 운용환경·보안성·인수 기준이 엄격한 방산 통신체계의 품질을 끝까지 책임졌습니다. 제노코에서는 K2 조준경과 무인기 정비장비 개발 PL로 국책·방산형 과제의 요구사항·설계·시험 산출물을 체계화하고 회로·PCB 설계검토와 검사기관 대응을 수행했습니다. 인텔리안테크놀로지스에서는 400여 항목 이상의 설계검증 체크리스트를 기반으로 field issue, defect log, SW log, 시험데이터, 네트워크 packet, EMI/EMC 및 진동 등 환경시험 결과를 연결해 release readiness와 출하 판단 근거를 관리했고, 그 결과 반복불량을 제품 신뢰성 개선 과제로 전환해 월간 불량률 약 30% 감소와 고객 이슈 처리 리드타임 약 25% 단축에 기여했습니다. 저는 시험을 단순 통과 절차가 아니라 설계 품질과 고객 신뢰를 증명하는 체계로 봅니다. 제노코에서도 시험기관·검사기관·고객 대응에 필요한 기술 근거를 명확히 정리하고, 설계 변경이 일정·품질·산출물에 미치는 영향을 선제적으로 관리하겠습니다.'

$entries = @(
 @{ m='제노코 시스템 체계 직무는 위성체 시스템 설계'; text=$ji1 },
 @{ m='제 경력은 회로와 보드에서 시작해'; text=$ji2 },
 @{ m='위성·항공·방산 시스템은 설계가 완료'; text=$ji4 }
)

$word=New-Object -ComObject Word.Application; $word.Visible=$false; $word.DisplayAlerts=0
$doc=$word.Documents.Open($tmp,$false,$false)
foreach($e in $entries){
  $done=$false
  foreach($p in $doc.Paragraphs){
    $t=($p.Range.Text -replace "[\r\n\a\x07]","").Trim()
    if($t -like ($e.m+'*')){
      $trail = TrailCtl $p.Range.Text
      $r=$p.Range.Duplicate; $r.End=$r.End-$trail; $r.Text=$e.text
      $r.Font.Bold=[int]0
      Write-Output ("OK: "+$e.m+" -> "+$e.text.Length+" chars")
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
Write-Output "JASO (#1,#2,#4) fixed in rev2"
