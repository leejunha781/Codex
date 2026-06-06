$ErrorActionPreference="Stop"

$dir="D:\이력서\AVEVA - Marine Principal Technical Support & Consultant – PLM SME, Busan\Proposal"
$out=Join-Path $dir "AVEVA_MCoE_최종PPT_예상질문_답변_가이드.docx"
$pdf=Join-Path $dir "AVEVA_MCoE_최종PPT_예상질문_답변_가이드_QA.pdf"

$word=$null;$doc=$null
function Add-P($text,$style,$before,$after,$keep){
 $p=$script:doc.Paragraphs.Add()
 $p.Range.Text=$text
 if($style -and $style -ne "Normal"){
  $styleId=switch($style){
   "Normal" {-1}
   "Heading 1" {-2}
   "Heading 2" {-3}
   "Title" {-63}
   "Subtitle" {-75}
   default {$style}
  }
  try{$p.Range.Style=[int]$styleId}catch{Write-Host ("STYLE ERROR style=<{0}> id=<{1}> type=<{2}>" -f $style,$styleId,$styleId.GetType().FullName);throw}
 }
 $p.Format.SpaceBefore=[single]$before;$p.Format.SpaceAfter=[single]$after
 if($keep){$p.Format.KeepWithNext=-1}
 return $p
}
function Add-Bullet($text){
 $p=Add-P $text "Normal" 0 4 $false
 $p.Range.ListFormat.ApplyBulletDefault()
 $p.Range.Font.Name="Aptos";$p.Range.Font.Size=[single]9.5
 return $p
}
function Add-QA($num,$q,$key,$answer,$follow){
 $p=Add-P ("Q"+$num+". "+$q) "Heading 2" 10 3 $true
 $p=Add-P ("핵심 메시지: "+$key) "Normal" 0 4 $true
 $p.Range.Font.Bold=1;$p.Range.Font.Color=[int]0x9C5700
 $p=Add-P $answer "Normal" 0 4 $false
 if($follow){$p=Add-P ("추가 질문 시: "+$follow) "Normal" 0 7 $false;$p.Range.Font.Italic=1;$p.Range.Font.Color=[int]0x666666}
}

try{
 $word=New-Object -ComObject Word.Application;$word.Visible=$false
 $doc=$word.Documents.Add()
 $sec=$doc.Sections.Item(1);$sec.PageSetup.TopMargin=$word.CentimetersToPoints(1.8);$sec.PageSetup.BottomMargin=$word.CentimetersToPoints(1.8);$sec.PageSetup.LeftMargin=$word.CentimetersToPoints(2.0);$sec.PageSetup.RightMargin=$word.CentimetersToPoints(2.0)

 $normal=$doc.Styles.Item(-1);$normal.Font.Name="Aptos";$normal.Font.Size=[single]9.5;$normal.ParagraphFormat.LineSpacingRule=0;$normal.ParagraphFormat.SpaceAfter=[single]4
 $h1=$doc.Styles.Item(-2);$h1.Font.Name="Aptos Display";$h1.Font.Size=[single]16;$h1.Font.Bold=1;$h1.Font.Color=[int]0x6B3A12;$h1.ParagraphFormat.SpaceBefore=[single]14;$h1.ParagraphFormat.SpaceAfter=[single]6;$h1.ParagraphFormat.KeepWithNext=-1
 $h2=$doc.Styles.Item(-3);$h2.Font.Name="Aptos";$h2.Font.Size=[single]11;$h2.Font.Bold=1;$h2.Font.Color=[int]0xA65B00;$h2.ParagraphFormat.KeepWithNext=-1

 $p=Add-P "AVEVA MCoE 2차 면접 후속 준비" "Title" 0 4 $true;$p.Range.Font.Name="Aptos Display";$p.Range.Font.Size=[single]22;$p.Range.Font.Bold=1;$p.Range.Font.Color=[int]0x6B3A12
 $p=Add-P "최종 제안 PPT 기반 예상 질문·답변 가이드" "Subtitle" 0 10 $false;$p.Range.Font.Name="Aptos";$p.Range.Font.Size=[single]12;$p.Range.Font.Color=[int]0x666666
 Add-P "목적" "Heading 1" 8 4 $true|Out-Null
 Add-P "김남배 전무님 또는 AVEVA MCoE 개발·기획 관계자가 최종 제안 덱을 검토한 뒤 물을 가능성이 높은 질문에 대비하기 위한 개인 답변 가이드입니다. 답변은 결론 → 근거 → HD현대중공업 적용 → 검증 방법 순서로 말합니다." "Normal" 0 6 $false|Out-Null
 Add-P "답변 원칙" "Heading 1" 8 4 $true|Out-Null
 foreach($b in @("공식 AVEVA 제품 기능과 제안 개념을 명확히 구분한다.","경쟁사를 낮추기보다 AVEVA가 연결·조율하여 더 큰 가치를 만든다고 설명한다.","AI는 승인권자가 아니라 오류·누락 탐지와 설명을 지원하는 보조자라고 강조한다.","모든 제안은 파일럿, WBS, KPI, 증거와 고객·Class 승인으로 검증한다고 답한다.")){Add-Bullet $b|Out-Null}

 Add-P "핵심 예상 질문과 답변" "Heading 1" 12 5 $true|Out-Null
 Add-QA 1 "이 자료를 추가로 만든 이유는 무엇입니까?" "면접 후에도 고객 수주와 실행 가능성을 주도적으로 고민했다." "HD현대중공업이 Siemens를 포함한 여러 대안을 검토하는 상황에서, AVEVA가 단순 기능 비교를 넘어 어떤 차별화된 실행 방향을 제시할 수 있는지 구체화하고 싶었습니다. 그래서 경쟁 비교, 개방형 API 아키텍처, 실제 프로세스 예시, WBS와 KPI까지 개발·기획팀이 논의 가능한 형태로 정리했습니다." "공식 제안서가 아니라 내부 검토와 파일럿 가설 수립을 위한 제안 개념이며, 제품팀과 고객 검증을 거쳐 구체화해야 한다고 말씀드립니다."
 Add-QA 2 "왜 AVEVA가 Siemens보다 유리하다고 보십니까?" "Siemens의 PLM 강점을 인정하면서 AVEVA의 Engineering-to-Operations 차별성을 강조한다." "Siemens는 Configuration, BOM, Change 분야에서 강력합니다. AVEVA의 차별점은 E3D/Marine, AIM, PI System, CONNECT를 통해 설계부터 운영 피드백까지 연결할 수 있다는 점입니다. 제안한 Open Control Plane은 Siemens급 형상관리 규율을 벤치마킹하면서, AVEVA의 엔지니어링·운영 문맥과 고객 소유 API를 결합하는 방향입니다." "고객이 기존 도구를 폐기하도록 요구하지 않고, 전문 도구 위에서 증거·승인·변경 영향 흐름을 조율한다고 설명합니다."
 Add-QA 3 "이 내용은 현재 AVEVA 제품 기능입니까, 아니면 신규 개발 제안입니까?" "사실·공개 신호·제안 개념의 경계를 분명히 한다." "E3D/Marine, AIM, PI System, CONNECT와 같은 제품은 현재 AVEVA 역량입니다. Open API Control Plane, Continuous Naval Assurance, 고객 구성형 프로세스, AI 게이트는 이 역량과 공개된 방향을 기반으로 한 제안 개념입니다. 따라서 실제 범위와 제품화 여부는 내부 제품팀, 고객, Class와 파일럿을 통해 확정해야 합니다." "슬라이드의 Evidence Base와 Assumption Boundary가 이 경계를 명시한다고 연결합니다."
 Add-QA 4 "왜 NAPA를 경쟁사가 아니라 연동 대상으로 보십니까?" "인증 계산의 Source of Truth는 전문 솔버에 남기고 AVEVA가 전체 흐름을 조율한다." "NAPA는 Naval Architecture와 Stability 분야의 강력한 전문 솔버입니다. 이를 대체하는 것은 위험하고 고객 신뢰도 얻기 어렵습니다. AVEVA는 NAPA와 Class 솔버 결과를 변경 이벤트, 영향 객체, 증거 팩, 승인 상태와 연결하여 설계·운영 전체 흐름을 조율할 수 있습니다." "AI는 계산 결과를 대체하지 않고 영향 후보를 우선순위화하고 설명한다고 강조합니다."
 Add-QA 5 "Open API가 보안과 거버넌스를 약화시키지 않습니까?" "Open은 무통제가 아니라 계약·권한·감사 가능한 개방성이다." "API는 OpenAPI 계약, OIDC/OAuth2, RBAC, JWT, WAF, 감사 로그와 버전 정책으로 통제합니다. 사용자가 프로세스를 구성할 수 있어도 관리자 권한과 승인된 정의 안에서만 변경하며, 모든 변경은 버전과 감사 이력을 남깁니다." "보안 게이트웨이, HA/DR, Zero Trust 접근을 파일럿 비기능 요구사항에 포함한다고 답합니다."
 Add-QA 6 "AI 게이트가 잘못 판단하면 생산이 지연되지 않습니까?" "AI는 설명·추천·누락 탐지를 담당하고 최종 승인권은 사람과 Class에 있다." "AI는 오류·누락·요구사항 미반영 후보를 탐지하고 근거를 제시합니다. 자동 승인이 아니라 위험 기반 검토를 돕는 구조이며, Release Authority는 승인자와 Class에 남습니다. AI 정밀도와 오탐률은 KPI로 측정하고, 초기에는 Conditional Pass와 Human Override를 운영합니다." "AI diagnostic precision 목표와 승인 증거 완전성을 함께 관리한다고 설명합니다."
 Add-QA 7 "Requirement, VCRM, Checklist를 왜 PLM에 포함합니까?" "생산 착수 전에 요구사항 반영 증거를 닫기 위해서다." "요구사항이 문서에만 남으면 설계 결과, 시험, 변경, 승인과 분리됩니다. Requirement 입력 시 VCRM과 Checklist를 생성하고 결과물과 검증 상태를 연결하면, Impact + Approval 단계에서 모든 요구사항이 반영되었는지 증거 기반으로 승인할 수 있습니다. 이 승인이 완료되어야 생산을 Release하도록 제안했습니다." "AI 생성 기본항목의 삭제는 관리자만 가능하고, 변경 이력을 남긴다고 답합니다."
 Add-QA 8 "E3D/Hull 데이터의 버전관리는 어떻게 접근합니까?" "E3D를 Git의 대용량 파일처럼 다루지 않고 이벤트 기반 디지털 스레드로 관리한다." "E3D는 데이터베이스 중심이므로 매 저장마다 전체 파일을 복제하는 방식은 부적합합니다. Element 변경 이력과 PLM 메타데이터·관계를 관리하고, 무거운 형상은 중복 제거 저장소에 보관합니다. Promote, AI Gate, Baseline 승인 같은 거버넌스 이벤트에서 비동기 스냅샷을 생성하는 방향이 적합합니다." "실제 E3D 데이터 모델과 성능은 도구 전문가와 파일럿 부하 시험으로 검증한다고 답합니다."
 Add-QA 9 "26주 WBS가 현실적입니까?" "완제품 개발 일정이 아니라 범위가 제한된 파일럿과 검증 증분 일정이다." "WBS는 Readiness, Phase-1 Hardening, Configuration Core, Change Impact, Naval Assurance, Pilot Adoption의 6개 패키지로 구성됩니다. 각 패키지는 API 계약, Seed Truth Set, E2E Test, KPI, Evidence와 Demo를 산출해야 다음 게이트로 진행합니다. 전체 제품화 일정이 아니라 한 개의 E3D/Hull 변경 시나리오를 닫는 파일럿 일정입니다." "리소스와 고객 가용성에 따라 재추정하며, G1에서 범위와 역할을 먼저 확정한다고 답합니다."
 Add-QA 10 "파일럿 성공 여부는 무엇으로 판단합니까?" "주관적 만족도가 아니라 정확도·증거·응답시간·승인 결과로 판단한다." "API 계약 100%, E2E 통과율 95% 이상, BOM/Effectivity 정확도 99% 이상, Impact Graph Recall 95% 이상, Solver Provenance와 Evidence Pack 100%, AI 진단 정밀도 85% 이상, Dashboard/Approval 응답 5초 이하를 기준으로 제안했습니다. 최종 Release는 고객과 Class의 승인으로 판단합니다." "KPI 기준은 시작 전 개발·기획·Naval Architecture·고객이 공동 승인해야 한다고 답합니다."
 Add-QA 11 "Principal Consultant의 역할은 무엇입니까?" "기술 설계자이면서 의사결정 구조와 고객 결과를 책임지는 통합 조정자다." "Principal Consultant는 고객 요구를 기술 계약으로 바꾸고, 개발·기획·도메인 전문가·보안·인프라·Class 사이의 결정과 의존성을 관리합니다. WBS와 KPI 게이트를 운영하고, 위험과 가정을 투명하게 관리하며, 기술 데모가 고객의 사업 결과로 연결되는지 책임집니다." "직접 모든 코드를 작성하는 역할보다 올바른 문제·범위·검증 기준을 정렬하는 역할이라고 설명합니다."
 Add-QA 12 "가장 큰 위험은 무엇입니까?" "범위 과대화, 전문 솔버/Class 수용성, 데이터 매핑, AI 과신, 운영 주체 부재다." "처음부터 전체 PLM을 만들려는 Scope Creep이 가장 큰 위험입니다. 한 개의 변경 시나리오와 제한된 어댑터로 시작하고, 전문 솔버 결과는 Source of Truth로 유지하며, Seed Truth Case로 매핑을 검증합니다. AI는 제안만 하고 승인하지 않으며, 고객·AVEVA 공동 운영위원회가 위험과 게이트를 관리합니다." "슬라이드 28의 Risk-to-Mitigation 구조로 답변을 연결합니다."
 Add-QA 13 "이 제안을 실제로 시작한다면 첫 30일에 무엇을 하겠습니까?" "고객 결정·데이터·시나리오를 고정하고 실행 가능한 기준선을 만든다." "첫째, HD현대중공업의 대표 E3D/Hull 변경 시나리오와 승인 라인을 확정합니다. 둘째, 데이터 소스·권한·API 계약·Seed Truth Set을 정의합니다. 셋째, 현재 레퍼런스 패키지를 고객 환경에서 실행하고, KPI 기준과 데모 시나리오를 공동 승인합니다. 넷째, G1 Readiness 결과를 바탕으로 파일럿 범위를 확정합니다." "첫 30일의 산출물은 합의된 시나리오, API 계약, 역할표, Seed Data, KPI Baseline이라고 답합니다."
 Add-QA 14 "자료를 만들면서 가장 중요하게 배운 점은 무엇입니까?" "완전한 단일 도구보다 전문 도구를 연결하는 고객 통제형 의사결정 흐름이 중요하다." "조선 분야는 PLM, 3D 설계, Naval Architecture, Class, 생산, 운영이 각각 전문성을 갖고 있어 한 도구가 모두를 대체하기 어렵습니다. AVEVA의 기회는 각 전문 도구의 권위를 유지하면서 변경·증거·승인·운영 피드백을 하나의 고객 통제 흐름으로 연결하는 것이라고 판단했습니다." "그래서 제안의 중심을 기능 목록이 아니라 Assurance Loop와 Control Plane으로 잡았다고 설명합니다."

 Add-P "전무님께 역으로 확인할 질문" "Heading 1" 14 5 $true|Out-Null
 foreach($b in @(
 "HD현대중공업 논의에서 현재 가장 큰 장애물은 제품 기능, 도입 위험, 가격, 조직 신뢰 중 무엇입니까?",
 "AVEVA MCoE가 우선적으로 증명하고 싶은 한 개의 고객 시나리오는 무엇입니까?",
 "본사 제품팀과 지역 MCoE 사이에서 신규 제안 기능을 파일럿으로 추진하는 절차는 어떻게 됩니까?",
 "Principal Consultant가 첫 90일 동안 달성해야 할 가장 중요한 결과는 무엇입니까?"
 )){Add-Bullet $b|Out-Null}

 Add-P "마지막 전달 문장" "Heading 1" 12 4 $true|Out-Null
 $p=Add-P "이 자료의 모든 기능이 현재 제품에 존재한다고 주장하는 것이 아닙니다. AVEVA의 실제 강점과 공개 방향을 바탕으로, HD현대중공업과 함께 검증할 수 있는 차별화된 파일럿 방향을 제시한 것입니다. 저는 이 가설을 고객·개발·기획팀과 실행 가능한 범위와 KPI로 전환하는 역할을 하고 싶습니다." "Normal" 0 8 $false
 $p.Range.Font.Bold=1;$p.Range.Font.Color=[int]0x6B3A12

 $footer=$doc.Sections.Item(1).Footers.Item(1).Range
 $footer.Text="AVEVA MCoE — Final PPT Interview Q&A Guide  |  Personal preparation"
 $footer.Font.Name="Aptos";$footer.Font.Size=[single]8;$footer.Font.Color=[int]0x777777;$footer.ParagraphFormat.Alignment=2
 $doc.SaveAs2($out,16)
 $doc.ExportAsFixedFormat($pdf,17)
 $doc.Close([ref]$false)
}finally{
 if($doc){try{$doc.Close([ref]$false)}catch{}}
 if($word){try{if($word.Documents.Count-eq0){$word.Quit()}}catch{}}
}
Write-Output("DOCX="+$out)
Write-Output("PDF="+$pdf)
