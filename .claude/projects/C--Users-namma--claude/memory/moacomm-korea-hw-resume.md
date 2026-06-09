---
name: moacomm-korea-hw-resume
description: "Korean resume for 모아컴코리아(주) HW 설계 (부장급) role; files in D:\\이력서\\모아컴코리아 - HW 개발; canonical = ..._최종_v2.docx"
metadata:
  type: project
---

The user is applying to **모아컴코리아(주) (rfmkorea.com) — HW 설계, 부장급 이상** (용인 기흥; 방산혁신기업 100, RF Filter/Transceiver/Digital HW/중계기, 수출 주력). JD core: **안테나 제어 및 신호처리 모듈의 디지털 회로 설계 및 개발관리 (FPGA, CPU, 플래시 메모리 회로 설계 포함)**; 자격: HW 설계 10년+, RF 장비 Digital H/W 설계 분석·디버깅; 우대: 5G/LTE RU, 디지털 광중계기, 영어회화. Posting appears to be via a search firm — the 양식 doc footer carries **"(주)제이앤씨에이치알" (J&C HR)** branding; keep the footer.

Working folder `D:\이력서\모아컴코리아 - HW 개발\`. Files: `이준하_모아컴 코리아_HW 개발.docx` = **양식 기준** (NOTE: its body text is the old 쏠리드 위성통신 version — layout reference only); `..._HW설계_최종.docx` = user's MoaComm draft; **`..._HW설계_최종_v2.docx` (+pdf) = my revised submission version (2026-06-10, 14 pp)**. Format rules from user: 양식 변형 금지, **맑은 고딕, 한줄 간격(lsr=0)**, formal/strong non-AI tone. Scripts/QA in `C:\Users\namma\.claude\moacom_work\` (extract/edit/render .ps1 + PNG renders).

v2 changes vs 최종 draft (12 edits + cleanup): (1) **경력 사항 엑시큐어하이트론 기간 오류 fix** "2004.01~2007.04" → **2006.01~2007.05** (now matches 상세 [기간: 2006.01~2007.05]; 2004.01 was a copy of the 네스랩 line); (2) 출장경력 #2 연도 오타 "2024.07~**2025**.08(1개월)" → 2024.08; (3) 희망연봉 "회사 보상 밴드 및 직무·역량 수준에 따라 협의" → **"회사 내규 협의"** (one line, matches 양식); (4) 고아 마침표 단락 2개 삭제 (에세텔/네스랩 추진내용 뒤); (5) 주요강점: "소규모 고기술 조직" → "소수 정예 기술조직", 대양 불릿에 **Flash/Memory**, 인텔리안 불릿에 **FPGA/DFE·ADC/DAC** 키워드 추가; (6) 인텔리안 담당분야 → "통합검증 및 Digital H/W 설계 분석·디버깅, 양산 전환 개발관리" (JD 동사 echo); (7) 인텔리안 퇴직사유 재작성 (slash-chain 제거, "퇴직을 결정하였습니다" closing); (8) 대양 담당분야에 Flash/Memory 주변회로 추가; (9) 대양 C4I 단락·개발관리 단락 마무리 문장 자연화 ("부장급 직무" 문구는 상세경력에서 제거, 자소서에는 의도적으로 1회 유지). 연봉 facts: 기본 8,397만 (총 9,484만), 변동상여 677/고정상여 377/복지포인트 50만.

**COM lesson learned:** for the LAST paragraph in a table cell, `Range.Text` shows the cell marker as TWO chars (`\r\a`) but it occupies **ONE story position** — `$r.End = $r.End - 1` is correct (same as body paragraphs); subtracting 2 leaves the last content char and produces doubled-char corruption (e.g. "협의의"). Body-paragraph End-1 trick per [[office-docs-com-automation]].

Candidate facts shared with [[itt-cannon-fae-resume]] (대양전기공업 ~15y 함정/잠수함 통신 PM/PL + CPU/IO board, 제노코, 인텔리안 OneWeb ESA SIT, 부산대 석사). User's original kept untouched (was open in user's Word during the session — edit via ReadOnly open + SaveAs2 to v2).
