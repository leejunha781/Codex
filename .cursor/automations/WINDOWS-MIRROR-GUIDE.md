# Windows 미러링 가이드 (Cursor Pro / Codex)

## ⚠️ 가장 흔한 오류
`C:\Users\namma\.cursor\automations\mirror-linkedin-run-to-codex.ps1` 가 **없다**고 나오면  
아직 **install** 을 하지 않은 상태입니다. (repo 안에만 있고 로컬 Cursor 폴더에는 없음)

---

## 1단계 — 최초 1회 설치 (필수)

```powershell
cd C:\Users\namma\Documents\Codex
git fetch origin
git checkout cursor/daily-linkedin-marine-plm-04c5
git pull
powershell -ExecutionPolicy Bypass -File .\.cursor\automations\install-linkedin-mirror.ps1
```

또는 sync 스크립트 (설치 포함):
```powershell
cd C:\Users\namma\Documents\Codex
.\.cursor\automations\sync-daily-linkedin-automation.ps1
```

설치 후 생성되는 파일:
- `C:\Users\namma\.cursor\automations\mirror-linkedin-run-to-codex.ps1`
- `C:\Users\namma\.cursor\automations\mirror-linkedin-run-to-codex.bat`
- `C:\Users\namma\.cursor\automations\codex-repo.path`  ← repo 경로 저장

---

## 2단계 — 미러링 (어디서든)

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.cursor\automations\mirror-linkedin-run-to-codex.ps1" -Latest
```

또는 더블클릭:
`C:\Users\namma\.cursor\automations\mirror-linkedin-run-to-codex.bat`

### 진단 (Verify)
```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.cursor\automations\mirror-linkedin-run-to-codex.ps1" -Verify
```

## 출력 경로

`C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
- `linkedin-post.md`
- `<topic-slug>-infographic.png`
- `ready-for-posting.json`

## 미러링이 안 될 때 (체크리스트)

### 1) 브랜치 확인
2026-07-06 아티팩트는 **`main`이 아니라** `cursor/daily-linkedin-marine-plm-04c5` 브랜치에 있습니다.
PR #16이 merge되기 전에는 `main`에서 `git pull`만 하면 runs 폴더가 비어 있습니다.

```powershell
git fetch origin
git branch -a
git checkout cursor/daily-linkedin-marine-plm-04c5
git pull
```

### 2) PNG 파일 존재 확인
```powershell
dir .cursor\automations\daily-linkedin-marine-plm-post\runs\2026-07-06\marine-plm-design-change-control-governance\
```
`*-infographic.png`가 없으면 브랜치를 다시 pull 하세요.

### 3) 진단 모드
```powershell
powershell -ExecutionPolicy Bypass -File .\.cursor\automations\mirror-linkedin-run-to-codex.ps1 -Verify
```
RepoRoot / RunsRoot / CodexRoot / SourceDir 경로를 출력합니다.

### 4) 수동 지정
```powershell
powershell -ExecutionPolicy Bypass -File .\.cursor\automations\mirror-linkedin-run-to-codex.ps1 `
  -Date 2026-07-06 `
  -TopicSlug marine-plm-design-change-control-governance `
  -RepoRoot "C:\Users\namma\Documents\Codex" `
  -CodexRoot "C:\Users\namma\Documents\Codex"
```

## 이전 버그 (수정됨)
- RepoRoot 자동 탐지: `.git` 상위 walk-up → `daily-linkedin-marine-plm-post\runs` 탐색
- 단일 PNG 파일 Count 오류 수정 (`@(...)` 배열 강제)
- `ready-for-posting.json`도 함께 복사
- `.gitignore`에서 `*-infographic.png` 예외 추가 (force-add 불필요)
