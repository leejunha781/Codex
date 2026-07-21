# Visual QA Report — Navarino AM CV (2026-07-21)

## Environment
- Cloud agent: LibreOffice Writer PDF export + `pdftoppm` PNG (proxy for Word COM)
- Word COM local script prepared: `../Copy-ToLocalAndWordComQa.ps1` (requires Windows + Word + E:)

## CV
| Check | Result |
|-------|--------|
| Page count | **2** (after tighten; initial render was 3 with sparse page) |
| Title / brand | Account Manager · Maritime Connectivity… present |
| Clipped / overlap | None observed on page 1–2 PNGs |
| Sparse trailing page | Fixed |
| Genohco satcom/TVAC claim | Absent (K2 / UAV / PCB only) |
| Employed AVEVA Marine claim | Absent (self-directed disclaimer present) |
| Invented sales numbers | Absent |
| Navarino product ownership claim | Absent (“learning, not product ownership”) |
| Footer path leak on CV | Removed from outward CV |

## Strategy doc
| Check | Result |
|-------|--------|
| Page count | 2 |
| Canonical paths stated preserved | Yes |
| Claim bans section | Present |

## Local Word COM (user action on Windows PC)
Cloud cannot mount `E:\`. Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Copy-ToLocalAndWordComQa.ps1 `
  -SourceDir "<repo>\career\navarino-account-manager" `
  -DestDir "E:\이력서\Account Manager"
```

Then inspect `E:\이력서\Account Manager\qa-renders\navarino-am-20260721\`.
