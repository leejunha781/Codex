---
name: cursor-pro-max-orchestration
description: Runs Cursor Pro Max Mode with Claude and Codex on every substantive task. Use for multi-model pipelines, Max Mode sessions, skill routing, embedded or career work, or when the user asks to deepen Cursor Pro with Claude/Codex.
---

# Cursor Pro Max ↔ Claude ↔ Codex Orchestration

## Default for Substantive Work

For any non-trivial task (multi-file, architecture, firmware, JD/resume, long RCA):

1. **Cursor Agent** — repo, skills, git, MCP, Cloud Agents  
2. **Claude (Opus/Sonnet, Max)** — deep reasoning, trade-offs, document polish  
3. **Codex / GPT-5.3 Codex** — heavy implementation, refactors, tests  

Skip paid duplicates unless a second pass clearly improves quality.

## When to Turn On Max Mode

| Situation | Max Mode |
|-----------|----------|
| Large repo / many `@` files | **On** |
| STM32/Taeha firmware design or RCA | **On** |
| Resume/JD long context | **On** |
| Single-file typo / rename | Off (Auto/Mini) |

User enables Max Mode in **Cursor → Settings → Models**. Agent reminds once per session if a hard task starts without Max.

## Skill Router (Load Before Acting)

| User intent | Load skills (in order) |
|-------------|------------------------|
| STM32 / driver / HAL / DMA | `embedded-codex-workflow` → `stm32-device-drivers` |
| Taeha / dispense / recipe / Modbus | `embedded-codex-workflow` → `taeha-dispense-controller` (+ STM32 if code) |
| JD / resume / cover | `30-resume-jd-workflow` → `resume-jd-match` |
| Interview / STAR | `career-interview-prep` |
| PLM / marine / defence | `plm-systems-consulting` |
| Any hard multi-step | this skill + domain skill |

Always boot `@MEMORY.md` for career/domain-accurate claims.

## Universal Pipeline

```
요구사항 정리 (Cursor)
  → 설계·트레이드오프 (Claude Max / Opus·Sonnet)
  → 구현 (Codex / GPT-5.3 Codex in Cursor Agent)
  → 검증·증거 (Cursor: build/test/measure)
  → 2차 리뷰 (Claude 또는 Bugbot)
  → 커밋·PR (Cursor)
```

### Embedded specialization

```
MCU/보드 확정 → @stm32-device-drivers (+ @taeha-dispense-controller)
  → Claude: SM·타이밍 버짓·안전
  → Codex: 드라이버/앱 코드
  → Bench/HIL 증거 → MEMORY/skill reference 학습 루프
```

See `embedded-codex-workflow` for firmware prompt template.

### Career specialization

```
JD → Cursor 분석 → Claude 문장·전략 → Cursor 파일 저장(_Targeted_*)
```

Never overwrite canonical resume/cover templates.

## Session Checklist

```
Task Progress:
- [ ] 1. Restate goal, constraints, success criteria
- [ ] 2. Route skills from table above; read SKILL.md
- [ ] 3. Choose model layer (Claude design vs Codex implement)
- [ ] 4. Max Mode if context/reasoning heavy
- [ ] 5. Execute smallest vertical slice
- [ ] 6. Verify with evidence (log/test/scope)
- [ ] 7. Persist durable facts to MEMORY or skill reference
```

## Claude Project Handoff

When using claude.ai Projects alongside Cursor:

1. Paste: goal + constraints + Cursor draft excerpt  
2. Ask for: architecture review, risk list, or English polish  
3. Bring decisions back into Cursor as concrete file edits  
4. Do not leave the only copy of design in the Claude chat

Project tip: upload `MEMORY.md` + relevant skill `SKILL.md` into **"Joonha Career & Codex"**.

## Codex Handoff Snippet

```markdown
Follow project skills:
@.cursor/skills/cursor-pro-max-orchestration/SKILL.md
@.cursor/skills/<domain-skill>/SKILL.md
@MEMORY.md

Goal: [...]
Constraints: [...]
Deliverables: [...]
Verify with: [...]
```

## Anti-Patterns

- Claiming done without build/test/measure evidence  
- Inventing Taeha-internal or MCU facts not in docs/MEMORY  
- Equating LPC1769 with STM32  
- Running Opus + Codex Max on the same trivial edit  

## Related

- [reference.md](reference.md) — model map, cost tips, skill catalog  
- `embedded-codex-workflow`, `stm32-device-drivers`, `taeha-dispense-controller`  
- Rule: `.cursor/rules/10-model-orchestration.mdc`  
- Setup: `docs/SETUP_CURSOR_PRO_CLAUDE_CODEX.md`
