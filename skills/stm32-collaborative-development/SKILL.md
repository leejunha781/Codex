---
name: stm32-collaborative-development
description: Coordinate evidence-based STM32 firmware creation, modification, reconstruction, audit, debugging, and release work across Codex, Claude, and Cursor. Use when an STM32 task benefits from requirements traceability, exact-device evidence, independent multi-engine review, host/target/HIL verification, open-source qualification, or durable Project Memory and Obsidian updates.
---

# STM32 Collaborative Development

Treat this skill as the orchestration layer. Use the installed Taeha STM32 domain skills for detailed implementation, audit, and test rules; never silently replace them with generic STM32 knowledge.

## Canonical STM32 extensions

Treat the following canonical personal skills as named extensions of this STM32 skill and load the smallest applicable set whenever this skill is invoked:

- `$embedded-code-audit`: firmware audit and release findings.
- `$embedded-test-debug`: verification, target debugging, HIL, and qualification evidence.
- `$taeha-stm32-workflow`: end-to-end Taeha implementation and release workflow.
- `$taeha-dispense-control-simulation`: executable dispensing models, controller comparison, and evidence boundaries.

Combine their rules into one task contract; do not copy, rename, or maintain alternate definitions under another skill root. The canonical skill name is the frontmatter `name` above, and every cross-skill call must use that name.

## Boot

1. Read the applicable repository instructions and the relevant durable-memory index.
2. Load the smallest matching domain skill:
   - `stm32-device-driver` for BSP, boot, peripheral, external-device, motor, GUI, ADC, or reconstruction work.
   - `taeha-stm32-workflow` for end-to-end implementation or release delivery.
   - `embedded-code-audit` for source, concurrency, bounds, protocol, startup, or safe-state review.
   - `embedded-test-debug` for test planning, debugging, target evidence, or release qualification.
   - `stm32-dispense-mass-control` for pumps, valves, augers, syringes, material profiles, load-cell feedback, shot-weight correction, density/temperature compensation, two-component ratio, or calibration persistence.
   - `taeha-dispense-control-simulation` for Python/OpenModelica/Renode/Simulink selection, plant identification, delayed-sensor models, control-theory comparison, Monte Carlo, SIL/PIL/HIL, or learning materials.
3. Resolve all Taeha STM32 capabilities through the canonical personal skills under the user skill directory. If a client cannot resolve one, read that canonical `SKILL.md` directly and report the loader limitation; do not reactivate a duplicate plugin merely to obtain the same capability.
4. Identify the exact MCU/board revision, schematic, reference manual, datasheet, errata, toolchain, HAL/LL/RTOS versions, build configuration, and acceptance criteria. Mark absent facts `UNKNOWN`.
5. Read `references/multi-engine-review.md` for substantive or safety-relevant work. Read `references/open-source-qualification.md` before adopting or upgrading third-party code.

## Task contract

Define the goal, preserved baselines, authority boundaries, evidence requirements, safe state, tests, and exit criteria before editing. Separate `CONFIRMED`, `INFERRED`, `UNKNOWN`, and `UNVERIFIED` claims. Never invent pins, clocks, polarity, electrical levels, motor interfaces, limits, memory maps, or boot security policy.

## Execution loop

1. Convert requirements into traceable behaviors and tests.
2. Keep portable state machines, parsers, codecs, and control policy separate from STM32 HAL/LL and board adapters.
3. Define ownership, timing, ISR/task interaction, DMA/cache behavior, buffer lifetime, overflow, timeout, fault latching, recovery, and reset-safe outputs before implementation.
4. Make the smallest correct change. Preserve CubeMX user regions and existing user work.
5. For each created or modified C/C++ logical block, add one nearby dated English marker in the form `//YYYYMMDD Concise English change description`. Do not relabel unchanged code or edit vendor code only to add a marker.
6. Run the strongest available evidence ladder: warnings/static analysis, host tests, cross-build/link-map review, on-target smoke test with outputs disabled, instrumented peripheral tests, then HIL and fault injection.
7. Review the diff, safe-state behavior, recovery paths, and evidence gaps before delivery.
8. For dispensing work, treat motor movement, delivered volume, delivered mass, and A/B ratio as separate claims. Route the control design through `$stm32-dispense-mass-control`; route its audit and verification through `$embedded-code-audit` and `$embedded-test-debug`.

## Cross-engine roles

- **Codex:** own scoped implementation, local commands, builds/tests, evidence reconciliation, and final delivery.
- **Cursor:** independently challenge repository integration, build configuration, CubeMX/HAL boundaries, compiler/linker behavior, concurrency, and maintainability.
- **Claude:** independently challenge requirements interpretation, state machines, edge cases, safe startup/shutdown, fault recovery, motor/ADC safety, and missing tests.

Use the same frozen, read-only review packet for Cursor and Claude. Keep their first passes independent. Other engines advise; the executing agent verifies and reconciles findings against primary documents, code, builds, tests, and measurements. Report actual participation and any unavailable connector or authentication failure.
Before transmitting any source, diff, or proprietary artifact, obtain explicit approval for that task and frozen packet. A plugin/app mention is not standing authorization for future packets.

## Evidence and release gates

- Prefer exact ST/vendor primary documents and measured evidence over memory or broad examples.
- Pin every adopted dependency to an immutable tag and commit; record canonical upstream, licenses/notices, transitive dependencies, security status, configuration, local patches, SBOM impact, upgrade owner, and rollback/removal plan.
- Do not treat simulation, mocks, or host tests as proof of analog behavior, interrupt latency, DMA/cache coherency, electrical safety, or physical safe state.
- Preserve reproducible build inputs, map files, firmware/image hashes, manifests, board revision, tool versions, captures, and test results when applicable.
- Label hardware-only claims `UNVERIFIED` until measured on the exact target.

## Project Memory

When the user explicitly asks to learn, remember, register, persist, or update the Obsidian vault, follow the active Project Memory skill. Store concise durable knowledge, not source dumps or transcripts. Record repository-relative paths, intent, evidence, verification state, assumptions, residual risks, branch/revision, and actual Codex/Cursor/Claude participation. Route contradictions and semantic supersessions through the Promotion Inbox.

## Exit criteria

- Requirements map to code and tests.
- Host-test and target-build status are explicit.
- No unresolved critical safety or correctness finding is hidden.
- Hardware assumptions and unmeasured behavior are labeled.
- Third-party provenance and release artifacts are complete when applicable.
- The final report states changed files, verification performed, independent-review participation, residual risk, and any required restart for skill discovery.
