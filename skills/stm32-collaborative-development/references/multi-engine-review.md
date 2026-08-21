# Multi-engine review protocol

Use this protocol for substantive STM32 implementation, reconstruction, debugging, motor-control, ADC-calibration, GUI, or skill-policy changes. Codex remains the executor and evidence integrator; Cursor and Claude are independent reviewers, not authorities.

## Authorization

Before transmitting source, diffs, build logs, schematics, or proprietary context, show the exact minimal packet scope and obtain explicit user approval for this task. Prior app mentions, earlier approvals, or successful reviews are not standing authorization. Continue with local verification and report an incomplete cross-review if approval is declined or a reviewer is unavailable.

Default reviewers to read-only and exclude secrets, credentials, unrelated files, and artifacts outside the approved scope.

## Frozen packet

Give Cursor and Claude the same read-only packet:

- objective, requirements, and acceptance criteria;
- exact MCU, board, schematic, primary-document revisions, toolchain, HAL/LL/RTOS, and build configuration;
- repository revision, clean/dirty status, relevant files, diff, generated-code boundaries, and build/test commands;
- known assumptions, `UNKNOWN` items, safe state, fault model, and evidence already collected.

Exclude secrets, credentials, unrelated files, prior reviewer conclusions, and claimed expected findings.

## Roles

- **Codex:** frame requirements, implement the smallest correct change, run tests, reconcile findings, and own delivery.
- **Cursor:** challenge repository integration, call paths, generated-code boundaries, build configuration, HAL/LL, concurrency, maintainability, and the smallest safe fix.
- **Claude:** challenge requirements, state transitions, edge cases, safe startup/shutdown, fault recovery, architecture, and missing tests.


## Independent prompts

Ask Cursor to return severity, exact file/line or configuration location, failure scenario, evidence, smallest safe fix, challenged assumptions, and missing tests, emphasizing repository/build/HAL/linker/concurrency integration.

Ask Claude for the same fields, emphasizing requirements, state transitions, edge cases, safe startup/shutdown, fault recovery, motor/ADC safety, and test completeness.

Do not show one reviewer the other's first-pass findings. A review counts only when the engine actually received the frozen packet and returned findings.

## Reconciliation

Classify each finding as accepted, rejected, deferred, or needs target evidence. Verify accepted findings against primary documents and reproducible tests before editing. Preserve disagreements and the evidence used to resolve them. Re-run affected tests and, for material changes, request a focused second pass.


Resolve conflicts in this order: exact datasheet/reference manual/errata and schematic; compiler/linker output and executable tests; instrumented measurements; version-pinned vendor examples; reviewed open-source evidence; then model opinion. Reject unsupported consensus and accept minority findings when evidence proves them.

Track each finding as accepted, rejected, deferred, or needs target evidence, with reproduction and verification. Exit only after P0/P1 findings are resolved or explicitly blocked and affected tests are rerun.

## Availability

- On this Windows host, invoke Claude CLI through the approved escalated `claude.exe -p` command; do not bypass TLS, firewall, authentication, or sandbox policy.
- Run Cursor in read-only ask mode for the approved packet and report if its Windows sandbox is unavailable.
- If a reviewer is absent, unauthenticated, times out, or cannot access the packet, continue safely with local evidence and report that it did not participate.
- Do not delay an emergency safe-state fix solely for reviewer availability; request retrospective review afterward.
If a connector is missing or authorization fails, continue only when safe and report the missing review. Never synthesize a review in another engine's name.
