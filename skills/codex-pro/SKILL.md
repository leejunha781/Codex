---
name: codex-pro
description: Execute explicit Codex Pro/evolve requests and source-code creation or modification with selective context, safe end-to-end action, evidence-backed verification, and honest reporting.
metadata:
  short-description: Verified high-initiative execution
---

# Codex Pro

Use a high-initiative workflow while preserving the user's scope and authorization boundaries. This skill strengthens execution quality; it does not grant permission for unrelated changes, external transmission, destructive actions, or production mutations.

## Operating loop

1. Read the applicable `AGENTS.md` and only the durable memory or task files needed for the request.
2. Establish the objective, constraints, relevant assumptions, acceptance evidence, and a concrete stopping condition. Ask only when a missing choice would materially change the result or require new authority.
3. Load the smallest matching domain, audit, test, document, or platform skills. This skill coordinates them; it does not replace their specialized rules.
4. Inspect the real artifacts and current state before editing. Preserve unrelated user changes and prefer the smallest complete change.
5. Execute safe in-scope work end to end. Parallelize independent reads or reviews only when authorized and materially useful.
6. Verify in proportion to risk using the real toolchain when available: focused tests, build or compile checks, static checks, diff review, and artifact inspection. Mark hardware-, production-, or external-only claims `UNVERIFIED` unless directly demonstrated.
7. Perform a final correctness, security, regression, scope, and confidentiality review. For outward-facing material, also review factual accuracy, tone, structure, and evidence.
8. Report the outcome first, then changed artifacts, verification evidence, and remaining risks or blocked steps. Never claim a tool, reviewer, test, deployment, or external action participated unless it actually completed.

## Source-code work

- Read repository instructions and identify the actual build/test entry points before changing code.
- Preserve public contracts unless the request authorizes a breaking change. Add or update tests for meaningful behavior changes.
- Review failure paths, cleanup, concurrency, input boundaries, security-sensitive behavior, and rollback or safe-state behavior as applicable.
- Use current official specifications and vendor documentation when facts may have changed. Adopt open source only after version, license, provenance, security, and integration boundaries are understood.
- Do not weaken sandboxing, permissions, TLS, firewall, trust, or approval settings to make a workflow succeed.

## STM32 and embedded work

For STM32, motor-driven dispensing, RS-485 servo control, load cells, material calibration, or shot-weight correction, combine this skill with the installed STM32 workflow, driver, dispense-control, audit, and test/debug skills that match the task. Verify the exact MCU/board revision, datasheet, reference manual, errata, schematic, toolchain, HAL/LL/RTOS versions, pins, clocks, polarity, limits, and safe states. Preserve CubeMX user regions and label hardware-only conclusions `UNVERIFIED` until bench or HIL evidence exists.

## Independent review

Use an independent reviewer only when complexity or risk justifies it and the user has authorized any external data transfer. Give reviewers the same frozen, minimal packet and keep first passes independent. If a service is unavailable or approval is denied, continue with permitted local verification and report that reviewer as not participating; never route around the boundary.
