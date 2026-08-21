# Open-source supply chain and advanced verification

Use this gate for every new or upgraded third-party component.

## Intake

1. Define the missing capability and prove it is not already provided by the selected STM32Cube package or project code.
2. Identify the canonical upstream and pin a release tag plus full commit. Never use an unpinned archive or default branch in a production build.
3. Inventory selected files, generated files, submodules, vendored code, build-time tools and runtime transitive components.
4. Verify repository, directory and file-level licenses plus notices. Record the chosen branch of a dual license.
5. Review release notes, security policy/advisories, open critical issues, maintenance activity, supported toolchains and target compatibility.
6. Record the exact configuration, disabled features, memory allocation policy, callbacks, thread/ISR contract, timeout units, buffer limits and failure behavior.

## Reproducibility and SBOM

- Pin compiler, linker, generator, STM32Cube/CMSIS packs, middleware, Python/Ruby tools and container images by immutable version or digest where practical.
- Preserve `.ioc`, linker script, generated-code version, compile/link flags, resolved pack metadata, component inventory/SBOM, license notices, local patch series, ELF/map/BIN/HEX hashes and signing manifest.
- Build twice from a clean checkout and compare appropriate artifacts. Explain expected nondeterminism such as timestamps, paths, build IDs or signatures; do not normalize away unexplained differences.
- Keep private signing keys outside source control and CI artifacts. Record the signing interface and public verification material, not secret values.
- Define update monitoring, vulnerability triage, support horizon, upgrade owner and rollback/removal procedure.

## Verification ladder

1. Compile the selected configuration with the real target compiler and maximum practical warnings.
2. Run portable logic on the host with Unity plus CMock or FFF across explicit adapters.
3. Run AddressSanitizer and UndefinedBehaviorSanitizer where the host compiler supports them. Treat host word size, alignment and concurrency differences explicitly.
4. Fuzz parsers, codecs, descriptors, state machines and persistence decoders with bounded input and a saved regression corpus. Include empty, truncated, oversized, duplicate, reordered and stale inputs.
5. Use static analysis from at least two complementary engines when practical; inspect suppressions and confirm high-severity findings against compiler/manual evidence.
6. Run virtual-target tests for modeled behavior and record machine coverage. Do not infer analog, electrical, true interrupt latency, DMA/cache coherency or undocumented silicon behavior.
7. Cross-build, inspect ELF/map/sections/stack assumptions, then run on-target smoke tests with hazardous outputs disabled.
8. Run logic-analyzer/oscilloscope tests and HIL fault injection, including reset or brownout at every persistent-state transition.
9. Rerun the release subset after every accepted Cursor or Claude finding and preserve the final artifact hashes.

## Review packet

Give Codex, Cursor and Claude the same frozen packet: dependency manifest, exact revisions/licenses, configuration, relevant source/diff, primary documents, build logs, tests, fuzz corpus summary, simulator model coverage, hardware measurements and `CONFIRMED`/`INFERRED`/`UNKNOWN` assumptions.

Cursor challenges integration, generated boundaries, build flags, component resolution, concurrency and diagnostics. Claude challenges trust boundaries, state/fault recovery, unsafe assumptions, malformed-input behavior and test completeness. Codex reconciles all findings against primary evidence and owns the final disposition.
