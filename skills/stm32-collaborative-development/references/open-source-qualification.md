# Open-source qualification

Qualify the exact dependency and production configuration, not a project name or reputation.

## Source hierarchy

1. Exact MCU/vendor reference manual, datasheet, errata, application note, and security bulletin.
2. Canonical upstream repository, signed release or immutable commit, release notes, license files, and security policy.
3. Reproducible local build, static analysis, host tests, simulation, target measurements, and fault-injection evidence.
4. Secondary commentary only as a discovery aid.

Do not combine evidence from different MCU families, board revisions, branches, tags, toolchains, or configurations without proving applicability.

## Adoption record

Record:

- canonical upstream URL, release/tag, immutable commit, retrieval date, and integrity/signature evidence;
- supported MCU/core/toolchain/configuration and known unsupported behavior;
- SPDX identifiers, file-level licenses/notices, transitive dependencies, and redistribution obligations;
- security policy, advisory channel, maintainer status, vulnerability review, and update owner;
- enabled features, compile options, local patches, generated configuration, and reproducible-build inputs;
- host sanitizer/fuzz/static-analysis evidence where applicable;
- target, power-loss, watchdog, malformed-input, resource-exhaustion, rollback, and recovery evidence;
- SBOM entry, field-update impact, removal plan, and residual risk.

Use `Candidate`, `Verified`, or `Superseded` as the knowledge status. `Verified` requires exact-version primary evidence plus a reproducible result relevant to the intended configuration. Hardware claims remain `UNVERIFIED` until measured.
