# Reconstruction delivery record

Create one traceability table with requirement, source artifact, evidence location, confidence, implementation module, host test, target/HIL test, and current result.

## Required outputs

- Immutable artifact manifest with SHA-256, origin, authorization, board/drive revision, and document revision.
- PDF-derived net/pin/electrical table with page coordinates and manual verification status.
- HEX memory/segment/vector summary and analysis project settings.
- `CONFIRMED`/`INFERRED`/`UNKNOWN` register, protocol, timing, unit, calibration, and fault matrix.
- IAR workspace/project/linker configuration, build options, generated map, firmware hash, and size/stack report.
- UART/servo contracts, safety state machine, diagnostics, tests, captures, mismatch log, and recovery procedure.

## Exit decision

Report the highest demonstrated L0-L4 equivalence level from `$stm32-device-driver`. Label all unmeasured hardware claims `UNVERIFIED`. A build, decompile, or successful flash alone does not prove motor behavior or safety equivalence.
