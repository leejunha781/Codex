# Open-Source Audit Tooling

Tools augment review; none proves hardware safety or full MISRA compliance by itself. Pin a released version and configuration, preserve reports, review licenses/notices, and record suppressions with owners.

Include every adopted analyzer, harness, runtime, and library in the project dependency record and SPDX or CycloneDX SBOM when distributed or used by the release process as applicable.

- Compiler warnings and sanitizers: exact target compiler warnings plus host Clang/GCC ASan and UBSan where portable code permits.
- Cppcheck: open-source C/C++ static analysis and addons. Canonical upstream: https://github.com/cppcheck-opensource/cppcheck
- Frama-C: C source analysis and proof-oriented plugins; requires scoped properties and qualified assumptions. https://github.com/Frama-C/Frama-C-snapshot
- CBMC: bounded model checking for assertions, array bounds, pointer safety, and bounded loop/state properties. Bounds are part of the claim. https://github.com/diffblue/cbmc
- clang-tidy/Clang Static Analyzer: use the version matching the selected toolchain or a documented host-analysis build.

For MISRA or safety-standard claims, distinguish a rules-oriented review/tool report from certified process or compliance. Do not claim qualification from clean output alone.
