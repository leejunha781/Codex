# Open-Source Test and Debug Tooling

Pin released versions/commits, review licenses and target support, preserve configuration and reports, and avoid claiming more than each tool proves.

Record distributed or release-process dependencies in the project dependency register and SPDX or CycloneDX SBOM as applicable, including test-only transitive components that affect reproducibility.

- Ceedling, Unity, CMock, FFF: portable C unit tests, generated mocks, and fakes. https://github.com/ThrowTheSwitch/Ceedling , https://github.com/ThrowTheSwitch/Unity , https://github.com/ThrowTheSwitch/CMock , https://github.com/meekrosoft/fff
- LLVM libFuzzer: deterministic in-process coverage fuzzing, normally with ASan/UBSan. Maintenance status and host-only scope must be recorded. https://llvm.org/docs/LibFuzzer.html
- AFL++: coverage-guided fuzzing for host harnesses; isolate parsers/calibration loaders from hardware dependencies. https://github.com/AFLplusplus/AFLplusplus
- Renode: repeatable system/peripheral simulation where the exact platform model exists. A passing simulation does not prove analog behavior, interrupt latency, DMA/cache correctness, or safe outputs. https://github.com/renode/renode
- pyOCD and OpenOCD: target flash/debug automation with exact pack/config/probe records. https://github.com/pyocd/pyOCD , https://github.com/openocd-org/openocd
- labgrid and Robot Framework: lab resource control and acceptance orchestration. Bind exact board/probe/instrument identity and collect artifacts. https://github.com/labgrid-project/labgrid , https://github.com/robotframework/robotframework

Use the real production compiler/linker and physical instruments for release evidence even when host and simulation coverage are strong.
