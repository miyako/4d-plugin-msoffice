# Ownership regressions

This is an ASan/UBSan harness for the plugin's ownership of 4D strings, variables, pictures and objects.

It compiles the real sources: `4DPlugin-MSOFFICE.cpp` (unity-included by `harness.mm`), `support/4DPlugin-JSON.cpp`, `C_TEXT.cpp`, `C_BLOB.cpp` and the bundled SDK's `4DPluginAPI.c`. These run against `fake_engine.cpp`, a synthetic 4D engine installed as `gCall4D`.

- Workbooks are synthetic files created with OpenXLSX in a temporary directory.
- The fake engine also models JSON Stringify, OB Copy and COPY DOCUMENT.
- Leaks are counted by the fake engine, because LSan is unavailable on macOS.

Usage:

    tests/ownership/run.sh [SOURCE_ROOT] [expected-failure ...]

`SOURCE_ROOT` defaults to this checkout. To run the same harness against an older tree:

1. Extract that tree, for example with `git archive <sha> | tar -x -C dir`.
2. Pass that directory, followed by the names of the tests expected to fail there.

Test prefixes:

- `live_`: reachable from plugin commands.
- `unreachable_`: built helpers with no caller.
- `control_`: must pass both before and after a fix.
