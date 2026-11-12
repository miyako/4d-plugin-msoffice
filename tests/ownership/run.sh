#!/bin/sh
# Build the real plugin sources + bundled SDK against the synthetic engine
# with ASan/UBSan, then run the ownership regressions.
#   tests/ownership/run.sh [SOURCE_ROOT] [expected-failure ...]
# The repo root is -iquote only: its Windows unistd.h shim must not shadow <unistd.h>.
# SOURCE_ROOT defaults to this checkout; pass another tree (e.g. a git archive
# of an older commit) to run this same harness against its sources.
set -eu
here=$(cd "$(dirname "$0")" && pwd)
root=$(cd "${1:-$here/../..}" && pwd)
[ $# -gt 0 ] && shift
out=$(mktemp -d)
trap 'rm -rf "$out"' EXIT
api="$root/4D Plugin API"
flags="-g -O1 -fno-omit-frame-pointer -fsanitize=address,undefined -fno-sanitize-recover=undefined -fno-objc-arc -w -Wno-error=missing-template-arg-list-after-template-kw"
build() {
    clang++ $flags -I"$api" -I"$api/Classes" -I"$root/support" -iquote "$root" -I"$here" \
        -I"$root/include/openxlsx" -I"$root/include" -I"$root/include/libxml2" \
        -I"$root/third_party/msoffice/include" -I"$root/third_party/cybozulib/include" \
        -std=c++17 -c "$@"
}
build -x objective-c++ "$api/4DPluginAPI.c" -o "$out/api.o"
build -x objective-c++ "$api/Classes/C_TEXT.cpp" -o "$out/ctext.o"
build -x objective-c++ "$api/Classes/C_BLOB.cpp" -o "$out/cblob.o"
build -x objective-c++ "$root/support/4DPlugin-JSON.cpp" -o "$out/json.o"
build -x c++ "$here/fake_engine.cpp" -o "$out/engine.o"
build -x objective-c++ "$here/harness.mm" -o "$out/harness.o" # includes 4DPlugin-MSOFFICE.cpp
a="$root/a"
clang++ $flags "$out"/*.o -o "$out/harness" \
    "$a/libOpenXLSX.a" "$a/libopc.a" "$a/libmce.a" "$a/libplib.a" "$a/libxml2.a" "$a/libiconv.a" \
    "$a/libjsoncpp.a" "$a/libcrypto.a" -lz \
    -framework Cocoa -framework CoreFoundation -framework Security
ASAN_OPTIONS=abort_on_error=0:halt_on_error=1 UBSAN_OPTIONS=print_stacktrace=1 "$out/harness" "$@"
