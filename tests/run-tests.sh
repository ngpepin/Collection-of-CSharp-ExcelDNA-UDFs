#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SOURCE_FILE="${SOURCE_FILE:-$ROOT_DIR/Custom-Excel-DNA-UDFs.cs}"
TMP_DIR="$(mktemp -d /tmp/exceldna-udf-tests.XXXXXX)"
trap 'rm -rf "$TMP_DIR"' EXIT

command -v mcs >/dev/null || { echo "mcs is required (Mono C# compiler)." >&2; exit 1; }
command -v mono >/dev/null || { echo "mono is required." >&2; exit 1; }

mcs -warn:4 \
  -out:"$TMP_DIR/AimlUdfTests.exe" \
  "$ROOT_DIR/tests/ExcelDnaStubs.cs" \
  "$SOURCE_FILE" \
  "$ROOT_DIR/tests/AimlUdfTests.cs"

mono "$TMP_DIR/AimlUdfTests.exe"
