#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_FILE="${SOURCE_FILE:-$ROOT_DIR/Custom-Excel-DNA-UDFs.cs}"
ARCH="${ARCH:-x64}"
EXCELDNA_VERSION="${EXCELDNA_VERSION:-1.8.0}"
PACK_VERSION="${PACK_VERSION:-1.5.1}"
INTEROP_VERSION="${INTEROP_VERSION:-15.0.4795.1001}"
PACKAGE_DIR="${PACKAGE_DIR:-$ROOT_DIR/.packages}"
BUILD_DIR="${BUILD_DIR:-$ROOT_DIR/.build}"
DIST_DIR="${DIST_DIR:-$ROOT_DIR/dist}"
ADDIN_NAME="CustomExcelDnaUdfs"

case "$ARCH" in
  x64) XLL_TEMPLATE="ExcelDna64.xll" ;;
  x86) XLL_TEMPLATE="ExcelDna.xll" ;;
  *) echo "ARCH must be x64 or x86." >&2; exit 2 ;;
esac

for command_name in curl python3 mcs mono; do
  command -v "$command_name" >/dev/null || { echo "$command_name is required." >&2; exit 1; }
done
[[ -f "$SOURCE_FILE" ]] || { echo "Source file not found: $SOURCE_FILE" >&2; exit 1; }

mkdir -p "$PACKAGE_DIR" "$BUILD_DIR" "$DIST_DIR"
ADDIN_PACKAGE="$PACKAGE_DIR/exceldna.addin.$EXCELDNA_VERSION.nupkg"
PACK_PACKAGE="$PACKAGE_DIR/exceldnapack.$PACK_VERSION.nupkg"
INTEROP_PACKAGE="$PACKAGE_DIR/microsoft.office.interop.excel.$INTEROP_VERSION.nupkg"

fetch() {
  local url="$1" destination="$2"
  if [[ ! -s "$destination" ]]; then
    echo "Downloading $(basename "$destination")"
    curl -fL --retry 3 --retry-delay 2 "$url" -o "$destination"
  fi
}

fetch "https://api.nuget.org/v3-flatcontainer/exceldna.addin/$EXCELDNA_VERSION/exceldna.addin.$EXCELDNA_VERSION.nupkg" "$ADDIN_PACKAGE"
fetch "https://api.nuget.org/v3-flatcontainer/exceldnapack/$PACK_VERSION/exceldnapack.$PACK_VERSION.nupkg" "$PACK_PACKAGE"
fetch "https://api.nuget.org/v3-flatcontainer/microsoft.office.interop.excel/$INTEROP_VERSION/microsoft.office.interop.excel.$INTEROP_VERSION.nupkg" "$INTEROP_PACKAGE"

python3 - "$ADDIN_PACKAGE" "$PACK_PACKAGE" "$INTEROP_PACKAGE" "$PACKAGE_DIR" <<'PY'
import sys, zipfile
from pathlib import Path
addin, pack, interop, destination = map(Path, sys.argv[1:])
destination.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(addin) as z:
    for member in [
        "tools/net452/ExcelDna.Integration.dll",
        "tools/net452/ExcelDna.xll",
        "tools/net452/ExcelDna64.xll",
    ]:
        target = destination / Path(member).name
        target.write_bytes(z.read(member))
with zipfile.ZipFile(pack) as z:
    for member in ["tools/ExcelDnaPack.exe", "tools/ExcelDnaPack.exe.config"]:
        (destination / Path(member).name).write_bytes(z.read(member))
with zipfile.ZipFile(interop) as z:
    member = "lib/net20/Microsoft.Office.Interop.Excel.dll"
    (destination / Path(member).name).write_bytes(z.read(member))
PY

rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR" "$DIST_DIR"

mcs -target:library -optimize+ -warn:4 \
  -out:"$BUILD_DIR/$ADDIN_NAME.dll" \
  -r:"$PACKAGE_DIR/ExcelDna.Integration.dll" \
  -r:"$PACKAGE_DIR/Microsoft.Office.Interop.Excel.dll" \
  "$SOURCE_FILE"

cp "$PACKAGE_DIR/ExcelDna.Integration.dll" "$BUILD_DIR/"
cp "$PACKAGE_DIR/Microsoft.Office.Interop.Excel.dll" "$BUILD_DIR/"
cp "$PACKAGE_DIR/$XLL_TEMPLATE" "$BUILD_DIR/$ADDIN_NAME.xll"

cat > "$BUILD_DIR/$ADDIN_NAME.dna" <<DNA
<DnaLibrary Name="Custom Excel-DNA UDFs" RuntimeVersion="v4.0">
  <ExternalLibrary Path="$ADDIN_NAME.dll" Pack="true" LoadFromBytes="true" ExplicitExports="false" />
  <Reference Path="Microsoft.Office.Interop.Excel.dll" Pack="true" />
</DnaLibrary>
DNA

rm -rf "$DIST_DIR"
mkdir -p "$DIST_DIR"
cp "$BUILD_DIR/$ADDIN_NAME.xll" "$DIST_DIR/"
cp "$BUILD_DIR/$ADDIN_NAME.dna" "$DIST_DIR/"
cp "$BUILD_DIR/$ADDIN_NAME.dll" "$DIST_DIR/"
cp "$BUILD_DIR/ExcelDna.Integration.dll" "$DIST_DIR/"
cp "$BUILD_DIR/Microsoft.Office.Interop.Excel.dll" "$DIST_DIR/"

OUTPUT_XLL="$DIST_DIR/$ADDIN_NAME.xll"
PACKED_XLL="$DIST_DIR/$ADDIN_NAME-packed.xll"

# ExcelDnaPack uses Windows resource APIs. Run it when the script is executed
# on Windows or under Wine; otherwise leave the standard XLL deployment bundle.
if [[ "${PACK_XLL:-auto}" != "false" ]]; then
  PACK_COMMAND=()
  case "$(uname -s 2>/dev/null || true)" in
    MINGW*|MSYS*|CYGWIN*) PACK_COMMAND=(mono) ;;
    *)
      if command -v wine >/dev/null 2>&1; then
        PACK_COMMAND=(wine)
      elif [[ "${PACK_XLL:-auto}" == "true" ]]; then
        echo "PACK_XLL=true requires Windows or Wine because ExcelDnaPack uses Windows resource APIs." >&2
        exit 1
      fi
      ;;
  esac

  if [[ ${#PACK_COMMAND[@]} -gt 0 ]]; then
    "${PACK_COMMAND[@]}" "$PACKAGE_DIR/ExcelDnaPack.exe" \
      "$BUILD_DIR/$ADDIN_NAME.dna" \
      /Y /NoMultiThreading /O "$PACKED_XLL"
    [[ -s "$PACKED_XLL" ]] || { echo "Packed XLL was not created." >&2; exit 1; }
  fi
fi

[[ -s "$OUTPUT_XLL" ]] || { echo "XLL deployment bundle was not created." >&2; exit 1; }
printf 'Built Excel-DNA deployment in %s (%s)\n' "$DIST_DIR" "$ARCH"
sha256sum "$OUTPUT_XLL"
if [[ -s "$PACKED_XLL" ]]; then sha256sum "$PACKED_XLL"; fi
