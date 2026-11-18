#!/usr/bin/env bash
set -euo pipefail

usage() {
    printf 'Usage: %s [--force] <workbook.xlsm> [output-directory]\n' "${0##*/}" >&2
}

force=0
while [[ $# -gt 0 ]]; do
    case "$1" in
        --force)
            force=1
            shift
            ;;
        --help|-h)
            usage
            exit 0
            ;;
        --)
            shift
            break
            ;;
        *)
            break
            ;;
    esac
done

if [[ $# -lt 1 || $# -gt 2 ]]; then
    usage
    exit 2
fi

repo_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
input_file="$1"
if [[ "$input_file" != /* ]]; then
    input_file="$PWD/$input_file"
fi
if [[ ! -f "$input_file" ]]; then
    printf 'Workbook not found: %s\n' "$input_file" >&2
    exit 1
fi

file_name="${input_file##*/}"
file_stem="${file_name%.*}"
output_dir="${2:-"$repo_root/samples/${file_stem}_extracted"}"
if [[ "$output_dir" != /* ]]; then
    output_dir="$PWD/$output_dir"
fi
if [[ -e "$output_dir" && "$force" -ne 1 ]]; then
    printf 'Output path already exists: %s\n' "$output_dir" >&2
    exit 1
fi
if [[ -e "$output_dir" && ! -d "$output_dir" ]]; then
    printf 'Output path is not a directory: %s\n' "$output_dir" >&2
    exit 1
fi

export PYTHONPATH="$repo_root/vendor/oletools${PYTHONPATH:+:$PYTHONPATH}"
"${PYTHON:-python}" - "$input_file" "$output_dir" "$force" <<'PY'
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

input_file = Path(sys.argv[1])
output_dir = Path(sys.argv[2])
force = sys.argv[3] == "1"
output_dir.mkdir(parents=True, exist_ok=force)
source_dir = output_dir / "vba_source"
source_dir.mkdir(exist_ok=force)
parser = VBA_Parser(str(input_file))
module_count = 0
seen_filenames = set()

try:
    for _, _, filename, source in parser.extract_macros():
        module_path = Path(filename)
        if module_path.name != filename or filename in ("", ".", ".."):
            raise ValueError(f"Unexpected VBA module filename: {filename!r}")
        if filename in seen_filenames:
            raise ValueError(f"Duplicate VBA module filename: {filename!r}")
        seen_filenames.add(filename)
        destination = source_dir / filename
        destination.write_text(source, encoding="utf-8", newline="")
        module_count += 1

    pcode = parser.extract_pcode()
    (output_dir / "pcode.txt").write_text(pcode or "", encoding="utf-8", newline="")
finally:
    parser.close()

print(f"Exported {module_count} VBA modules and P-code to {output_dir}")
PY