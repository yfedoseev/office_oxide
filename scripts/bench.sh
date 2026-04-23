#!/usr/bin/env bash
# scripts/bench.sh — single reproducible entry point for BENCHMARKS.md
#
# Usage:
#   scripts/bench.sh CORPUS_DIR [OUTPUT_DIR]
#
# CORPUS_DIR: path containing {docx,xlsx,pptx,doc,xls,ppt} subdirectories
#             (see BENCHMARKS.md for the canonical corpus sources).
# OUTPUT_DIR: where to write results + environment manifest
#             (default: ./bench-results-$(date +%Y%m%d-%H%M%S))
#
# What it does:
#   1. Captures machine spec, rustc / python / apt versions to machine.json
#   2. Installs pinned Python competitor libs (scripts/bench-requirements.txt)
#      into an isolated venv
#   3. Verifies apt-side native competitors (catdoc, antiword, xls2csv)
#   4. Builds office_oxide in release mode with LTO
#   5. Runs the Rust bench (office_oxide + calamine + dotext + docx-rs)
#   6. Runs the Python bench (markitdown, python-docx, openpyxl,
#      python-pptx, python-calamine, xlrd)
#   7. Emits a consolidated results.json
#
# NOTE: this script does not modify BENCHMARKS.md. Update tables by hand
# after reviewing results.json.

set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 CORPUS_DIR [OUTPUT_DIR]" >&2
  exit 1
fi

CORPUS_DIR="$(cd "$1" && pwd)"
OUTPUT_DIR="${2:-$(pwd)/bench-results-$(date +%Y%m%d-%H%M%S)}"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

mkdir -p "$OUTPUT_DIR"
echo "Corpus: $CORPUS_DIR"
echo "Output: $OUTPUT_DIR"

# ── 1. Machine spec ─────────────────────────────────────────────────────
echo "==> capturing machine spec"
{
  echo "{"
  printf '  "timestamp_utc": "%s",\n' "$(date -u +%Y-%m-%dT%H:%M:%SZ)"
  printf '  "uname": "%s",\n' "$(uname -srmo | sed 's/"/\\"/g')"
  printf '  "kernel": "%s",\n' "$(uname -r)"
  printf '  "cpu_model": "%s",\n' \
    "$(grep -m1 '^model name' /proc/cpuinfo 2>/dev/null \
        | sed 's/.*: //; s/"/\\"/g' \
        || echo unknown)"
  printf '  "cpu_count": %s,\n' "$(nproc 2>/dev/null || echo 0)"
  printf '  "mem_total_kb": %s,\n' \
    "$(grep -m1 '^MemTotal:' /proc/meminfo 2>/dev/null | awk '{print $2}' \
        || echo 0)"
  printf '  "rustc": "%s",\n' "$(rustc --version 2>/dev/null || echo absent)"
  printf '  "cargo": "%s",\n' "$(cargo --version 2>/dev/null || echo absent)"
  printf '  "python": "%s",\n' \
    "$(python3 --version 2>/dev/null || echo absent)"
  printf '  "catdoc": "%s",\n' \
    "$(catdoc -V 2>&1 | head -1 | sed 's/"/\\"/g' || echo absent)"
  printf '  "antiword": "%s",\n' \
    "$(antiword 2>&1 | head -1 | sed 's/"/\\"/g' || echo absent)"
  printf '  "xls2csv": "%s"\n' \
    "$(xls2csv -V 2>&1 | head -1 | sed 's/"/\\"/g' || echo absent)"
  echo "}"
} > "$OUTPUT_DIR/machine.json"

# ── 2. Python venv + pinned competitors ─────────────────────────────────
echo "==> creating bench venv"
python3 -m venv "$OUTPUT_DIR/.venv"
# shellcheck disable=SC1091
source "$OUTPUT_DIR/.venv/bin/activate"
pip install --quiet --upgrade pip
pip install --quiet -r "$REPO_ROOT/scripts/bench-requirements.txt"
pip freeze > "$OUTPUT_DIR/python-libs.txt"

# ── 3. Native competitor sanity check ───────────────────────────────────
echo "==> checking native competitors"
for tool in catdoc antiword xls2csv; do
  if ! command -v "$tool" >/dev/null 2>&1; then
    echo "  WARN: $tool not installed; legacy tables will be incomplete" >&2
  fi
done

# ── 4. Build office_oxide (release + LTO) ───────────────────────────────
echo "==> building office_oxide release binary"
cargo build --release --manifest-path "$REPO_ROOT/Cargo.toml" \
  -p office_oxide_cli

# ── 5. Run Rust bench harness ───────────────────────────────────────────
echo "==> running bench_rust (office_oxide + calamine + docx-rs + dotext)"
(
  cd "$REPO_ROOT/bench_rust"
  cargo run --release --bin bench_rust -- \
    --json "$OUTPUT_DIR/rust.json" \
    --lib all "$CORPUS_DIR" \
    > "$OUTPUT_DIR/rust.txt" 2> "$OUTPUT_DIR/rust.err" || {
      echo "bench_rust exited non-zero; see $OUTPUT_DIR/rust.err" >&2
    }
)

# ── 6. Run Python bench harness ─────────────────────────────────────────
echo "==> running bench_python"
python3 "$REPO_ROOT/bench_python.py" "$CORPUS_DIR" \
  --json "$OUTPUT_DIR/python.json" \
  > "$OUTPUT_DIR/python.txt" 2> "$OUTPUT_DIR/python.err" || {
    echo "bench_python exited non-zero; see $OUTPUT_DIR/python.err" >&2
  }

# ── 7. Summary ──────────────────────────────────────────────────────────
echo
echo "==> done"
echo "Results in: $OUTPUT_DIR"
ls -1 "$OUTPUT_DIR"
