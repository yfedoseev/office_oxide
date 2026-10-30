#!/usr/bin/env bash
# Run the numbered example catalogue for one binding and check what each prints.
#
#   scripts/run-examples.sh <rust|python|go|csharp|javascript|c> [id ...]
#
# Every example must finish within EXAMPLE_TIMEOUT seconds (default 120), exit 0,
# and print `OK <id>` as its last non-empty line.
#
# A timeout is treated as a defect, not as an ordinary failure: something is
# stuck (a deadlock, a blocked read). The example's whole process tree is killed,
# the run stops at once, and no later example runs. Other failures (non-zero
# exit, missing marker) are collected so one run lists every one of them.
#
# Run from the repository root, with the native library already built
# (`cargo build --release --lib`) for every binding except Rust.
set -uo pipefail

cd "$(dirname "$0")/.."

lang="${1:-}"
[ -n "$lang" ] || { echo "usage: $0 <rust|python|go|csharp|javascript|c> [id ...]" >&2; exit 2; }
shift

ids=(01_extract 02_create_from_markdown 03_edit 04_open_from_bytes 05_xlsx_writer
     06_pptx_writer 07_all_outputs 08_limits_strict_damage 09_concurrency
     10_detect_and_batch)
[ "$#" -gt 0 ] && ids=("$@")

limit="${EXAMPLE_TIMEOUT:-120}"

# Run "$@" in its own process group and kill the whole group when it overruns.
# `go run`, `dotnet run` and `cargo run` start the real program as a child, so
# signalling only the launcher would leave it running. Job control (`set -m`)
# makes the background job a group leader we can signal with `kill -- -PID`.
# Returns 124 when the limit was hit, otherwise the command's own status.
run_limited() {
  local flag="$1"; shift
  set -m
  "$@" &
  local pid=$!
  (
    sleep "$limit"
    : > "$flag"
    kill -TERM -- "-$pid" 2>/dev/null || kill -TERM "$pid" 2>/dev/null
    sleep 3
    kill -KILL -- "-$pid" 2>/dev/null || kill -KILL "$pid" 2>/dev/null
  ) &
  local watchdog=$!
  wait "$pid"
  local rc=$?
  kill "$watchdog" 2>/dev/null
  wait "$watchdog" 2>/dev/null
  set +m
  if [ -e "$flag" ]; then return 124; fi
  return "$rc"
}

command_for() {
  local id="$1"
  case "$lang" in
    rust)       echo "cargo run --quiet --example $id" ;;
    python)     echo "uv run python examples/python/$id/main.py" ;;
    go)         echo "bash -c 'cd examples/go/$id && go run -tags office_oxide_dev .'" ;;
    csharp)     echo "dotnet run --project examples/csharp/$id/$id.csproj" ;;
    javascript) echo "node examples/javascript/$id.mjs" ;;
    c)          echo "examples/c/build/$id" ;;
    *) echo "unknown binding: $lang" >&2; return 1 ;;
  esac
}

if [ "$lang" = c ]; then
  make -C examples/c all || { echo "C examples failed to build" >&2; exit 1; }
fi

failed=()
for id in "${ids[@]}"; do
  cmd="$(command_for "$id")" || exit 2
  out="$(mktemp)"
  echo "::group::$lang $id"
  echo "+ $cmd"
  flag="$out.timeout"
  run_limited "$flag" bash -c "$cmd" >"$out" 2>&1
  rc=$?
  cat "$out"
  echo "::endgroup::"
  last="$(grep -v '^[[:space:]]*$' "$out" | tail -n 1 | tr -d '\r')"
  rm -f "$out" "$flag"
  if [ "$rc" -eq 124 ]; then
    echo "::error::TIMEOUT: $lang $id did not finish within ${limit}s. It was killed and the run stopped here; the examples after it were not run. A hang is a defect."
    exit 124
  elif [ "$rc" -ne 0 ]; then
    echo "::error::$lang $id exited with status $rc"
    failed+=("$id (exit $rc)")
  elif [ "$last" != "OK $id" ]; then
    echo "::error::$lang $id did not end with 'OK $id' (last line: '$last')"
    failed+=("$id (no OK marker)")
  else
    echo "pass: $lang $id"
  fi
done

echo
if [ "${#failed[@]}" -gt 0 ]; then
  echo "$lang examples: ${#failed[@]} of ${#ids[@]} failed:"
  printf '  %s\n' "${failed[@]}"
  exit 1
fi
echo "$lang examples: all ${#ids[@]} passed"
