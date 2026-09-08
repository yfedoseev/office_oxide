#!/usr/bin/env bash
# Revert-check: prove that the tests a branch adds actually cover the
# production code it adds.
#
# A test that still passes with its own production hunk reverted is testing
# nothing. This has shipped twice: once a guard whose unit test hand-built the
# struct instead of going through the parser, and once nine of thirty-two new
# tests that passed with their production hunk reverted — one of them
# asserting `.all()` over a vector that is always empty, which is vacuously
# true under any implementation, including one that does nothing.
#
# How it works: revert this branch's changes to `src/` (keeping the new tests),
# then run the test suite. It MUST fail. If it passes, the new tests do not
# exercise the new code.
#
# Usage:  scripts/revert-check.sh [base-ref]      (default: origin/main)
set -euo pipefail

BASE="${1:-origin/main}"
cd "$(git rev-parse --show-toplevel)"

if ! git diff --quiet || ! git diff --cached --quiet; then
  echo "revert-check: working tree is dirty; commit or stash first" >&2
  exit 2
fi

MERGE_BASE=$(git merge-base "$BASE" HEAD)

# Match on the directory, not on `src/**/*.rs`. In a git pathspec `*` spans
# directory separators, so `src/**/*.rs` still requires a second `/` and never
# matched a top-level file: `src/lib.rs`, `src/ir_render.rs` and every
# `src/convert_*.rs` were silently excluded from the revert, and the gate
# reported PASS having left the largest production files in place.
#
# `--no-renames` keeps the status vocabulary to A/M/D, so a rename arrives as
# a delete plus an add and each half is handled by the rule below.
CHANGED=$(git diff --no-renames --name-status "$MERGE_BASE"..HEAD \
            -- 'src/' 'crates/*/src/' | awk '$2 ~ /\.rs$/')

if [ -z "$CHANGED" ]; then
  echo "revert-check: no production changes against $BASE — nothing to check"
  exit 0
fi

CHANGED_SRC=$(echo "$CHANGED" | cut -f2-)

echo "revert-check: reverting production changes against $MERGE_BASE:"
echo "$CHANGED_SRC" | sed 's/^/  /'

cleanup() {
  # Restores modified files, re-creates ones we deleted, and drops ones we
  # restored that HEAD does not have — every case comes back from the index.
  # shellcheck disable=SC2086
  git checkout -- $CHANGED_SRC 2>/dev/null || true
}
trap cleanup EXIT

# Reverting an *added* file means removing it; `git checkout <base> -- <path>`
# cannot, because the path does not exist at the base and git errors out. That
# aborted the whole check the first time this branch added a CLI command.
while IFS=$'\t' read -r status path; do
  [ -z "$path" ] && continue
  case "$status" in
    A) rm -f "$path" ;;
    *) git checkout "$MERGE_BASE" -- "$path" ;;
  esac
done <<< "$CHANGED"

set +e
cargo test --quiet >/tmp/revert-check.log 2>&1
STATUS=$?
set -e

if [ "$STATUS" -eq 0 ]; then
  echo
  echo "revert-check: FAIL — the suite still passes with the production changes reverted."
  echo "  The new tests do not exercise the new code. See /tmp/revert-check.log"
  exit 1
fi

FAILED=$(grep -cE '^test .* FAILED|panicked at' /tmp/revert-check.log || true)
echo
echo "revert-check: PASS — the suite fails without the production changes (${FAILED} failure lines)."
