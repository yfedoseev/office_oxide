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
CHANGED_SRC=$(git diff --name-only "$MERGE_BASE"..HEAD -- 'src/**/*.rs' 'crates/*/src/**/*.rs' || true)

if [ -z "$CHANGED_SRC" ]; then
  echo "revert-check: no production changes against $BASE — nothing to check"
  exit 0
fi

echo "revert-check: reverting production changes against $MERGE_BASE:"
echo "$CHANGED_SRC" | sed 's/^/  /'

cleanup() {
  git checkout -- $CHANGED_SRC 2>/dev/null || true
}
trap cleanup EXIT

# Restore the base version of every changed production file, leaving the
# branch's tests in place.
# shellcheck disable=SC2086
git checkout "$MERGE_BASE" -- $CHANGED_SRC

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
