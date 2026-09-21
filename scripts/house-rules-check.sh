#!/usr/bin/env bash
# House-rules check: mechanical rules from CONTRIBUTING.md that a review
# should never have to catch by eye. Each rule is one grep; the script fails
# if any rule finds a violation and prints every offending line as
# `path:line: text` so the fix is a click away.
#
#   1. Every Rust test function is named `test_*`. A test is the one kind of
#      function that is never called by name, so nothing else distinguishes
#      it from a helper in a diff or a grep.
#   2. No `allow(dead_code)`. Dead code is deleted, not silenced; a
#      test-only helper is `#[cfg(test)]`.
#   3. No issue/PR numbers in code or comments. Tests are named by defect
#      class and reporters are credited in CHANGELOG.md; a ticket number in a
#      comment is a pointer that rots the moment the tracker moves.
#
# Usage: scripts/house-rules-check.sh   (from the repository root)
set -euo pipefail

cd "$(dirname "$0")/.."

# Rust sources subject to the rules: the library, its tests, and the
# workspace crates. Bindings in other languages keep their own conventions.
rust_files() {
  git ls-files -- 'src/*.rs' 'tests/*.rs' 'crates/*.rs' 'fuzz/*.rs' 'examples/*.rs' 'bench_rust/*.rs'
}

status=0
# $1 = rule title, stdin = offending lines. Returns 1 when there are any, so
# a `| report ... || status=1` pipeline records the failure in the parent
# shell (an assignment inside the pipeline would be lost to the subshell).
report() {
  local title=$1 lines
  lines=$(cat)
  [[ -z "$lines" ]] && return 0
  printf '\n== %s ==\n%s\n' "$title" "$lines"
  return 1
}

# Rule 1: `#[test]` (or `#[tokio::test]`, with or without arguments), then
# any further attributes, then `fn <name>` — <name> must start with `test_`.
rust_files | xargs awk '
  /^[[:space:]]*#\[(tokio::)?test(\(.*\))?\][[:space:]]*$/ { pending = 1; next }
  pending && /^[[:space:]]*#\[/ { next }
  pending && /^[[:space:]]*(pub([[:space:]]*\([^)]*\))?[[:space:]]+)?(async[[:space:]]+)?fn[[:space:]]+/ {
    name = $0
    sub(/^[[:space:]]*(pub([[:space:]]*\([^)]*\))?[[:space:]]+)?(async[[:space:]]+)?fn[[:space:]]+/, "", name)
    sub(/[^A-Za-z0-9_].*$/, "", name)
    if (name !~ /^test_/) printf "%s:%d: fn %s\n", FILENAME, FNR, name
    pending = 0; next
  }
  { pending = 0 }
' | report "test functions must be named test_* (rule 1)" || status=1

# Rules 2 and 3 are plain greps. `|| true` inside the group: grep exits 1 on
# no match, which `pipefail` would otherwise report as a failure.
#
# Rule 2: no dead-code suppression, file-wide or targeted. The one
# structural exception is `tests/common/`: Cargo compiles it into every
# integration-test binary separately, so a builder that any one binary does
# not use warns there, and no attribute narrower than `allow` covers that.
{ rust_files | grep -v '^tests/common/' | xargs grep -HnE 'allow\(dead_code\)' || true; } \
  | report "allow(dead_code) is not permitted — delete the code or mark it #[cfg(test)] (rule 2)" || status=1

# Rule 3: no issue/PR numbers in code or comments: `#` followed by two or
# more digits, standing alone. Benign look-alikes are masked out of each
# line before matching (not used to skip the line, so a ticket number next
# to one is still caught): `&#160;` character references and `#112233`
# colours. `#,##0` number formats never put digits directly after `#`.
{
  while IFS= read -r f; do
    sed -E 's/&#[0-9]+;//g; s/#[0-9A-Fa-f]{6}\b//g' "$f" \
      | grep -nE '(^|[^A-Za-z0-9_])#[0-9]{2,}\b' | sed "s|^|$f:|" || true
  done < <(rust_files)
} | report "issue/PR numbers do not belong in code or comments — credit reporters in CHANGELOG.md (rule 3)" || status=1

if [[ $status -ne 0 ]]; then
  printf '\nhouse-rules-check: violations found (see above)\n' >&2
  exit 1
fi
printf 'house-rules-check: ok\n'
