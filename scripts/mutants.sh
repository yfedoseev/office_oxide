#!/usr/bin/env bash
# Mutation testing of the limits, options, package and writer code with
# cargo-mutants: each mutant is a small change to one line, and a mutant
# the tests still pass with ("missed") is code no test checks.
#
# Each file is run against the tests that exercise it, not the whole suite:
# a mutant of a writer can only be caught by the tests that write, and
# rebuilding and running every test binary for each of a thousand mutants
# takes a day.
#
# Usage: scripts/mutants.sh [limits|opc|create|all] [JOBS]
#   (from the repository root; needs `cargo install cargo-mutants`)
# Results: mutants.out-<set>/{caught,missed,timeout,unviable}.txt
set -euo pipefail

cd "$(dirname "$0")/.."

set_name=${1:-limits}
jobs=${2:-3}

# The tests that read limits and options, by binary.
limits_tests=(
  --lib
  --test test_parse_options --test test_package_limits --test test_part_size_limits
  --test test_write_limits --test test_text_budget --test test_image_budget
  --test test_xlsx_amplification --test test_ooxml_amplification
  --test test_embedded_package_nesting --test test_limits_docs
  --test test_link_target_amplification --test test_xlsx_padding_parity
)
# The tests that read packages.
opc_tests=(
  --lib
  --test test_package_limits --test test_part_size_limits --test test_ooxml_unreadable_parts
  --test test_damaged_part_salvage --test test_robustness_and_safety
  --test test_core_integration --test test_lazy_pictures
)
# The tests that write documents.
create_tests=(
  --lib
  --test test_write_integration --test test_write_path_validity --test test_write_limits
  --test test_ir_roundtrip_idempotence --test test_docx_integration --test test_office_integration
  --test test_xlsx_integration --test test_pptx_integration
)

run() {
  local name=$1; shift
  local -n tests=$1; shift
  # Debug info off: each mutant rebuilds the crate and links its tests.
  CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
    cargo mutants --jobs "$jobs" --timeout 600 --build-timeout 900 \
      "$@" --output "mutants.out-$name" -- "${tests[@]}"
}

case $set_name in
  limits) run limits limits_tests --file src/limits.rs --file src/options.rs ;;
  opc) run opc opc_tests --file src/core/opc.rs ;;
  create) run create create_tests --file src/create.rs ;;
  all)
    run limits limits_tests --file src/limits.rs --file src/options.rs
    run opc opc_tests --file src/core/opc.rs
    run create create_tests --file src/create.rs
    ;;
  *) echo "usage: $0 [limits|opc|create|all] [JOBS]" >&2; exit 2 ;;
esac
