#!/bin/bash
# Verify every shipping manifest carries the same version string.
# Invoked by the release workflow (validate job) and the local pre-commit hook.

set -euo pipefail

CARGO=$(grep '^version = ' Cargo.toml | head -1 | sed 's/version = "\(.*\)"/\1/')
PY=$(grep '^version = ' pyproject.toml | head -1 | sed 's/version = "\(.*\)"/\1/')
WASM=$(node -p "require('./wasm-pkg/package.json').version")
JS=$(node -p "require('./js/package.json').version")
CSHARP=$(grep '<Version>' csharp/OfficeOxide/OfficeOxide.csproj | head -1 \
  | sed -E 's|.*<Version>(.*)</Version>.*|\1|')
GO=$(grep 'const defaultVersion' go/cmd/install/main.go \
  | sed -E 's/.*"(.*)".*/\1/')

echo "Cargo.toml             ${CARGO}"
echo "pyproject.toml         ${PY}"
echo "wasm-pkg/package.json  ${WASM}"
echo "js/package.json        ${JS}"
echo "csharp .csproj         ${CSHARP}"
echo "go/cmd/install         ${GO}"

fail=0
for v in "${PY}" "${WASM}" "${JS}" "${CSHARP}" "${GO}"; do
  if [ "${CARGO}" != "${v}" ]; then
    echo "FAIL: version drift — Cargo.toml=${CARGO} vs ${v}" >&2
    fail=1
  fi
done

if [ "${fail}" -ne 0 ]; then
  echo "Run .github/scripts/bump-version.sh <version> to realign." >&2
  exit 1
fi

echo "OK: all manifests at ${CARGO}"
