#!/bin/bash
# Bump the version string across every manifest that the release workflow
# validates. Run from repo root.
#
# Usage:
#   .github/scripts/bump-version.sh 0.2.0

set -euo pipefail

if [ $# -ne 1 ]; then
  echo "usage: $0 <new-version>" >&2
  exit 1
fi
NEW="$1"

# Cargo workspace (propagates to all crates via version.workspace = true)
sed -i.bak -E "s/^(version = )\"[^\"]+\"/\1\"${NEW}\"/" Cargo.toml
# Workspace-internal dep pins
sed -i.bak -E "s/(office_oxide = \{ version = )\"[^\"]+\"/\1\"${NEW}\"/g" \
  crates/office_oxide_cli/Cargo.toml \
  crates/office_oxide_mcp/Cargo.toml

# Python
sed -i.bak -E "s/^(version = )\"[^\"]+\"/\1\"${NEW}\"/" pyproject.toml

# npm packages
node -e "
  for (const f of ['wasm-pkg/package.json', 'js/package.json']) {
    const j = JSON.parse(require('fs').readFileSync(f));
    j.version = '${NEW}';
    require('fs').writeFileSync(f, JSON.stringify(j, null, 2) + '\n');
  }
"

# C# csproj
sed -i.bak -E "s|<Version>[^<]+</Version>|<Version>${NEW}</Version>|" \
  csharp/OfficeOxide/OfficeOxide.csproj

# Go installer default
sed -i.bak -E "s/(const defaultVersion = )\"[^\"]+\"/\1\"${NEW}\"/" \
  go/cmd/install/main.go

# README snippets
sed -i.bak -E "s/(office_oxide = )\"[^\"]+\"/\1\"${NEW}\"/g" README.md docs/getting-started-rust.md

# Clean sed backups
find . -name '*.bak' -type f -delete

echo "Bumped to ${NEW}. Verify with:"
echo "  .github/scripts/check-versions.sh"
