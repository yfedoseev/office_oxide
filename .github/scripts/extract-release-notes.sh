#!/bin/bash
# Extracts release notes for a given version from CHANGELOG.md
# Usage: extract-release-notes.sh <version>
# Outputs:
#   release-title.txt  — "v0.3.5 | Performance, ..."
#   release-notes.md   — Full body (changelog section + installation footer)

set -euo pipefail

VERSION="$1"
CHANGELOG="CHANGELOG.md"

if [ ! -f "$CHANGELOG" ]; then
  echo "Error: $CHANGELOG not found" >&2
  exit 1
fi

# Extract subtitle from "> ..." line after version header
SUBTITLE=$(awk "/^## \[${VERSION}\]/{found=1; next} found && /^>/{gsub(/^> */, \"\"); print; exit}" "$CHANGELOG")

# Build title
if [ -n "$SUBTITLE" ]; then
  echo "v${VERSION} | ${SUBTITLE}" > release-title.txt
else
  echo "v${VERSION}" > release-title.txt
fi

# Extract body: everything between this version's ## and the next ##
awk "/^## \[${VERSION}\]/{flag=1; next} /^## \[/{flag=0} flag" "$CHANGELOG" \
  | sed '/^> /d' \
  | sed '1{/^$/d}' > changelog-section.md

if [ ! -s changelog-section.md ]; then
  echo "Error: No changelog content found for version ${VERSION}" >&2
  echo "  Add a '## [${VERSION}]' section to ${CHANGELOG} before tagging." >&2
  exit 2
fi

# Build release body = changelog section + installation footer
cat changelog-section.md > release-notes.md
cat >> release-notes.md << 'FOOTER'

---

### Install

**Rust** &nbsp; `cargo add office_oxide`

**Python** &nbsp; `pip install office-oxide`

**JavaScript (WASM, universal)** &nbsp; `npm install office-oxide-wasm`

**Node.js (native)** &nbsp; `npm install office-oxide`

**Go** &nbsp;
```bash
go get github.com/yfedoseev/office_oxide/go
# fetch the native library matching your platform:
go run github.com/yfedoseev/office_oxide/go/cmd/install@latest
```

**C# / .NET** &nbsp; `dotnet add package OfficeOxide`

**CLI**
```bash
cargo binstall office_oxide_cli      # pre-built binary
brew install yfedoseev/tap/office-oxide
scoop bucket add yfedoseev https://github.com/yfedoseev/scoop-bucket && scoop install office-oxide
```

**Raw C FFI** — download the `native-<platform>-<arch>` asset below and include `include/office_oxide_c/office_oxide.h`.

### Changelog
Full history: [CHANGELOG.md](https://github.com/yfedoseev/office_oxide/blob/main/CHANGELOG.md)
FOOTER

# Cleanup
rm -f changelog-section.md

echo "Generated release-title.txt and release-notes.md for v${VERSION}"
