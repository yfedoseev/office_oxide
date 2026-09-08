#!/usr/bin/env node
// Post-install guard. When a user installs office-oxide via npm, we expect a
// prebuilt native library for their platform under prebuilds/<platform>-<arch>/.
// If none exists, fall back to a clear error — but don't fail the install
// itself, so downstream consumers who intend to set OFFICE_OXIDE_LIB manually
// can still complete `npm install`.
'use strict';

const fs = require('node:fs');
const path = require('node:path');
const process = require('node:process');

const ext =
  process.platform === 'win32' ? '.dll' :
  process.platform === 'darwin' ? '.dylib' : '.so';
const prefix = process.platform === 'win32' ? '' : 'lib';

// The platform package installed via optionalDependencies is the normal
// source; `prebuilds/` inside this package is the previous layout, kept so
// an existing install or a repo checkout still works.
const platformPackage = `office-oxide-${process.platform}-${process.arch}`;
const candidates = [];
try {
  candidates.push(require.resolve(`${platformPackage}/${prefix}office_oxide${ext}`));
} catch {
  // Optional dependency absent for this platform.
}
candidates.push(path.join(
  __dirname, '..', 'prebuilds',
  `${process.platform}-${process.arch}`,
  `${prefix}office_oxide${ext}`,
));

if (candidates.some((c) => fs.existsSync(c))) {
  process.exit(0);
}

process.stderr.write(
  `\n[office-oxide] No prebuilt native library for ${process.platform}-${process.arch}.\n` +
  `  Tried:\n    ${candidates.join('\n    ')}\n` +
  `  Set OFFICE_OXIDE_LIB=/path/to/liboffice_oxide.{so,dylib,dll}\n` +
  `  or build from source in the office_oxide monorepo:\n` +
  `    cargo build --release --lib\n` +
  `    export OFFICE_OXIDE_LIB=$(pwd)/target/release/liboffice_oxide.so\n\n`,
);
// Exit 0 so `npm install` still succeeds for headless CI / custom setups.
process.exit(0);
