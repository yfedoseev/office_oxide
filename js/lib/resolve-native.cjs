// Resolve the platform's native library.
//
// The six prebuilt libraries used to ship inside the main package, so every
// install downloaded all of them — five of which can never be loaded. They
// are now published as one package per platform, declared in
// `optionalDependencies` with `os`/`cpu` fields so npm installs only the
// matching one.
//
// Resolution order, first hit wins:
//   1. `OFFICE_OXIDE_LIB` — an explicit override, for custom builds.
//   2. The platform package `office-oxide-<platform>-<arch>`.
//   3. `prebuilds/<platform>-<arch>/` inside this package — the previous
//      layout, kept so an install that still carries bundled prebuilds (or
//      a checkout of this repo) keeps working.
//   4. The bare library name, letting the OS loader search its own paths.
'use strict';

const path = require('node:path');
const process = require('node:process');

const EXT =
  process.platform === 'win32' ? '.dll' :
  process.platform === 'darwin' ? '.dylib' : '.so';
const PREFIX = process.platform === 'win32' ? '' : 'lib';

/** File name of the shared library on this platform. */
function libFileName() {
  return `${PREFIX}office_oxide${EXT}`;
}

/** npm package name holding this platform's prebuilt library. */
function platformPackageName() {
  return `office-oxide-${process.platform}-${process.arch}`;
}

/**
 * Candidate paths in resolution order.
 *
 * @param {(id: string) => string} resolve - a `require.resolve` bound to the
 *   caller's module, so the platform package resolves from the right place.
 */
function candidatePaths(resolve) {
  const paths = [];
  if (process.env.OFFICE_OXIDE_LIB) {
    paths.push(process.env.OFFICE_OXIDE_LIB);
  }
  try {
    paths.push(resolve(`${platformPackageName()}/${libFileName()}`));
  } catch {
    // The optional dependency is not installed — fall through.
  }
  try {
    const hereDir = path.dirname(resolve('office-oxide/package.json'));
    paths.push(path.join(
      hereDir, 'prebuilds',
      `${process.platform}-${process.arch}`,
      libFileName(),
    ));
  } catch {
    // Not resolvable from here (e.g. a repo checkout); the caller adds its
    // own package-relative path below.
  }
  paths.push(libFileName());
  return paths;
}

module.exports = { candidatePaths, libFileName, platformPackageName };
