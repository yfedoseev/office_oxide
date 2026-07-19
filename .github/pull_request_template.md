<!--
Read CONTRIBUTING.md first. The two rules that matter most:
  1. Non-trivial PRs must reference an issue the maintainer has ACCEPTED.
     Drive-by PRs with no accepted issue may be closed without detailed review.
  2. Any parsing/extraction/IR change must be proven not to regress on a corpus
     of real Office files (yours — the project corpus is private).
Write this PR in your own words. Delete the guidance comments as you fill it in.

NOTE: a PR opened without filling in this template is closed automatically —
just fill it in and reopen. (Maintainers, drafts, and `skip-template-check` are
exempt.)
-->

## Linked issue

<!-- Required for features/behavior changes. Bug/typo/docs fixes may omit. -->
Closes #

## What and why

<!-- In your own words: what problem does this solve, and why this approach? -->

## Type of change

- [ ] Bug fix
- [ ] New feature (has an accepted issue: #___)
- [ ] Performance
- [ ] Refactor / internal
- [ ] Docs / CI / chore
- [ ] Breaking change

## Tests

- [ ] I added a test that **fails before this change and passes after**
      (revert-checked). For bug fixes the reproducer is a **minimal synthetic
      document built in code** where practical — no third-party/customer file is
      committed.
- [ ] Test named by defect **class**, not an issue/PR number; no
      contributor/company names in code or fixtures.
- [ ] `cargo test` passes, plus the affected feature tiers / bindings: ____
- [ ] `cargo fmt --check` and `cargo clippy -- -D warnings` are clean.

## Regression on real Office files

<!-- Required for ANY parsing / extraction / IR change.
     The project corpus is private; use your OWN corpus. -->

**Corpus I tested** (count + kinds + source): <!-- e.g. "~80 files: docx/xlsx/pptx
+ a few legacy .doc/.xls, real-world, from my own collection" -->

- [ ] Ran the extraction over my corpus on my branch vs **main** — no unintended
      regressions (no dropped content, garbled text, or crashes).
- [ ] Checked against the **latest release** as well.

**Diff summary:** <!-- what changed, and confirmation it's only the intended fix -->

<!-- N/A only if this PR touches no parsing/extraction/IR code. -->

## AI assistance disclosure

- [ ] AI assistance: **none**, or **assisted** — tool: ______, extent: ______
- [ ] I understand and can explain every line; the description and my review
      replies are written by me, not generated. This PR is not fully or
      predominantly AI-generated.

## Checklist

- [ ] One logical change (no bundled refactor/perf/correctness).
- [ ] Commits follow Conventional Commits.
- [ ] Non-trivial contribution: I will sign the CLA when the bot asks (see CLA.md).
- [ ] Docs/`CHANGELOG.md` updated if user-facing; reporters credited in the
      CHANGELOG (not in code).
