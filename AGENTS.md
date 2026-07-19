# AGENTS.md — guidance for AI coding agents

OfficeOxide is a pure-Rust parser/extractor for Office documents — OOXML
(`.docx`/`.xlsx`/`.pptx`, a zip of XML) and the legacy CFB compound formats
(`.doc`/`.xls`/`.ppt`). Read [`CONTRIBUTING.md`](CONTRIBUTING.md) in full before
proposing changes. The essentials that apply to agent-assisted work:

## Contribution rules (must follow)

1. **Issue-first.** Non-trivial changes require an issue the maintainer has
   accepted, with the approach agreed *before* code is written. Do not open
   drive-by PRs — they may be closed without review. Bug/typo/docs fixes may
   skip this; the reproducer may not.
2. **No autonomous PRs.** An autonomous agent must not open issues or pull
   requests on its own. A human directs the work, reviews every line, and is
   accountable for it. Disclose AI assistance in the PR; the human owns
   correctness, licensing, and provenance. Non-trivial contributions are signed
   under the project's [CLA](CLA.md).
3. **The human must understand and be able to explain every line.** If the
   change can't be explained without the AI, it isn't ready.
4. **Every bug fix ships a regression test that fails before the fix.** Build the
   reproducer as a **minimal synthetic document in code** where practical (a
   small OOXML part or CFB stream) — do not commit third-party/customer/real
   documents that you don't have the rights to redistribute.
5. **Prove no corpus regression.** Office parsers consume untrusted input;
   heuristics regress silently. For any change to parsing/extraction/IR, run the
   test suite **and** your own corpus of representative real `.docx/.xlsx/.pptx/
   .doc/.xls/.ppt` files (the project's corpus is not distributed — bring your
   own), and report what you tested in the PR.
6. **Robustness contract:** no input may panic, overflow, or hang. Malformed
   documents (bad zip central directory, cyclic OPC relationships, bogus CFB
   FAT chains, oversized dimensions) must surface as `Err`, not a crash. A fuzz
   target lives in `fuzz/` — extend it when you touch a parser.
7. **House rules:** no issue/PR numbers or contributor/company names in code,
   comments, or fixture names — name tests by defect *class*; credit reporters
   only in `CHANGELOG.md`. One logical change per PR. Fail loudly, never fall
   back to a silent plausible-but-wrong result.
8. **Green across the feature matrix**, not just the default `cargo test` — run
   the tiers your change touches (bindings, optional features).

## Format references
- **OOXML** — ECMA-376 / ISO/IEC 29500 (Office Open XML): the `.docx/.xlsx/.pptx`
  package is an OPC (Open Packaging Conventions) zip of parts + relationships.
- **CFB** — MS-CFB (Compound File Binary) underlies the legacy `.doc/.xls/.ppt`.
- Prefer spec-accurate parsing over guessing; when a document violates the spec,
  degrade gracefully with a warning rather than a panic.
