# `.doc` structured extraction — regression & risk analysis

Read-only analysis of the current `.doc` path against the proposed
"structure-aware second extraction path" (PAPX-driven paragraphs → real
`Element::Table` / `Element::List`, text sliced from the character stream).

No code was changed. No build/test was run.

---

## 0. Headline

**The single highest-severity risk is character-position (CP) misalignment.**
The proposed path slices "each paragraph's text out of the character stream
directly" using CP boundaries from a property index. But *neither* string that
exists today in this codebase is CP-indexed:

* `extract_text` (`src/doc/piece_table.rs:128-167`) **silently drops entire
  pieces** whose byte range is out of bounds (`:144`, `:154`) and **collapses
  UTF-16 surrogate pairs** into one `char` (`:161`). The concatenated result is
  therefore not guaranteed to have one `char` per CP.
* `sanitize_text` (`src/doc/piece_table.rs:205-218`) **deletes** `0x01`, `0x08`,
  `0x13`, `0x14`, `0x15` (`:213`). Every field in the document shortens the
  string relative to CP space.

Word documents are saturated with field characters — page numbers, dates, TOCs,
hyperlinks, cross-references, mail-merge, `SEQ`/caption numbering. Any CP→string
offset arithmetic against the sanitized text drifts by the number of deleted
control characters seen so far, so paragraph boundaries garble progressively
**worse the further into the document you go** — and it fails silently, because
the text is still valid UTF-8 and still "looks like a document". On top of that,
`&s[a..b]` on a non-`char`-boundary is a **panic**, and `sanitize_text` never
changes byte lengths uniformly (`\r`→`\n` and `\x07`→`\t` are 1:1, but the
CP1252 map at `:170-202` turns 1 byte into up to 3 UTF-8 bytes).

Everything below elaborates.

---

## 1. Text pipeline trace (raw bytes → the string `doc_to_ir` consumes)

Entry: `Document::open` → `src/lib.rs:213` → `DocDocument::open`
(`src/doc/document.rs:98-101`) → `DocDocument::from_reader`
(`src/doc/document.rs:23-95`).

| # | Step | Location | What it does / adds / drops / rewrites |
|---|------|----------|----------------------------------------|
| 1 | CFB container open | `src/doc/document.rs:24` → `src/cfb/reader.rs` | Parses the OLE2 header, FAT, DIFAT, directory. |
| 2 | Read `WordDocument` stream | `src/doc/document.rs:26-28` | `read_chain` (`src/cfb/reader.rs:261-291`) concatenates whole sectors and **pads to the sector boundary** — the returned `Vec<u8>` is usually *longer* than the real stream size. It also tolerates truncation (`n == 0 → break`, `:279`) so the buffer may be *shorter* than the FIB believes. Missing stream → `Err(MissingStream)`. |
| 3 | Parse FIB | `src/doc/document.rs:30-38` → `src/doc/fib.rs:38-103` | Requires ≥68 bytes; accepts `wIdent` `0xA5EC` (Word 97+) **and `0xA5DC` (Word 6/95)** (`fib.rs:48`). Reads `ccpText`@`0x4C`, footnote/header/comment/endnote/textbox lengths, and `fcClx`@`0x1A2` / `lcbClx`@`0x1A6` (`fib.rs:83-88`). **On any FIB error it returns an empty document, not an `Err`** (`document.rs:32-37`) — first silent-failure site. |
| 4 | Pick table stream | `src/doc/document.rs:41-56` | `fWhichTblStm` (bit 9 of flags@`0x0A`) selects `1Table`/`0Table`, with a fallback to the other one. Neither present → **silently empty document**. |
| 5 | Slice out the CLX | `src/doc/document.rs:59-73` | `clx_start = fcClx`, `clx_end = clx_start + lcbClx`, clamped to stream length (`:72`). Guard `clx_size_zero_or_oob` (`:141-143`) permits up to `stream_len + 1024` of slack, then relies on the `.min()` clamp. Out of range → **silently empty document**. |
| 6 | Parse CLX | `src/doc/piece_table.rs:28-62` | **Skips every `Grpprl` (type `0x01`) block** (`:32-38`) — these are the direct-formatting property sets referenced by `PCD.prm`, i.e. exactly the data a structured path may need. Requires a `Pcdt` (`0x02`) marker (`:41-46`), reads the `PlcPcd` size, tolerates an oversized size by clamping (`:56-60`). |
| 7 | Parse `PlcPcd` | `src/doc/piece_table.rs:71-125` | `n = (len - 4) / 12`; reads `n+1` CPs then `n` 8-byte PCDs. Extracts `fc` from PCD bytes 2..6, sets `is_compressed = fc & 0x4000_0000`. **`PCD.prm` (bytes 6..8) is read but discarded** (`:105` comment only) — no direct-formatting SPRM survives. No validation that CPs are monotonically increasing. |
| 8 | Reassemble text | `src/doc/piece_table.rs:128-167` | Per piece, `char_count = cp_end.min(max_chars) - cp_start` (`:136`); `break` once `cp_start >= max_chars` (`:132-134`). Compressed pieces: byte offset `(fc & !0x4000_0000)/2`, 1 byte/char through `cp1252_to_char` (`:170-202`, maps `0x80-0x9F` to the Windows-1252 punctuation block). Unicode pieces: UTF-16LE via `String::from_utf16_lossy` (`:161`). **Both branches silently skip the whole piece when the byte range doesn't fit the buffer** (`:144`, `:154`) — a truncated file loses text with no signal. |
| 9 | Truncate to main document | `src/doc/document.rs:85` (`max_chars = fib.text_len`) | Only `ccpText` characters are kept, so footnotes, headers/footers, comments, endnotes, and textboxes are excluded. This depends on the CP order in the FIB (main text first). |
| 10 | Sanitize | `src/doc/piece_table.rs:205-218`, called at `src/doc/document.rs:86` | `\r` (paragraph mark) → `\n`; `\x07` (**both** cell-end and row-end mark) → `\t`; `\x0C` (page/section break) → `\n`; `\x0B` (soft line break) → `\n`; `\x01`, `\x08`, `\x13`, `\x14`, `\x15` **deleted**. Everything else passes through verbatim — including `\x02`/`\x05` (note/annotation refs), `\x0E` (column break), `\x1E`/`\x1F` (non-breaking / optional hyphen), and **field instruction text**: only the `0x13/0x14/0x15` delimiters are removed, so `HYPERLINK "http://…"` and `PAGE \* MERGEFORMAT` remain in the output (see the test at `piece_table.rs:344-346`). |
| 11 | Store | `src/doc/document.rs:94` | The sanitized string is the *only* text state `DocDocument` keeps. Piece table, FIB, and CP mapping are all dropped. |
| 12a | `plain_text()` / `to_markdown()` | `src/doc/document.rs:109-138` | Return the flat string directly (markdown = trim each line, blank-line separate). **These do NOT go through the IR** (`src/lib.rs:148-153` `dispatch_inner!`). |
| 12b | `to_ir()` | `src/lib.rs:352` → `src/convert_doc.rs:4-63` | `text.lines()` (`:9`); blank lines skipped (`:11-13`); per line, a heuristic (`:17-24`): `len < 100` && doesn't end in `.`/`,` && (all alphabetic chars uppercase **or** it's the first element and `len < 60`) → `Heading` (level 1 if first else 2, `bold: true`); otherwise `Paragraph` **built from the untrimmed `line`** (`:37`) while headings use `trimmed` (`:31`). The first heading's text becomes both `metadata.title` and `Section::title` (`:43-61`). |
| 13 | Render | `src/ir_render.rs:125-158` | `render_section_plain` (`:166-175`) and `render_section_html` (`:422-436`) **prepend `Section::title`** — so for `.doc`, `to_html()` and `to_ir().plain_text()` today emit the title **twice** (once as the section title, once as the `Heading` element). Pre-existing bug that a structured rewrite will change the shape of. |

### Key structural consequences of the current pipeline

* Because **cell-end and row-end are both `0x07`**, an entire table is flattened
  into a single tab-separated **line** — `.lines()` at `convert_doc.rs:9`
  therefore emits **one `Paragraph` per table**, not one per row.
* Because `0x0B` (Shift+Enter soft break) becomes `\n`, a single Word paragraph
  containing soft breaks currently becomes **several** IR `Paragraph`s.
* Auto-numbering and bullet glyphs from `LFO`/`LST` are **not in the character
  stream at all**, so today's `.doc` output has no list markers whatsoever.

---

## 2. What breaks when the IR is built from a different point in the pipeline

Ranked by how often the affected feature appears in real Word 97-2003 files.

### R1 — CP↔string offset drift from field-character deletion (severity: critical, frequency: very high)
Fields appear in essentially every non-trivial business document. `sanitize_text`
deletes 5 control characters (`piece_table.rs:213`); `extract_text` drops whole
pieces (`:144`, `:154`) and collapses surrogate pairs (`:161`). A CP-indexed
slice against either string is wrong, and the error **accumulates**.
*User sees:* paragraphs whose text starts mid-word and ends mid-word; table cells
containing the tail of the previous cell; content shifting further out of place
toward the end of the file. No error, no warning.
*Also:* byte-slicing at a non-`char` boundary **panics** — `cp1252_to_char`
produces multi-byte UTF-8 for `0x80-0xFF`, which is ubiquitous (curly quotes,
em-dashes, accented names).
*Required:* the structured path must slice **from the piece table**, per CP
range, re-running steps 8+10 on that range — never index into the flat string.
If a CP→char index map is built instead, it must be built inside `extract_text`
while the pieces are still in scope, and must record dropped pieces.

### R2 — Content set change: headers/footers/footnotes/textboxes leaking in (severity: high, frequency: high)
Today `max_chars = ccpText` (`document.rs:85`) confines output to the main
document. The `.doc` CP space is *contiguous*: main text, then footnotes,
headers, comments, endnotes, textboxes (`Fib` already carries all those lengths,
`fib.rs:22-33`). A PAPX/FKP walk enumerates paragraphs across **all** of it. If
the new walker doesn't clamp to `[0, ccpText)`, every document with a header,
footer, page-number footer, or footnote suddenly gains that text in the IR.
*User sees:* "Page 1 of 12" and company footers interleaved into body content;
`to_html`/`save_as` output grows.

### R3 — Paragraph count / boundary semantics change (severity: high, frequency: very high)
Soft line breaks (`0x0B`) are extremely common (addresses, headings, poetry,
signature blocks). Today each becomes its own IR `Paragraph` (step 10 + 12b).
Structurally, they belong to one paragraph. The structured path will merge them.
Likewise `0x0C` page/section breaks currently split paragraphs and will not.
*User sees:* different element counts from `to_ir()`; different `.docx` output
from `save_as`; anything downstream that indexes elements positionally
(**pdf_oxide consumes this crate** — see project memory) shifts.

### R4 — Loss of `metadata.title` / `Section::title` (severity: medium-high, frequency: very high)
Today the title is manufactured by the ALL-CAPS/short-first-line heuristic
(`convert_doc.rs:17-24, 43-49`). Word 97 documents very often apply **no**
heading style — the visual heading is just bold 14pt Normal. A style/outline-level
based structured path will find zero headings for those files, so
`metadata.title` and `Section::title` both become `None`.
*User sees:* `office_oxide info file.doc` loses the title; the MCP `document_info`
tool (`crates/office_oxide_mcp/src/protocol.rs:116`) reports no title; the
title line stops being duplicated in `to_html()` (arguably a fix, but a diff).
*Guardrail:* keep the heuristic as a *title fallback* even on the structured path.

### R5 — `plain_text()` / `to_markdown()` vs `to_ir()` divergence (severity: medium-high, frequency: universal)
For `.doc`, `plain_text()` and `to_markdown()` bypass the IR entirely
(`src/lib.rs:148-153`; `src/doc/document.rs:109-138`). Adding tables and lists to
the IR only improves `to_html()`, `to_ir()`, `save_as()`, and the MCP `ir`/`html`
modes — while `office_oxide text` and `office_oxide markdown`
(`crates/office_oxide_cli/src/commands/{text,markdown}.rs`) keep emitting the old
flat text. Two "markdown" outputs for the same file that no longer agree.
*Decide explicitly:* either route DOC `to_markdown()` through the IR (a large,
separate behaviour change) or document the divergence.

### R6 — Table rendering shape change (severity: medium, frequency: high)
Once tables become `Element::Table`, `render_table_plain`
(`src/ir_render.rs:212-229`) joins cells with `\t` **and rows with `\n`**, where
today the whole table is one `\t`-joined line. `to_html()` starts emitting real
`<table>`. That is the goal, but it is a hard output diff for every table-bearing
document, and nested tables (which `0x07` alone cannot disambiguate — you need
`sprmPFInnerTableCell`/`sprmPFInnerTtp`) are a correctness trap: mis-detected
nesting produces a table with the wrong row count and no error.
Watch `TableCell::default()` → `col_span: 0, row_span: 0`
(`src/ir.rs:876-896`); most consumers do `.max(1)` (`src/docx/write.rs:1176`,
`src/create.rs:502`) but not all — set them to 1 explicitly.

### R7 — List marker duplication or invention (severity: medium, frequency: high)
Auto-numbers/bullets live in `LFO`/`LST`, not in the character stream, so today
they are absent. `render_list_markdown`/`render_list_plain`
(`src/ir_render.rs:231-247`) *synthesize* a marker. Two failure modes:
(a) a document using literal typed `"1."` or a Symbol-font bullet character in the
text stream gets **double** markers; (b) `ordered`/`start_number` guessed wrong
gives wrong numbering that reads as authoritative.

### R8 — Field instruction text (severity: medium, frequency: high)
Today the field *instructions* survive (`HYPERLINK "…"`, `PAGE \* MERGEFORMAT`);
only the `0x13/0x14/0x15` delimiters are stripped. If the structured path
"cleans this up" as a side effect, output for TOC-heavy documents changes
dramatically. If it doesn't, per-paragraph instruction text now lands inside
individual table cells where it is far more visually obvious than in a flat blob.
Either way it must be a **deliberate, separately-tested** decision.

### R9 — Word 6/95 files (severity: medium, frequency: low-moderate)
`Fib::parse` accepts `wIdent == 0xA5DC` (`fib.rs:48`) but then reads Word-97
`FibRgLw97`/`FibRgFcLcb97` offsets. For those files `ccpText` and `fcClx` are
already garbage; a PAPX pointer read the same way will be garbage too, and will
be *trusted* by a new binary walker. Gate the structured path on
`nFib >= 0x00C1`.

### R10 — Astral-plane characters (severity: low-medium, frequency: low)
`String::from_utf16_lossy` (`piece_table.rs:161`) turns a surrogate **pair**
(2 CPs) into 1 `char`. Any CP↔char arithmetic breaks for emoji / CJK Ext-B /
rare scripts. Also the reason a "1 CP = 1 char" assumption is unsound in general.

### R11 — Tracked changes / deleted pieces (severity: low, frequency: low-moderate)
Neither path distinguishes revision-marked text. A structured path that gains
access to `prm`/SPRMs may start (inconsistently) honouring or exposing them.

---

## 3. Silent-failure surface

**As described, the design degrades silently in at least six places, and the
existing code sets a bad precedent it would inherit.**

Existing silent-failure sites on the `.doc` path — all return an
**empty document with `Ok(...)`**, never an `Err`, never a log line:

* `src/doc/document.rs:30-38` — FIB unparseable.
* `src/doc/document.rs:48-56` — no table stream.
* `src/doc/document.rs:62-70` — CLX out of range.
* `src/doc/document.rs:74-82` — CLX/piece-table malformed.
* `src/doc/piece_table.rs:144`, `:154` — a piece whose bytes don't fit is
  dropped mid-document; the surrounding text still returns fine.
* `src/doc/piece_table.rs:56-58` — `if pos + pcdt_size > data.len() { /* Be
  tolerant */ }` — an empty `if` body; the comment is the whole handler.

This directly contradicts `AGENTS.md:37-38` ("Fail loudly, never fall back to a
silent plausible-but-wrong result").

### What a partial structured parse yields, unguarded
* **PAPX FKP found, but half the bins unreadable** → some paragraphs are
  structured, the rest vanish (they are not in any FKP the walker could read).
  Output: a document that is *missing* content, with no signal.
* **Table start detected, table end not** → the `0x07` run consumed into one
  `Table` swallows the following body paragraphs as extra rows.
* **List properties partially read (`ilfo` yes, `LFO`/`LST` no)** → `ordered`
  defaults to `false`, so numbered lists render as bullets.
* **Structured path silently falls back to the line heuristic per-document** →
  two documents from the same source produce structurally different IR with no
  way for a caller to know which path ran.
* **CP mapping drifts (R1)** → the worst case: nothing is missing, everything is
  *slightly wrong*. This is undetectable without a byte-level baseline.

### Guardrails I would require

1. **Total-text conservation check.** After building structured paragraphs,
   compare `structured_paragraphs.concat()` (whitespace-normalized) against the
   existing flat `sanitize_text` output. If more than a small, explicitly-chosen
   epsilon of characters is missing or added, **discard the structured result and
   fall back**, and log at `warn`. This single check catches R1, R2, and the
   partial-FKP case. It is cheap — both strings already exist.
2. **CP coverage check.** The structured walk must cover `[0, ccpText)`
   contiguously with no gaps and no overlaps. A gap or overlap → fall back.
   Explicitly clamp the upper bound to `ccpText`; never walk past it (R2).
3. **Explicit path signal.** Record which path produced the IR (an enum on
   `DocDocument`, or at minimum a `log::info!`/`log::warn!`). Silently choosing
   between two structurally different extractors is not acceptable for a library
   whose consumers (`pdf_oxide`) index into the result.
4. **All-or-nothing per document.** Do not mix: a per-paragraph fallback
   produces IR where some tables are `Element::Table` and others are tab-joined
   `Paragraph`s, which is worse for consumers than either pure mode.
5. **Never derive text from the sanitized flat string by index.** Enforce this
   structurally — have the structured extractor take `&[Piece]` + `&word_doc`
   and produce its own per-range text, so no API even exists to mis-slice.
6. **Replace the four `Ok(empty)` returns** in `document.rs` (or at least add
   `log::warn!` at each) as part of this work; the new path adds more of them
   otherwise.
7. **Preserve the title fallback.** If the structured path finds no heading,
   run the existing heuristic on the structured paragraph texts so
   `metadata.title` does not silently disappear (R4).

---

## 4. Panic-safety rulebook for a new binary-structure parser here

### Conventions the codebase already uses (follow these)

`src/ppt/persist.rs` and `src/ppt/records.rs` are the house style — copy them.

* **Slice with `.get(a..b)?`, not `&data[a..b]`.**
  `src/ppt/persist.rs:112`, `:121`, `:135`, `:139`, `:140`, `:152`, `:161`, `:172`.
* **`saturating_add` before every `.min(len)`.**
  `src/ppt/persist.rs:119`, `:159`; `src/ppt/records.rs:132`; `src/ppt/text.rs:331`.
* **Bound every record by *its own* declared length, clamped to the enclosing
  slice** — never let a corrupt length swallow siblings.
  `src/ppt/records.rs:88-143` (the doc comment there states the rule explicitly).
* **Named constant caps on every unbounded loop / recursion.**
  `MAX_EDIT_CHAIN_LEN = 4096` (`src/ppt/persist.rs:25`),
  `MAX_SHAPE_DEPTH = 64` (`src/ppt/text.rs:7`, checked at `:209`, `:353`).
* **Cycle detection on any offset chain** — `HashSet<usize>` of visited offsets
  (`src/ppt/persist.rs:70-72`); same idea as CFB's `visited > max_sectors`
  (`src/cfb/reader.rs:268-270`, `:303-306`).
* **Guarantee forward progress**: an iterator whose cursor could stay put on a
  zero-length record must still advance (`src/ppt/records.rs:128-135` advances
  past the 8-byte header even when `rec_len == 0`).
* **`Option`-returning helpers with `?`** for "malformed → give up on this
  sub-structure", reserving `Result` for "this document is unusable".
* **`saturating_sub` for lengths** (`src/cfb/reader.rs:315`).
* **Return `Err`, never `panic!`/`unwrap`/`expect`, from parsing code**
  (`AGENTS.md:30-33`).

### Additional rules this specific parser must follow

1. **Never index a `String`/`&str` by a number derived from file content.**
   Slice from `&[u8]` and build the `String` yourself. If you must map CP→char,
   use `char_indices()`/`chars().count()` — not byte arithmetic.
2. **Validate PLC/FKP shape before trusting it.** A `PLCF` of `n` entries is
   `(n+1)*4 + n*cbStruct` bytes: solve for `n` from the *actual slice length*
   (as `parse_plc_pcd` does at `piece_table.rs:77`), then re-verify the total
   fits, and reject `n == 0`.
3. **Assert CP monotonicity.** Reject (or clamp) any `cp[i+1] <= cp[i]`. This is
   the check whose absence produces the existing subtraction overflow (below).
4. **Clamp CP ranges to `[0, ccpText)`** before doing anything with them.
5. **Cap allocations by input size.** `Vec::with_capacity` must never be sized
   from an unvalidated file field. A 4-byte length field can request 4 GiB.
   Prefer capping the *count* against `slice.len() / entry_size`.
6. **Cap table nesting depth and list nesting depth** with named constants, in
   the style of `MAX_SHAPE_DEPTH`.
7. **Anchor FIB offsets on a single documented base.** `FibRgFcLcb97` begins at
   absolute `0x9A`; `fcClx` at absolute `0x1A2` is `0x9A + 66*4`. Read every new
   pointer as `0x9A + index*4` and validate `data.len()` covers it
   (`fib.rs:106-117`'s `read_u32` returning `0` on OOB is the right shape).
8. **Gate on `nFib`.** Only run the structured path for `wIdent == 0xA5EC` and a
   Word-97+ `nFib`.
9. **Extend `fuzz/fuzz_targets/fuzz_parse.rs`.** See §5 — the current target does
   **not** exercise this code path at all.

### Places the *existing* code already violates these rules

| Location | Violation |
|----------|-----------|
| `src/doc/piece_table.rs:136` | `piece.cp_end.min(max_chars) - piece.cp_start` — **unchecked `u32` subtraction**. If `cp_end < cp_start` (non-monotonic CP array — trivially producible in a crafted file) this **panics in debug** and wraps to a huge `char_count` in release, which then drives the byte-range checks at `:144`/`:154`. Nothing validates CP ordering in `parse_plc_pcd` (`:71-125`). **This is a live bug and a fuzz-reachable panic.** |
| `src/doc/piece_table.rs:32-38` | `pos += 3 + size` with `size` from a `u16` — no re-check that `pos` stayed in bounds before the next `data[pos]` read. The loop condition catches it, but only by luck of ordering; and `pos` can jump past `data.len()` leaving the `0x02` check (`:41`) to report a confusing error. |
| `src/doc/piece_table.rs:56-58` | `if pos + pcdt_size > data.len() { /* Be tolerant */ }` — **empty `if` body**. Dead code that documents an unhandled case. |
| `src/doc/piece_table.rs:90-111` | Direct `data[...]` indexing throughout instead of `.get()`. Safe *today* only because of the single aggregate check at `:83`; any edit to the index arithmetic silently loses that protection. |
| `src/doc/document.rs:59-60` | `clx_start + fib.clx_size as usize` computed **before** any bounds check; safe on 64-bit (two `u32`s), a latent overflow on 32-bit/WASM32 targets — and the crate ships a WASM binding. |
| `src/doc/document.rs:141-143` | `clx_size_zero_or_oob` allows `stream_len + 1024` of slack with the comment "allow some slack" — an unprincipled tolerance that lets a bogus `lcbClx` through to the clamp. |
| `src/doc/fib.rs:69-81` | A 13-line comment block that computes the CLX offset **two different ways** (`0x23C` vs `0x1A2`) and contradicts the code. The code is right; the comment is wrong. A new parser copying the comment's `0x9A + …` arithmetic will read the wrong fields. |
| `src/doc/fib.rs:48` | Accepts `0xA5DC` (Word 6/95) then applies Word-97 field offsets. |
| `src/doc/images.rs:16-51` | Byte-by-byte scan with `data[pos + …]` indexing and `data_start + rec_len` before the `.min()` (`:28`); `img_start = data_start + skip` (`:31`) is unchecked before the `img_start < data_end` comparison. Bounded in practice by the loop guard, but not by construction. |
| `src/lib.rs:132-141` | The panic *net* (`with_parse_stack` catching a thread panic) only exists when `needs_stack_thread()` is true — i.e. **only on Unix with a small `RLIMIT_STACK`**. In an ordinary Rust binary the closure runs inline (`:140`) and any parser panic propagates to the caller. Do not rely on it. |

---

## 5. Testing: what exists, what's missing, what I'd require

### What the current setup provides

* **Unit tests inside the `.doc` modules only**: `src/doc/fib.rs:119-170`
  (4 tests), `src/doc/piece_table.rs:220-370` (8 tests),
  `src/doc/images.rs:109-177` (5 tests), `src/doc/document.rs:155-233`
  (10 tests — of which the 8 IR tests construct `DocDocument` **directly from a
  `String`**, bypassing the entire binary pipeline).
* **A fuzz target** (`fuzz/fuzz_targets/fuzz_parse.rs:12-23`) that feeds
  arbitrary bytes to `Document::from_reader` for all six formats.

### What it does **not** provide

1. **No `.doc` integration test at all.** `tests/` contains
   `docx_integration.rs`, `xlsx_integration.rs`, `pptx_integration.rs`,
   `office_integration.rs`, `core_integration.rs`, `write_integration.rs`,
   `ir_roundtrip_idempotence.rs`, `xlsx_row_cap.rs` — **none exercises DOC**.
   Every `.doc` assertion in the repo is a unit test on a hand-built byte blob
   or a hand-written `String`.
2. **`samples-doc/{simple-list,table,table-merges}.doc` are referenced by
   nothing.** A repo-wide grep finds zero uses. They are fixtures with no test.
3. **`doc_to_ir` is never fuzzed.** The fuzz target calls `from_reader` only —
   `to_ir()` is never invoked, so `convert_doc.rs` and everything the structured
   path would add there gets **zero** fuzz coverage. If the structure parsing
   lands in `doc_to_ir` rather than `from_reader`, it will be entirely unfuzzed.
4. **No end-to-end test from real `.doc` bytes to IR.** Nothing would catch R1
   (offset drift), R2 (header/footer leakage), or R3 (paragraph count change).
5. **No `.doc` round-trip/idempotence coverage.** `ir_roundtrip_idempotence.rs`
   covers DOCX/XLSX/PPTX only — `save_as` from `.doc` to `.docx` is untested.
6. **No golden/snapshot baseline** for `.doc` extraction, so "did the text
   change?" is unanswerable except by eye.
7. **Debug vs release**: `[profile.release]` (`Cargo.toml:155-160`) does not set
   `overflow-checks`, so the `piece_table.rs:136` subtraction wraps silently in
   release builds and only panics under `cargo test`/fuzz. Good for fuzzing;
   means users hit corruption rather than a crash.

### What I would require before believing this change is safe

1. **Byte-level regression fixtures, built in code** (per `AGENTS.md:22-25`):
   minimal synthetic CFB `.doc` streams for — a plain paragraph; a paragraph with
   a soft line break (`0x0B`); a paragraph containing a field
   (`\x13 … \x14 … \x15`); a 2×2 table; a table with a vertically merged cell;
   a nested table; a bulleted list; a numbered list; a document with a header and
   a footnote (to prove they stay out); a compressed-and-Unicode mixed piece
   table; a piece table with a non-monotonic CP array (must not panic).
2. **A text-conservation property test**: for every fixture, the concatenated
   structured text must equal the flat `sanitize_text` output modulo
   whitespace/marker characters. This is the single test that catches R1.
3. **A "no new content" test**: assert `to_ir().plain_text()` for a
   header/footer/footnote fixture contains none of that text (R2).
4. **A path-selection test**: assert which extraction path ran for each fixture
   (requires guardrail #3 above) — otherwise a silent fallback makes all the
   structured assertions vacuously pass on the old path.
5. **Fuzz coverage extension**: add `let _ = doc.to_ir();` (and `to_html()`) to
   `fuzz/fuzz_targets/fuzz_parse.rs`, plus a dedicated structured-parser target
   fed the raw table-stream bytes. Run it before merge, not after.
6. **Fix `piece_table.rs:136` first**, with its own regression test — a new
   parser that reads the same CP array will hit the same malformed input.
7. **Wire `samples-doc/*.doc` into an integration test**, or delete them.
   Whatever these three files were built for, they should assert something.
8. **A corpus diff report** (per `AGENTS.md:26-29`): for N real `.doc` files,
   record before/after `plain_text()`, `to_markdown()`, IR element counts, and
   IR element-type histogram; review every file whose text length changed by more
   than a trivial amount. Text length changes are the R1 signal.
9. **`.doc` → `.docx` `save_as` smoke test**, since `Document::save_as`
   (`src/lib.rs:~410`) routes legacy formats through the IR — a structured IR
   change silently changes the conversion product.
10. **A decision, in writing, on R5** (whether DOC `to_markdown()` starts going
    through the IR) before any code lands, because retrofitting it later is
    another breaking output change for the same users.
