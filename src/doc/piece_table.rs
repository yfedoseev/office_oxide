//! Piece table parsing for Word binary documents.
//!
//! The piece table maps character positions to byte ranges in the WordDocument stream.
//! Each piece can be either:
//! - Compressed (CP1252): 1 byte per character, fc has bit 30 set, actual offset = (fc & ~0x40000000) / 2
//! - Unicode (UTF-16LE): 2 bytes per character, fc is used directly

use super::error::{DocError, Result};

/// A single piece descriptor.
#[derive(Debug, Clone)]
pub struct Piece {
    /// Character position range start (inclusive).
    pub cp_start: u32,
    /// Character position range end (exclusive).
    pub cp_end: u32,
    /// File offset in the WordDocument stream.
    pub fc: u32,
    /// Whether this piece uses compressed (CP1252) encoding.
    pub is_compressed: bool,
}

/// Parse the CLX structure to extract the piece table.
///
/// The CLX contains:
/// - Optional Grpprl entries (type 0x01): skip them.
/// - Pcdt entry (type 0x02): the piece table.
pub fn parse_clx(data: &[u8]) -> Result<Vec<Piece>> {
    let mut pos = 0;

    // Skip Grpprl entries.
    while pos < data.len() && data[pos] == 0x01 {
        if pos + 3 > data.len() {
            return Err(DocError::InvalidPieceTable("Grpprl truncated".into()));
        }
        let size = u16::from_le_bytes([data[pos + 1], data[pos + 2]]) as usize;
        pos += 3 + size;
    }

    // Now we should be at the Pcdt (type 0x02).
    if pos >= data.len() || data[pos] != 0x02 {
        return Err(DocError::InvalidPieceTable(format!(
            "expected Pcdt (0x02) at offset {pos}, found {:?}",
            data.get(pos)
        )));
    }
    pos += 1;

    if pos + 4 > data.len() {
        return Err(DocError::InvalidPieceTable("Pcdt size truncated".into()));
    }
    let pcdt_size =
        u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]) as usize;
    pos += 4;

    if pos + pcdt_size > data.len() {
        // Be tolerant — use what we have.
    }

    let pcd_data = &data[pos..data.len().min(pos + pcdt_size)];
    parse_plc_pcd(pcd_data)
}

/// Parse the PlcPcd structure (array of CPs + array of PCDs).
///
/// Structure:
/// - (n+1) u32 character positions (CPs)
/// - n PCD entries (8 bytes each)
///
/// Where n = (size - 4) / 12 (solve for: (n+1)*4 + n*8 = size)
fn parse_plc_pcd(data: &[u8]) -> Result<Vec<Piece>> {
    if data.len() < 8 {
        return Err(DocError::InvalidPieceTable("PlcPcd too small".into()));
    }

    // n pieces: (n+1)*4 + n*8 = data.len() → n = (data.len() - 4) / 12
    let n = (data.len() - 4) / 12;
    if n == 0 {
        return Ok(Vec::new());
    }

    let cp_array_size = (n + 1) * 4;
    if cp_array_size + n * 8 > data.len() {
        return Err(DocError::InvalidPieceTable("PlcPcd size mismatch".into()));
    }

    let mut pieces = Vec::with_capacity(n);

    for i in 0..n {
        let cp_start = u32::from_le_bytes([
            data[i * 4],
            data[i * 4 + 1],
            data[i * 4 + 2],
            data[i * 4 + 3],
        ]);
        let cp_end = u32::from_le_bytes([
            data[(i + 1) * 4],
            data[(i + 1) * 4 + 1],
            data[(i + 1) * 4 + 2],
            data[(i + 1) * 4 + 3],
        ]);

        // PCD at offset cp_array_size + i * 8.
        let pcd_offset = cp_array_size + i * 8;
        // PCD structure: [u16 unused][u32 fc][u16 prm]
        let fc = u32::from_le_bytes([
            data[pcd_offset + 2],
            data[pcd_offset + 3],
            data[pcd_offset + 4],
            data[pcd_offset + 5],
        ]);

        // Bit 30 of fc indicates compressed encoding.
        let is_compressed = (fc & 0x40000000) != 0;

        // A non-monotonic CP range would underflow later subtractions
        // (extract_text / decode_cp_range / fc_to_cp). Reject it here so the
        // malformed piece table surfaces as `Err`, not a panic (AGENTS.md #6).
        if cp_end < cp_start {
            return Err(DocError::InvalidPieceTable("non-monotonic piece CP range".into()));
        }

        pieces.push(Piece {
            cp_start,
            cp_end,
            fc,
            is_compressed,
        });
    }

    Ok(pieces)
}

/// Extract text from the WordDocument stream using the piece table.
/// `lid` is the FIB's language id (`Fib::lid`), which selects the
/// codepage compressed (8-bit) runs decode with (issue #310).
pub fn extract_text(word_doc: &[u8], pieces: &[Piece], max_chars: u32, lid: u16) -> String {
    extract_text_range(word_doc, pieces, 0, max_chars, lid)
}

/// Whether `pieces` fully covers `[0, text_len)` with no gaps.
///
/// A well-formed piece table starts at CP 0 and each piece picks up exactly
/// where the previous one ended, all the way to (at least) `text_len`. When
/// that doesn't hold — the first piece starts past 0, a gap sits between
/// two pieces, or the pieces stop short of `text_len` — every CP the gap
/// covers is silently absent from `extract_text`'s output with no error:
/// a `.doc` can end up 84% shorter than its own FIB says it is, and look
/// identical to a document that genuinely has that little text (issue
/// #230). This doesn't attempt to recover the missing text (a fallback
/// byte-range read risks producing *wrong*, garbled text for other files,
/// which is worse than an accurate partial extraction) — it only gives a
/// caller a way to tell the two cases apart.
pub fn covers_declared_length(pieces: &[Piece], text_len: u32) -> bool {
    if text_len == 0 {
        return true;
    }
    let mut expected_start = 0u32;
    for piece in pieces {
        if piece.cp_start != expected_start {
            return false;
        }
        expected_start = piece.cp_end;
        if expected_start >= text_len {
            return true;
        }
    }
    false
}

/// Extract the text for a character-position range `[cp_start, cp_end)`.
///
/// The piece table addresses one contiguous character space that holds the
/// main document *and* every subdocument in a fixed order: main (`ccpText`),
/// footnotes (`ccpFtn`), headers/footers (`ccpHdd`), comments (`ccpAtn`),
/// endnotes (`ccpEdn`), text boxes (`ccpTxbx`) and header text boxes
/// (`ccpHdrTxbx`). Reading only `[0, ccpText)` — which is all
/// [`extract_text`] does — leaves every one of those unread, which is why
/// the FIB's `ccp*` lengths were parsed and then never used.
pub fn extract_text_range(
    word_doc: &[u8],
    pieces: &[Piece],
    range_start: u32,
    range_end: u32,
    lid: u16,
) -> String {
    let mut text = String::new();
    if range_end <= range_start {
        return text;
    }
    let max_chars = range_end;

    for piece in pieces {
        if piece.cp_start >= max_chars {
            break;
        }
        if piece.cp_end <= range_start {
            continue;
        }
        // Clip the piece to the requested range.
        let skip = range_start.saturating_sub(piece.cp_start);
        let char_count = piece.cp_end.min(max_chars) - piece.cp_start - skip;

        if piece.is_compressed {
            // Compressed: 1 byte per character, CP1252.
            // Actual byte offset = (fc & ~0x40000000) / 2
            let byte_offset = ((piece.fc & !0x40000000) / 2) as usize + skip as usize;
            let byte_count = char_count as usize;

            if byte_offset + byte_count <= word_doc.len() {
                for &b in &word_doc[byte_offset..byte_offset + byte_count] {
                    text.push(super::codepage::decode_byte(b, lid));
                }
            }
        } else {
            // Unicode: 2 bytes per character, UTF-16LE.
            let byte_offset = piece.fc as usize + skip as usize * 2;
            let byte_count = char_count as usize * 2;

            if byte_offset + byte_count <= word_doc.len() {
                let chars: Vec<u16> = (0..char_count as usize)
                    .map(|i| {
                        let o = byte_offset + i * 2;
                        u16::from_le_bytes([word_doc[o], word_doc[o + 1]])
                    })
                    .collect();
                text.push_str(&String::from_utf16_lossy(&chars));
            }
        }
    }

    text
}

/// The largest CP of `piece` that is actually backed by bytes in a stream of
/// `word_doc_len` bytes.
///
/// A malformed piece can declare a CP range whose bytes lie outside the
/// stream. Decoding must clamp to this bound so the work stays proportional to
/// the backed range, not the (possibly huge) declared range — a DoS guard
/// (AGENTS.md rule 6). `decode_cp_range` uses this to bound each segment; it is
/// exposed as a pure helper so the bound itself can be asserted directly (the
/// output is identical with or without the clamp, so only the bound pins it).
pub(crate) fn piece_backed_cp_end(piece: &Piece, word_doc_len: usize) -> u32 {
    let base = if piece.is_compressed {
        (piece.fc & !0x4000_0000) as u64 / 2
    } else {
        piece.fc as u64
    };
    let avail = (word_doc_len as u64).saturating_sub(base);
    let extra = if piece.is_compressed {
        avail
    } else {
        // Each Unicode char is 2 bytes; `avail` bytes hold `avail / 2` full
        // chars (a trailing lone byte is not a complete char and is skipped,
        // matching the per-cp `off + 1 < len` check in `decode_cp_range`).
        avail / 2
    };
    piece
        .cp_start
        .saturating_add(extra.min(u64::from(u32::MAX)) as u32)
}

/// Decode the text in character-position range `[cp_start, cp_end)` straight
/// from `word_doc`, walking the piece table.
///
/// Unlike [`extract_text`] this is *per range*: callers slice by CP without
/// first collapsing the whole document into a flat `String`, so a surrogate
/// pair (2 UTF-16 code units) in a Unicode piece is decoded into one `char`
/// exactly where it belongs, and compressed (CP1252) pieces are decoded by
/// their own stride. A truncated/out-of-range segment is skipped per-CP
/// rather than dropping the entire piece, which keeps later ranges aligned.
pub(crate) fn decode_cp_range(
    word_doc: &[u8],
    pieces: &[Piece],
    cp_start: u32,
    cp_end: u32,
    lid: u16,
) -> String {
    let mut out = String::new();
    if cp_end <= cp_start {
        return out;
    }
    for piece in pieces {
        if piece.cp_end <= cp_start || piece.cp_start >= cp_end {
            continue;
        }
        let seg_start = cp_start.max(piece.cp_start);
        let mut seg_end = cp_end.min(piece.cp_end);
        // Clamp the segment to the bytes that are actually present in the
        // stream. A malformed piece can declare a CP range whose bytes are not
        // in `word_doc`; without this clamp the loop would iterate the entire
        // (possibly huge) declared range while producing nothing — a DoS
        // (AGENTS.md rule 6). The clamp preserves output (unbacked CPs yield no
        // characters) while bounding the work to the backed range.
        let backed_cp = piece_backed_cp_end(piece, word_doc.len());
        seg_end = seg_end.min(backed_cp);
        if seg_end <= seg_start {
            continue;
        }
        if piece.is_compressed {
            let base = ((piece.fc & !0x4000_0000) / 2) as usize;
            for cp in seg_start..seg_end {
                let off = base + (cp - piece.cp_start) as usize;
                if off < word_doc.len() {
                    out.push(super::codepage::decode_byte(word_doc[off], lid));
                }
            }
        } else {
            let base = piece.fc as usize;
            let mut u16s: Vec<u16> = Vec::with_capacity((seg_end - seg_start) as usize);
            for cp in seg_start..seg_end {
                let off = base + (cp - piece.cp_start) as usize * 2;
                if off + 1 < word_doc.len() {
                    u16s.push(u16::from_le_bytes([word_doc[off], word_doc[off + 1]]));
                }
            }
            out.push_str(&String::from_utf16_lossy(&u16s));
        }
    }
    out
}

/// Convert a CP1252 byte to a Unicode char.
pub(crate) fn cp1252_to_char(b: u8) -> char {
    // CP1252 is identical to Latin-1 except for bytes 0x80-0x9F.
    match b {
        0x80 => '\u{20AC}', // €
        0x82 => '\u{201A}', // ‚
        0x83 => '\u{0192}', // ƒ
        0x84 => '\u{201E}', // „
        0x85 => '\u{2026}', // …
        0x86 => '\u{2020}', // †
        0x87 => '\u{2021}', // ‡
        0x88 => '\u{02C6}', // ˆ
        0x89 => '\u{2030}', // ‰
        0x8A => '\u{0160}', // Š
        0x8B => '\u{2039}', // ‹
        0x8C => '\u{0152}', // Œ
        0x8E => '\u{017D}', // Ž
        0x91 => '\u{2018}', // '
        0x92 => '\u{2019}', // '
        0x93 => '\u{201C}', // "
        0x94 => '\u{201D}', // "
        0x95 => '\u{2022}', // •
        0x96 => '\u{2013}', // –
        0x97 => '\u{2014}', // —
        0x98 => '\u{02DC}', // ˜
        0x99 => '\u{2122}', // ™
        0x9A => '\u{0161}', // š
        0x9B => '\u{203A}', // ›
        0x9C => '\u{0153}', // œ
        0x9E => '\u{017E}', // ž
        0x9F => '\u{0178}', // Ÿ
        _ => b as char,
    }
}

/// A `HYPERLINK` field's display-text span, as a byte range into the string
/// [`sanitize_text_with_hyperlinks`] returned it alongside, paired with the
/// URL parsed from the field's own instruction text (issue #249).
#[derive(Debug, Clone, PartialEq)]
pub struct HyperlinkSpan {
    pub range: std::ops::Range<usize>,
    pub url: String,
}

/// Convert special Word characters to readable text.
///
/// A `.doc` field ([MS-DOC] §2.8.24 — `0x13` begin, `0x14` separator,
/// `0x15` end) has BOTH its instruction text (`HYPERLINK "url" \o "tip"`,
/// `DATE \@ "..."`, …) and its cached result between the boundary markers.
/// Only the cached result — the part between `0x14` and `0x15` — is what
/// Word itself displays; the instruction text must be dropped entirely,
/// not just the three boundary characters around it (issue #249).
pub fn sanitize_text(text: &str) -> String {
    strip_fields(text).0
}

/// As [`sanitize_text`], but also returns each `HYPERLINK` field's display
/// text as a [`HyperlinkSpan`] (byte range in the *returned* string) paired
/// with its target URL, so a caller building structured inline content can
/// attach `TextSpan::hyperlink` to exactly that span (issue #249).
pub fn sanitize_text_with_hyperlinks(text: &str) -> (String, Vec<HyperlinkSpan>) {
    strip_fields(text)
}

/// Shared implementation for [`sanitize_text`] / [`sanitize_text_with_hyperlinks`].
///
/// A `depth` counter tracks field nesting (a field's instruction can itself
/// contain another field, e.g. `{ IF {PAGE} > 1 "yes" "no" }`) so an inner
/// field's own `0x14`/`0x15` never gets mistaken for the outer field's.
/// Only the OUTERMOST field's own separator/end are meaningful here: its
/// instruction text (which may itself contain nested fields) is dropped in
/// full, and its cached result becomes visible text.
fn strip_fields(text: &str) -> (String, Vec<HyperlinkSpan>) {
    let mut out = String::with_capacity(text.len());
    let mut hyperlinks = Vec::new();

    let mut depth: u32 = 0;
    let mut in_result = false; // past the outermost field's own 0x14
    let mut instruction = String::new(); // outermost field's instruction text
    let mut result_start: usize = 0; // byte offset in `out` where the result began

    for ch in text.chars() {
        match ch {
            '\x13' => {
                depth += 1;
                if depth == 1 {
                    instruction.clear();
                }
            },
            // A stray separator (depth == 0) or a nested field's own
            // separator (depth > 1) is never a boundary that matters here
            // — dropped either way, same as every other field control char.
            '\x14' => {
                if depth == 1 {
                    in_result = true;
                    result_start = out.len();
                }
            },
            '\x15' => {
                if depth >= 1 {
                    if depth == 1 {
                        if in_result {
                            if let Some(url) = parse_hyperlink_url(&instruction) {
                                hyperlinks.push(HyperlinkSpan {
                                    range: result_start..out.len(),
                                    url,
                                });
                            }
                        }
                        in_result = false;
                    }
                    depth -= 1;
                }
            },
            '\x01' | '\x08' => {}, // Picture placeholder, historic field-mark — always skip
            _ => {
                if depth == 0 {
                    push_mapped(ch, &mut out);
                } else if depth == 1 && !in_result {
                    instruction.push(ch); // outermost instruction text
                } else if depth == 1 && in_result {
                    push_mapped(ch, &mut out); // outermost cached result
                }
                // depth > 1: nested field's own instruction/result — never visible.
            },
        }
    }

    (out, hyperlinks)
}

/// Apply `sanitize_text`'s non-field control-character mappings to a single
/// character and push the result onto `out`.
fn push_mapped(ch: char, out: &mut String) {
    match ch {
        '\r' => out.push('\n'),   // Paragraph mark
        '\x07' => out.push('\t'), // Cell/row mark → tab
        '\x0C' => out.push('\n'), // Page break / section break
        '\x0B' => out.push('\n'), // Vertical tab → newline
        _ => out.push(ch),
    }
}

/// Parse a `HYPERLINK` field's instruction text for its target URL.
///
/// Handles the common external-URL shape (`HYPERLINK "http://..."`) and the
/// internal-bookmark shape (`HYPERLINK \l "bookmark"`, surfaced as
/// `#bookmark`). Returns `None` for any other field type or a `HYPERLINK`
/// field whose instruction has no quoted argument at all.
fn parse_hyperlink_url(instruction: &str) -> Option<String> {
    let trimmed = instruction.trim_start();
    let after_kw = trimmed.strip_prefix("HYPERLINK").or_else(|| {
        // Word's own writer always uppercases the keyword, but tolerate a
        // lowercase one rather than silently missing a real hyperlink.
        let first_word_len = trimmed.find(char::is_whitespace).unwrap_or(trimmed.len());
        let (first, rest) = trimmed.split_at(first_word_len);
        first.eq_ignore_ascii_case("HYPERLINK").then_some(rest)
    })?;

    let quote_start = after_kw.find('"')?;
    let after_quote = &after_kw[quote_start + 1..];
    let quote_end = after_quote.find('"')?;
    let target = &after_quote[..quote_end];

    let before_quote = after_kw[..quote_start].trim_end();
    if before_quote.ends_with("\\l") {
        Some(format!("#{target}"))
    } else {
        Some(target.to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_clx_with_one_piece() {
        let mut clx = Vec::new();
        // Pcdt marker.
        clx.push(0x02);
        // Size of PlcPcd: (1+1)*4 + 1*8 = 16
        clx.extend_from_slice(&16u32.to_le_bytes());
        // CP[0] = 0
        clx.extend_from_slice(&0u32.to_le_bytes());
        // CP[1] = 10
        clx.extend_from_slice(&10u32.to_le_bytes());
        // PCD: [u16 unused=0][u32 fc=0x40000100 (compressed, offset=0x80)][u16 prm=0]
        clx.extend_from_slice(&0u16.to_le_bytes());
        clx.extend_from_slice(&0x40000100u32.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());

        let pieces = parse_clx(&clx).unwrap();
        assert_eq!(pieces.len(), 1);
        assert_eq!(pieces[0].cp_start, 0);
        assert_eq!(pieces[0].cp_end, 10);
        assert!(pieces[0].is_compressed);
    }

    #[test]
    fn parse_clx_with_grpprl_prefix() {
        let mut clx = Vec::new();
        // Grpprl: type=0x01, size=3, data=[0,0,0]
        clx.push(0x01);
        clx.extend_from_slice(&3u16.to_le_bytes());
        clx.extend_from_slice(&[0, 0, 0]);
        // Pcdt
        clx.push(0x02);
        clx.extend_from_slice(&16u32.to_le_bytes());
        clx.extend_from_slice(&0u32.to_le_bytes());
        clx.extend_from_slice(&5u32.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());
        clx.extend_from_slice(&0x40000000u32.to_le_bytes());
        clx.extend_from_slice(&0u16.to_le_bytes());

        let pieces = parse_clx(&clx).unwrap();
        assert_eq!(pieces.len(), 1);
        assert_eq!(pieces[0].cp_end, 5);
    }

    #[test]
    fn extract_compressed_text() {
        // Build a word_doc with "Hello" at byte offset 0x80 (fc=0x40000100, offset = 0x100/2 = 0x80)
        let mut word_doc = vec![0u8; 256];
        let text_offset = 0x80;
        word_doc[text_offset..text_offset + 5].copy_from_slice(b"Hello");

        let pieces = vec![Piece {
            cp_start: 0,
            cp_end: 5,
            fc: 0x40000100, // compressed, offset = 0x100/2 = 0x80
            is_compressed: true,
        }];

        let text = extract_text(&word_doc, &pieces, 5, 0);
        assert_eq!(text, "Hello");
    }

    #[test]
    fn extract_unicode_text() {
        let mut word_doc = vec![0u8; 256];
        let fc = 100u32;
        // "Hi" in UTF-16LE at offset 100
        word_doc[100] = b'H';
        word_doc[101] = 0;
        word_doc[102] = b'i';
        word_doc[103] = 0;

        let pieces = vec![Piece {
            cp_start: 0,
            cp_end: 2,
            fc,
            is_compressed: false,
        }];

        let text = extract_text(&word_doc, &pieces, 2, 0);
        assert_eq!(text, "Hi");
    }

    #[test]
    fn extract_multiple_pieces() {
        let mut word_doc = vec![0u8; 512];
        // Piece 1: compressed "AB" at offset 0x80
        word_doc[0x80] = b'A';
        word_doc[0x81] = b'B';
        // Piece 2: compressed "CD" at offset 0x90
        word_doc[0x90] = b'C';
        word_doc[0x91] = b'D';

        let pieces = vec![
            Piece {
                cp_start: 0,
                cp_end: 2,
                fc: 0x40000100, // offset = 0x80
                is_compressed: true,
            },
            Piece {
                cp_start: 2,
                cp_end: 4,
                fc: 0x40000120, // offset = 0x90
                is_compressed: true,
            },
        ];

        let text = extract_text(&word_doc, &pieces, 4, 0);
        assert_eq!(text, "ABCD");
    }

    #[test]
    fn sanitize_paragraph_marks() {
        assert_eq!(sanitize_text("Hello\rWorld"), "Hello\nWorld");
        assert_eq!(sanitize_text("A\x0CB"), "A\nB");
        assert_eq!(sanitize_text("A\x07B"), "A\tB");
    }

    /// issue #249 — the instruction text between `0x13` and `0x14` is
    /// never visible in Word; only the cached result (between `0x14` and
    /// `0x15`) is. The old behaviour kept both, mashed together.
    #[test]
    fn sanitize_field_codes_stripped() {
        assert_eq!(sanitize_text("before\x13FIELD\x14result\x15after"), "beforeresultafter");
    }

    /// The real corpus shape from `hyperlink.doc` (issue #249): a
    /// `HYPERLINK` field's instruction and quoted URL must vanish from
    /// visible text, its cached display text must survive, and the URL
    /// must be recovered as a `HyperlinkSpan` over exactly that text.
    #[test]
    fn hyperlink_field_strips_instruction_and_yields_url_span() {
        let raw = "Before text; \x13 HYPERLINK \"http://testuri.org/\" \x14Hyperlink text\x15; after text";
        let (text, links) = sanitize_text_with_hyperlinks(raw);
        assert_eq!(text, "Before text; Hyperlink text; after text");
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].url, "http://testuri.org/");
        assert_eq!(&text[links[0].range.clone()], "Hyperlink text");
    }

    /// A `\l` switch targets an internal bookmark, not an external URL —
    /// surfaced as a `#bookmark`-shaped target the way a browser-style
    /// consumer would expect.
    #[test]
    fn hyperlink_field_with_bookmark_switch_gets_hash_prefix() {
        let raw = "\x13 HYPERLINK \\l \"SectionTwo\" \x14Jump to Section Two\x15";
        let (text, links) = sanitize_text_with_hyperlinks(raw);
        assert_eq!(text, "Jump to Section Two");
        assert_eq!(links[0].url, "#SectionTwo");
    }

    /// A non-`HYPERLINK` field (e.g. `CREATEDATE`) must still have its
    /// instruction text stripped, but must never produce a hyperlink span.
    #[test]
    fn non_hyperlink_field_produces_no_hyperlink_span() {
        let raw = "\x13 CREATEDATE   \\* MERGEFORMAT \x1419/11/2010 14:49:00\x15";
        let (text, links) = sanitize_text_with_hyperlinks(raw);
        assert_eq!(text, "19/11/2010 14:49:00");
        assert!(links.is_empty(), "a CREATEDATE field must never yield a hyperlink span");
    }

    /// A field nested inside another field's instruction (e.g. `{ IF
    /// {PAGE} > 1 "yes" "no" }`) must not have its own `0x14`/`0x15`
    /// mistaken for the outer field's boundary — the whole nested field is
    /// swallowed as part of the outer instruction, and only the outer
    /// field's own cached result becomes visible.
    #[test]
    fn nested_field_does_not_confuse_the_outer_fields_boundary() {
        let raw = "\x13 IF \x13 PAGE \x141\x15 > 1 \"yes\" \"no\" \x14no\x15";
        let text = sanitize_text(raw);
        assert_eq!(text, "no");
    }

    #[test]
    fn cp1252_special_chars() {
        assert_eq!(cp1252_to_char(0x80), '€');
        assert_eq!(cp1252_to_char(0x93), '\u{201C}');
        assert_eq!(cp1252_to_char(0x94), '\u{201D}');
        assert_eq!(cp1252_to_char(0x41), 'A');
    }

    #[test]
    fn max_chars_limits_output() {
        let mut word_doc = vec![0u8; 256];
        word_doc[0x80..0x85].copy_from_slice(b"Hello");

        let pieces = vec![Piece {
            cp_start: 0,
            cp_end: 5,
            fc: 0x40000100,
            is_compressed: true,
        }];

        let text = extract_text(&word_doc, &pieces, 3, 0);
        assert_eq!(text, "Hel");
    }

    // --------------------------------------------------------------------
    // Regression tests for malformed-input robustness defects (AGENTS.md rule 6:
    // no panic / overflow / hang on untrusted input). Each fixture is a minimal
    // synthetic piece table, no third-party document.
    // --------------------------------------------------------------------

    /// A non-monotonic CP array (`CP[1] < CP[0]`) must surface as `Err`, not
    /// trigger a subtraction-with-overflow panic when the range is later
    /// walked. `parse_plc_pcd` validates CP ordering (piece_table.rs:119) and
    /// returns `Err` on a reversed range, so the later subtraction is always
    /// on a well-ordered piece.
    #[test]
    fn parse_clx_rejects_nonmonotonic_cps() {
        // CLX: Pcdt marker, PlcPcd size 16, CP[0]=10, CP[1]=5 (reversed), one PCD.
        let mut clx = Vec::new();
        clx.push(0x02); // Pcdt marker
        clx.extend_from_slice(&16u32.to_le_bytes()); // PlcPcd size = (n+1)*4 + n*8, n=1
        clx.extend_from_slice(&10u32.to_le_bytes()); // CP[0]
        clx.extend_from_slice(&5u32.to_le_bytes()); // CP[1]  (CP[1] < CP[0]!)
        clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: unused u16
        clx.extend_from_slice(&0u32.to_le_bytes()); // PCD: fc
        clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: prm

        let pieces = parse_clx(&clx);
        assert!(
            pieces.is_err(),
            "non-monotonic CP array must surface as Err, not panic on subtraction"
        );
    }

    /// The backing clamp must bound the *work* (the CP range actually walked),
    /// not merely the output. The decoded text is identical with or without the
    /// clamp (the loop has its own bounds checks), so the output alone cannot
    /// catch a regression. These assertions pin `piece_backed_cp_end` directly —
    /// that value changes the moment the clamp is dropped.
    #[test]
    fn decode_cp_range_backing_is_clamped_not_just_output() {
        let word_doc = vec![0u8; 64];

        // Fully unbacked: declared range starts at fc=4096 but the stream is
        // only 64 bytes, so nothing is backed → backed end collapses to cp_start.
        let unbacked = Piece {
            cp_start: 0,
            cp_end: 20_000_000,
            fc: 4096,
            is_compressed: false,
        };
        assert_eq!(
            piece_backed_cp_end(&unbacked, word_doc.len()),
            0,
            "fully unbacked piece must clamp to cp_start (no iteration)"
        );
        assert_eq!(
            decode_cp_range(&word_doc, &[unbacked], 0, 20_000_000, 0),
            "",
            "CP range with no backing bytes yields empty text"
        );

        // Partially backed: fc=0 with a 64-byte stream backs exactly 32 Unicode
        // chars (2 bytes each), even though the declared range spans 20M CPs.
        let backed = Piece {
            cp_start: 0,
            cp_end: 20_000_000,
            fc: 0,
            is_compressed: false,
        };
        assert_eq!(
            piece_backed_cp_end(&backed, word_doc.len()),
            32,
            "partially backed Unicode piece must clamp to 32, not 20M"
        );
        assert_eq!(
            decode_cp_range(&word_doc, &[backed], 0, 20_000_000, 0)
                .chars()
                .count(),
            32,
            "only the 32 Unicode chars that are actually backed are decoded"
        );

        // Compressed (CP1252) piece: 1 byte per char, so 64 bytes back exactly
        // 64 CPs. The helper must use the compressed stride, not the 2-byte one.
        let compressed = Piece {
            cp_start: 0,
            cp_end: 20_000_000,
            fc: 0,
            is_compressed: true,
        };
        assert_eq!(
            piece_backed_cp_end(&compressed, word_doc.len()),
            64,
            "compressed piece must clamp to 64 CPs (1 byte each), not 32"
        );
    }

    /// `piece_backed_cp_end` must add a non-zero `cp_start` to the backed
    /// length, and must apply the compressed-piece `fc` mask
    /// (`& !0x4000_0000`, then `/ 2`) when the piece actually has an offset.
    /// The unbacked/partial tests above all use `cp_start == 0` and `fc == 0`,
    /// so neither the `saturating_add` nor the offset mask was exercised there.
    #[test]
    fn piece_backed_cp_end_nonzero_cp_start_and_offset() {
        let word_doc = [0u8; 64];

        // Unicode piece with cp_start=100: backed end = 100 + 32.
        let unicode = Piece {
            cp_start: 100,
            cp_end: 20_000_000,
            fc: 0,
            is_compressed: false,
        };
        assert_eq!(
            piece_backed_cp_end(&unicode, word_doc.len()),
            132,
            "non-zero cp_start must be added to the backed length (100 + 32)"
        );

        // Compressed piece whose base offset exceeds the stream: fc=0x40000100
        // -> base = (0x40000100 & !0x40000000)/2 = 0x80 = 128, stream is only
        // 64 bytes, so nothing is backed. cp_start=5 still applies.
        let compressed_unbacked = Piece {
            cp_start: 5,
            cp_end: 20_000_000,
            fc: 0x40000100,
            is_compressed: true,
        };
        assert_eq!(
            piece_backed_cp_end(&compressed_unbacked, word_doc.len()),
            5,
            "compressed offset beyond the stream must clamp to cp_start only"
        );

        // Compressed piece with an offset that IS backed: fc=0x40000010 ->
        // base = (0x40000010 & !0x40000000)/2 = 8; stream 64 bytes backs
        // 56 chars, plus cp_start=7 -> 63. Exercises the mask + non-zero start.
        let compressed_backed = Piece {
            cp_start: 7,
            cp_end: 20_000_000,
            fc: 0x40000010,
            is_compressed: true,
        };
        assert_eq!(
            piece_backed_cp_end(&compressed_backed, word_doc.len()),
            7 + 56,
            "compressed backing = cp_start + (stream_len - base_offset)"
        );
    }

    /// `decode_cp_range` must decode a *mid-range* request inside a backed
    /// piece, exercising the per-cp offset `base + (cp - piece.cp_start)`. The
    /// other tests only decode from `cp_start` (offset 0), so the in-range
    /// offset arithmetic was never directly asserted.
    #[test]
    fn decode_cp_range_mid_range_into_backed_piece() {
        // "HelloWorld" in UTF-16LE (10 chars = 20 bytes) at fc=0.
        let mut word_doc = vec![0u8; 20];
        let text = b"HelloWorld";
        for (i, &b) in text.iter().enumerate() {
            word_doc[2 * i] = b;
        }
        let piece = Piece {
            cp_start: 0,
            cp_end: 10,
            fc: 0,
            is_compressed: false,
        };
        // Request only CP 4..7 -> "oWo".
        let out = decode_cp_range(&word_doc, &[piece], 4, 7, 0);
        assert_eq!(out, "oWo", "mid-range decode must use cp - cp_start offset");
    }

    /// Malformed CLX inputs must surface as `Err`, never panic (AGENTS.md rule 6).
    /// Each fixture exercises one of the truncation / wrong-marker branches in
    /// `parse_clx` that the happy-path fixtures never reach.
    #[test]
    fn parse_clx_truncation_is_err() {
        // Grpprl marker (0x01) with no size bytes at all -> "Grpprl truncated".
        assert!(parse_clx(&[0x01]).is_err());
        // Marker that is neither Grpprl nor Pcdt -> "expected Pcdt".
        assert!(parse_clx(&[0x03]).is_err());
        // Pcdt marker (0x02) with no size field -> "Pcdt size truncated".
        assert!(parse_clx(&[0x02]).is_err());
        // Pcdt present but its PlcPcd payload shorter than the minimum 8 bytes
        // -> "PlcPcd too small".
        let mut clx = vec![0x02u8];
        clx.extend_from_slice(&8u32.to_le_bytes()); // Pcdt size = 8
        clx.extend_from_slice(&[0u8; 3]); // only 3 payload bytes follow
        assert!(parse_clx(&clx).is_err());
    }

    /// Every CP1252 special byte (0x80..=0x9F) must resolve without panic, and
    /// the documented multi-byte code points must be correct. The conversion
    /// table's individual arms are otherwise only partially exercised.
    #[test]
    fn cp1252_special_bytes_all_covered() {
        for b in 0x80u8..=0x9F {
            let _ = cp1252_to_char(b);
        }
        assert_eq!(cp1252_to_char(0x80), '€');
        assert_eq!(cp1252_to_char(0x85), '\u{2026}'); // …
        assert_eq!(cp1252_to_char(0x91), '\u{2018}'); // '
        assert_eq!(cp1252_to_char(0x92), '\u{2019}'); // '
        assert_eq!(cp1252_to_char(0x93), '\u{201C}'); // "
        assert_eq!(cp1252_to_char(0x94), '\u{201D}'); // "
        assert_eq!(cp1252_to_char(0x95), '\u{2022}'); // •
        assert_eq!(cp1252_to_char(0x96), '\u{2013}'); // –
        assert_eq!(cp1252_to_char(0x97), '\u{2014}'); // —
        assert_eq!(cp1252_to_char(0x99), '\u{2122}'); // ™
    }

    /// `sanitize_text` must map every classified control character to its
    /// documented replacement, including the field-code markers it strips.
    #[test]
    fn sanitize_all_control_marks() {
        assert_eq!(sanitize_text("A\x01B"), "AB"); // picture placeholder stripped
        assert_eq!(sanitize_text("A\x08B"), "AB"); // historic field-mark stripped
        // An unclosed field (issue #249): "B" is instruction text with no
        // matching separator/end, so it's never promoted to visible output.
        assert_eq!(sanitize_text("A\x13B"), "A");
        assert_eq!(sanitize_text("A\x14B"), "AB"); // stray separator (no open field): dropped, "B" is plain text
        assert_eq!(sanitize_text("A\x15B"), "AB"); // stray end (no open field): dropped, "B" is plain text
        assert_eq!(sanitize_text("A\x0BB"), "A\nB"); // vertical tab -> newline
    }
}

#[cfg(test)]
mod multi_piece_tests {
    use super::*;

    /// Build a CLX holding `n` pieces over a Unicode text buffer.
    fn clx(pieces: &[(u32, u32, u32)]) -> Vec<u8> {
        let mut plc = Vec::new();
        for (cp, _, _) in pieces {
            plc.extend_from_slice(&cp.to_le_bytes());
        }
        plc.extend_from_slice(&pieces.last().unwrap().1.to_le_bytes());
        for (_, _, fc) in pieces {
            plc.extend_from_slice(&0u16.to_le_bytes());
            plc.extend_from_slice(&fc.to_le_bytes());
            plc.extend_from_slice(&0u16.to_le_bytes());
        }
        let mut out = vec![0x02];
        out.extend_from_slice(&(plc.len() as u32).to_le_bytes());
        out.extend_from_slice(&plc);
        out
    }

    /// A document acquires multiple pieces through ordinary editing
    /// history, so this is a common shape rather than an exotic one.
    /// `lcbClx = 45` is exactly the three-piece size from issue #168.
    #[test]
    fn a_three_piece_table_extracts_every_piece() {
        let text: Vec<u8> = "ABCDEF"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        let mut word_doc = vec![0u8; 512];
        word_doc.extend_from_slice(&text);
        let base = 512u32;
        let c = clx(&[(0, 2, base), (2, 4, base + 4), (4, 6, base + 8)]);
        assert_eq!(c.len(), 45, "the three-piece CLX size from the report");

        let pieces = parse_clx(&c).expect("parse");
        assert_eq!(pieces.len(), 3);
        assert_eq!(extract_text(&word_doc, &pieces, 6, 0), "ABCDEF");
    }

    /// Ranges must be clipped at both ends, which is what lets the
    /// subdocuments be read out of the same character space.
    #[test]
    fn a_range_starting_mid_piece_is_clipped_at_both_ends() {
        let text: Vec<u8> = "ABCDEF"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        let mut word_doc = vec![0u8; 512];
        word_doc.extend_from_slice(&text);
        let c = clx(&[(0, 6, 512)]);
        let pieces = parse_clx(&c).expect("parse");
        assert_eq!(extract_text_range(&word_doc, &pieces, 2, 5, 0), "CDE");
    }

    // ── Piece table coverage vs the FIB's declared text length (#230) ──────

    fn piece(cp_start: u32, cp_end: u32) -> Piece {
        Piece { cp_start, cp_end, fc: 0, is_compressed: true }
    }

    /// A single piece that covers the whole declared range.
    #[test]
    fn covers_declared_length_true_for_a_full_single_piece() {
        assert!(covers_declared_length(&[piece(0, 100)], 100));
        // Covering more than declared is fine too.
        assert!(covers_declared_length(&[piece(0, 200)], 100));
    }

    /// Several contiguous pieces that together reach the declared length.
    #[test]
    fn covers_declared_length_true_for_contiguous_pieces() {
        assert!(covers_declared_length(&[piece(0, 40), piece(40, 70), piece(70, 100)], 100));
    }

    /// The exact real-world shape from the issue: a single piece that
    /// starts well past CP 0 and doesn't reach `text_len` — the piece
    /// table itself is internally consistent (one valid piece), but it
    /// leaves the first 2816 of 3390 declared characters completely
    /// unmapped.
    #[test]
    fn covers_declared_length_false_when_the_only_piece_starts_past_zero() {
        assert!(!covers_declared_length(&[piece(2816, 3390)], 3368));
    }

    /// A gap between two otherwise-valid pieces.
    #[test]
    fn covers_declared_length_false_for_a_gap_between_pieces() {
        assert!(!covers_declared_length(&[piece(0, 40), piece(50, 100)], 100));
    }

    /// Pieces that stop short of the declared length with no gap before
    /// that point.
    #[test]
    fn covers_declared_length_false_when_pieces_stop_short() {
        assert!(!covers_declared_length(&[piece(0, 40)], 100));
    }

    /// No pieces at all, but the FIB declares text — an empty piece table
    /// with a nonzero `ccpText` is definitionally a gap, not "no text".
    #[test]
    fn covers_declared_length_false_for_no_pieces_with_nonzero_text_len() {
        assert!(!covers_declared_length(&[], 100));
    }

    /// `text_len == 0` trivially has nothing to cover.
    #[test]
    fn covers_declared_length_true_when_text_len_is_zero() {
        assert!(covers_declared_length(&[], 0));
        assert!(covers_declared_length(&[piece(0, 10)], 0));
    }
}
