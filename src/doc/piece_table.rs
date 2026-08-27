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
pub fn extract_text(word_doc: &[u8], pieces: &[Piece], max_chars: u32) -> String {
    let mut text = String::new();

    for piece in pieces {
        if piece.cp_start >= max_chars {
            break;
        }

        let char_count = piece.cp_end.min(max_chars) - piece.cp_start;

        if piece.is_compressed {
            // Compressed: 1 byte per character, CP1252.
            // Actual byte offset = (fc & ~0x40000000) / 2
            let byte_offset = ((piece.fc & !0x40000000) / 2) as usize;
            let byte_count = char_count as usize;

            if byte_offset + byte_count <= word_doc.len() {
                for &b in &word_doc[byte_offset..byte_offset + byte_count] {
                    text.push(cp1252_to_char(b));
                }
            }
        } else {
            // Unicode: 2 bytes per character, UTF-16LE.
            let byte_offset = piece.fc as usize;
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
                    out.push(cp1252_to_char(word_doc[off]));
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
fn cp1252_to_char(b: u8) -> char {
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

/// Convert special Word characters to readable text.
pub fn sanitize_text(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '\r' => result.push('\n'),                        // Paragraph mark
            '\x07' => result.push('\t'),                      // Cell/row mark → tab
            '\x0C' => result.push('\n'),                      // Page break / section break
            '\x0B' => result.push('\n'),                      // Vertical tab → newline
            '\x01' | '\x08' | '\x13' | '\x14' | '\x15' => {}, // Field codes, picture, etc. — skip
            _ => result.push(ch),
        }
    }
    result
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

        let text = extract_text(&word_doc, &pieces, 5);
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

        let text = extract_text(&word_doc, &pieces, 2);
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

        let text = extract_text(&word_doc, &pieces, 4);
        assert_eq!(text, "ABCD");
    }

    #[test]
    fn sanitize_paragraph_marks() {
        assert_eq!(sanitize_text("Hello\rWorld"), "Hello\nWorld");
        assert_eq!(sanitize_text("A\x0CB"), "A\nB");
        assert_eq!(sanitize_text("A\x07B"), "A\tB");
    }

    #[test]
    fn sanitize_field_codes_stripped() {
        assert_eq!(sanitize_text("before\x13FIELD\x14result\x15after"), "beforeFIELDresultafter");
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

        let text = extract_text(&word_doc, &pieces, 3);
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
            decode_cp_range(&word_doc, &[unbacked], 0, 20_000_000),
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
            decode_cp_range(&word_doc, &[backed], 0, 20_000_000)
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
        let word_doc = vec![0u8; 64];

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
        let out = decode_cp_range(&word_doc, &[piece], 4, 7);
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
        assert_eq!(sanitize_text("A\x01B"), "AB"); // field begin stripped
        assert_eq!(sanitize_text("A\x08B"), "AB"); // field separator stripped
        assert_eq!(sanitize_text("A\x13B"), "AB"); // field begin
        assert_eq!(sanitize_text("A\x14B"), "AB"); // field separator
        assert_eq!(sanitize_text("A\x15B"), "AB"); // field end
        assert_eq!(sanitize_text("A\x0BB"), "A\nB"); // vertical tab -> newline
    }
}
