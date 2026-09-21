//! Word 6.0 / Word 95 text extraction.
//!
//! The Word 6/95 FIB shares its first 0x18 bytes with Word 97's
//! `FibBase` (`wIdent`, `nFib`, `lid`, the flags, `fcMin`/`fcMac`) and
//! then follows its own fixed layout: the `ccp*` story lengths from 0x34
//! and the FC/LCB pairs from 0x58, with the piece-table pointer
//! `fcClx`/`lcbClx` at 0x160/0x164 ([Word 6.0 Binary File Format], FIB).
//! There is no table stream — the CLX, when the file is complex
//! (fast-saved), lives in the `WordDocument` stream itself — and every
//! character is one byte in the document's code page. A non-complex
//! file's text is simply the bytes from `fcMin`, story after story.
//!
//! This is the shape of Apache POI's `HWPFOldDocument`: enough to return
//! the text every other reader (catdoc, antiword, LibreOffice, POI)
//! returns for these files, with no formatting fidelity attempted.

use super::error::{DocError, Result};
use super::piece_table::{Piece, extract_text_range, parse_clx, sanitize_text};

/// The `wIdent` magics of the family: Word 6.0/95 for Windows and the
/// three Macintosh Word 6/95 variants.
pub(super) fn is_word6_magic(wident: u16) -> bool {
    matches!(wident, 0xA5DC | 0xA697..=0xA699)
}

/// Word for Windows 1.x (`0xA59B`) and 2.0 (`0xA5DB`). These are not
/// compound files: the FIB is at byte 0 of a flat file, laid out like
/// Word 6's up to the story lengths (`fcMin` 0x18, `fcMac` 0x1C, `ccpText`
/// … `ccpAtn` from 0x34), and the text is one byte per character in the
/// Windows code page from `fcMin`. catdoc and antiword read them; a
/// `.doc` archive from the early 1990s is nothing but these.
pub(super) fn is_word2_magic(wident: u16) -> bool {
    matches!(wident, 0xA59B | 0xA5DB)
}

/// A Word 1.x/2.0 FIB: the Word 6 fields it shares, with the flags this
/// format does not have (`fExtChar`) cleared and the CLX pointer unused —
/// a fast-saved Word 2 file's piece table has its own layout this reader
/// does not follow, so its text is read contiguously and reported as
/// possibly incomplete.
pub(super) fn parse_word2_fib(data: &[u8]) -> Result<Word6Fib> {
    if data.len() < 0x180 {
        return Err(DocError::InvalidFib(format!(
            "file too short for a Word 2.0 FIB: {} bytes",
            data.len()
        )));
    }
    let mut padded;
    let data = if data.len() < FIB_LEN {
        padded = data.to_vec();
        padded.resize(FIB_LEN, 0);
        padded.as_slice()
    } else {
        data
    };
    let mut fib = parse_fib(data)?;
    fib.ext_char = false;
    // Only the five stories Word 2 has: main text, footnotes,
    // headers/footers, macros, annotations.
    for slot in fib.ccp.iter_mut().skip(5) {
        *slot = 0;
    }
    Ok(fib)
}

/// The FIB fields the extractor needs, bounded to the stream.
#[derive(Debug)]
pub(super) struct Word6Fib {
    pub lid: u16,
    pub complex: bool,
    /// `fExtChar`: the text is stored in two-byte code units (UTF-16LE)
    /// rather than one byte per character — the Far East and Macintosh
    /// Word 6/95 builds write this.
    pub ext_char: bool,
    pub fc_min: u32,
    pub fc_mac: u32,
    /// Story lengths in characters, in piece-table order: main text,
    /// footnotes, headers, macros, comments, endnotes, text boxes, header
    /// text boxes.
    pub ccp: [u32; 8],
    pub fc_clx: u32,
    pub lcb_clx: u32,
}

const FIB_LEN: usize = 0x168;

pub(super) fn parse_fib(data: &[u8]) -> Result<Word6Fib> {
    if data.len() < FIB_LEN {
        return Err(DocError::InvalidFib(format!(
            "WordDocument stream too short for a Word 6/95 FIB: {} bytes",
            data.len()
        )));
    }
    let u16_at = |o: usize| u16::from_le_bytes([data[o], data[o + 1]]);
    let u32_at = |o: usize| u32::from_le_bytes([data[o], data[o + 1], data[o + 2], data[o + 3]]);
    let flags = u16_at(0x0A);
    // Word 6 FIB flags: fDot 0x0001, fGlsy 0x0002, fComplex 0x0004,
    // fHasPic 0x0008, cQuickSaves 0x00F0, fEncrypted 0x0100,
    // fExtChar 0x1000.
    if flags & 0x0100 != 0 {
        return Err(DocError::Encrypted);
    }
    let mut ccp = [0u32; 8];
    for (i, slot) in ccp.iter_mut().enumerate() {
        *slot = u32_at(0x34 + i * 4);
    }
    Ok(Word6Fib {
        lid: u16_at(0x06),
        complex: flags & 0x0004 != 0,
        ext_char: flags & 0x1000 != 0,
        fc_min: u32_at(0x18),
        fc_mac: u32_at(0x1C),
        ccp,
        fc_clx: u32_at(0x160),
        lcb_clx: u32_at(0x164),
    })
}

/// Encode a byte offset the way [`extract_text_range`] reads a compressed
/// (8-bit) piece: `fc` carries `byte_offset * 2` with bit 30 set.
fn compressed_fc(byte_offset: u32) -> Option<u32> {
    byte_offset.checked_mul(2).map(|v| v | 0x4000_0000)
}

/// The piece table for the whole character space: every piece 8-bit, or
/// every piece UTF-16 when the FIB says `fExtChar`.
pub(super) fn pieces(word_doc: &[u8], fib: &Word6Fib) -> Result<Vec<Piece>> {
    let stream_len = word_doc.len() as u32;
    if fib.complex && fib.lcb_clx != 0 {
        let start = fib.fc_clx as usize;
        let end = start.saturating_add(fib.lcb_clx as usize);
        if start >= word_doc.len() || end > word_doc.len() {
            return Err(DocError::InvalidPieceTable(format!(
                "CLX at offset {start} size {} is outside the {}-byte WordDocument stream",
                fib.lcb_clx,
                word_doc.len()
            )));
        }
        let mut pieces = parse_clx(&word_doc[start..end])?;
        // A Word 6/95 PCD's `fc` is a plain byte offset — there is no
        // `fCompressed` bit in this format; `fExtChar` decides the width
        // for the whole file — so re-express each one in the encoding the
        // shared extractor reads. A piece that points past the stream
        // contributes nothing.
        for piece in &mut pieces {
            let fc = if piece.is_compressed {
                (piece.fc & !0x4000_0000) / 2
            } else {
                piece.fc
            };
            let fc = fc.min(stream_len);
            if fib.ext_char {
                piece.fc = fc;
                piece.is_compressed = false;
            } else {
                piece.fc = compressed_fc(fc).ok_or_else(|| {
                    DocError::InvalidPieceTable(format!("piece offset {fc} is out of range"))
                })?;
                piece.is_compressed = true;
            }
        }
        return Ok(pieces);
    }
    Ok(vec![contiguous_piece(word_doc, fib)])
}

/// One piece covering the stories laid out contiguously from `fcMin`,
/// as a non-complex (not fast-saved) file stores them.
pub(super) fn contiguous_piece(word_doc: &[u8], fib: &Word6Fib) -> Piece {
    let stream_len = word_doc.len() as u32;
    let unit = if fib.ext_char { 2 } else { 1 };
    let (fc_min, fc_mac) = (fib.fc_min.min(stream_len), fib.fc_mac.min(stream_len));
    let available = fc_mac.saturating_sub(fc_min) / unit;
    let declared = fib
        .ccp
        .iter()
        .try_fold(0u32, |a, &c| a.checked_add(c))
        .unwrap_or(u32::MAX);
    let cp_end = declared.min(available);
    if fib.ext_char {
        Piece {
            cp_start: 0,
            cp_end,
            fc: fc_min,
            is_compressed: false,
        }
    } else {
        Piece {
            cp_start: 0,
            cp_end,
            fc: compressed_fc(fc_min).unwrap_or(0x4000_0000),
            is_compressed: true,
        }
    }
}

/// The main text and each non-empty subdocument, sanitised.
pub(super) fn stories(
    word_doc: &[u8],
    fib: &Word6Fib,
    pieces: &[Piece],
) -> (String, Vec<(usize, String)>) {
    // The character space cannot exceed what the pieces cover.
    let covered = pieces.iter().map(|p| p.cp_end).max().unwrap_or(0);
    let mut cp = 0u32;
    let mut main = String::new();
    let mut subs = Vec::new();
    for (i, &len) in fib.ccp.iter().enumerate() {
        if len == 0 || cp >= covered {
            continue;
        }
        let end = cp.saturating_add(len).min(covered);
        let raw = extract_text_range(word_doc, pieces, cp, end, fib.lid);
        let text = sanitize_text(&raw);
        if i == 0 {
            main = text;
        } else if !text.trim().is_empty() {
            subs.push((i, text));
        }
        cp = end;
    }
    (main, subs)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fib_bytes(
        complex: bool,
        ccp_text: u32,
        fc_min: u32,
        fc_mac: u32,
        clx: (u32, u32),
    ) -> Vec<u8> {
        let mut d = vec![0u8; 0x300];
        d[0..2].copy_from_slice(&0xA5DCu16.to_le_bytes());
        d[2..4].copy_from_slice(&101u16.to_le_bytes());
        d[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        d[0x0A..0x0C].copy_from_slice(&(if complex { 0x0004u16 } else { 0 }).to_le_bytes());
        d[0x18..0x1C].copy_from_slice(&fc_min.to_le_bytes());
        d[0x1C..0x20].copy_from_slice(&fc_mac.to_le_bytes());
        d[0x34..0x38].copy_from_slice(&ccp_text.to_le_bytes());
        d[0x160..0x164].copy_from_slice(&clx.0.to_le_bytes());
        d[0x164..0x168].copy_from_slice(&clx.1.to_le_bytes());
        d
    }

    #[test]
    fn test_non_complex_text_is_read_from_fc_min() {
        let text = b"Hello from Word 6.\rSecond paragraph.\r";
        let mut doc = fib_bytes(false, text.len() as u32, 0x300, 0x300 + text.len() as u32, (0, 0));
        doc.extend_from_slice(text);
        let fib = parse_fib(&doc).unwrap();
        let table = pieces(&doc, &fib).unwrap();
        let (main, subs) = stories(&doc, &fib, &table);
        assert_eq!(main, "Hello from Word 6.\nSecond paragraph.\n");
        assert!(subs.is_empty());
    }

    /// A fast-saved file keeps its text in pieces named by a CLX inside
    /// the WordDocument stream; the pieces' `fc` are plain byte offsets.
    #[test]
    fn test_complex_text_follows_the_piece_table_in_the_main_stream() {
        let a = b"second half\r";
        let b = b"first half, ";
        let fc_a = 0x300u32;
        let fc_b = fc_a + a.len() as u32;
        let clx_at = fc_b + b.len() as u32;
        // CLX: clxt=2, lcb, PlcPcd { cps: [0, 12, 24], pcds: 2 × 8 bytes }.
        let mut clx = vec![2u8];
        let plc_len = 3 * 4 + 2 * 8;
        clx.extend_from_slice(&(plc_len as u32).to_le_bytes());
        for cp in [0u32, b.len() as u32, (a.len() + b.len()) as u32] {
            clx.extend_from_slice(&cp.to_le_bytes());
        }
        for fc in [fc_b, fc_a] {
            clx.extend_from_slice(&0u16.to_le_bytes());
            clx.extend_from_slice(&fc.to_le_bytes());
            clx.extend_from_slice(&0u16.to_le_bytes());
        }
        let total = (a.len() + b.len()) as u32;
        let mut doc = fib_bytes(true, total, fc_a, clx_at, (clx_at, clx.len() as u32));
        doc.extend_from_slice(a);
        doc.extend_from_slice(b);
        doc.extend_from_slice(&clx);
        let fib = parse_fib(&doc).unwrap();
        let table = pieces(&doc, &fib).unwrap();
        let (main, _) = stories(&doc, &fib, &table);
        assert_eq!(main, "first half, second half\n");
    }

    /// The Macintosh and Far East builds set `fExtChar` and store the
    /// text in two-byte units.
    #[test]
    fn test_ext_char_text_is_utf16() {
        let text: Vec<u8> = "Mac Word\r"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let mut doc = fib_bytes(false, 9, 0x300, 0x300 + text.len() as u32, (0, 0));
        doc[0x0A..0x0C].copy_from_slice(&0x1000u16.to_le_bytes());
        doc.extend_from_slice(&text);
        let fib = parse_fib(&doc).unwrap();
        assert!(fib.ext_char);
        let table = pieces(&doc, &fib).unwrap();
        assert_eq!(stories(&doc, &fib, &table).0, "Mac Word\n");
    }

    /// A Word 2.0 file is the FIB and the text with no container: the
    /// text runs from `fcMin` in the Windows code page, the five story
    /// lengths sit where Word 6 keeps its eight.
    #[test]
    fn test_word_2_text_is_read_from_a_flat_file() {
        let text = b"Word 2.0 for Windows.\rZweite Zeile: \xfcber.\r";
        let footnote = b"A footnote.\r";
        let mut doc = vec![0u8; 0x180];
        doc[0..2].copy_from_slice(&0xA5DBu16.to_le_bytes());
        doc[2..4].copy_from_slice(&45u16.to_le_bytes());
        doc[6..8].copy_from_slice(&0x0407u16.to_le_bytes());
        doc[0x18..0x1C].copy_from_slice(&0x180u32.to_le_bytes());
        let fc_mac = 0x180 + (text.len() + footnote.len()) as u32;
        doc[0x1C..0x20].copy_from_slice(&fc_mac.to_le_bytes());
        doc[0x34..0x38].copy_from_slice(&(text.len() as u32).to_le_bytes());
        doc[0x38..0x3C].copy_from_slice(&(footnote.len() as u32).to_le_bytes());
        doc.extend_from_slice(text);
        doc.extend_from_slice(footnote);
        let fib = parse_word2_fib(&doc).unwrap();
        assert!(!fib.ext_char);
        let table = vec![contiguous_piece(&doc, &fib)];
        let (main, subs) = stories(&doc, &fib, &table);
        assert_eq!(main, "Word 2.0 for Windows.\nZweite Zeile: über.\n");
        assert_eq!(subs, vec![(1, "A footnote.\n".to_string())]);
        assert!(parse_word2_fib(&doc[..0x100]).is_err(), "shorter than a FIB");
    }

    /// Fuzzed corpus files declare 1.8 GB of text and CLX offsets past
    /// the end of the stream; nothing may allocate or index on those.
    #[test]
    fn test_absurd_fib_values_are_bounded_or_refused() {
        let mut doc = fib_bytes(false, u32::MAX, 768, 0, (0, 0));
        doc.extend_from_slice(b"x");
        let fib = parse_fib(&doc).unwrap();
        let table = pieces(&doc, &fib).unwrap();
        assert_eq!(stories(&doc, &fib, &table).0, "");
        let doc = fib_bytes(true, 100, 768, 900, (4_076_008_178, 15_728_882));
        let fib = parse_fib(&doc).unwrap();
        assert!(pieces(&doc, &fib).is_err());
    }
}
