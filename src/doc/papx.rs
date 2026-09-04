//! PAPX (paragraph properties) parsing for Word binary documents.
//!
//! The paragraph properties for a Word 97-2003 document live in PAPX FKP
//! (Formatted KPara Page) pages, indexed by the PlcfBtePapx in the Table
//! stream. Each 512-byte FKP page sits in the WordDocument stream at
//! `pn * 512` and holds:
//!
//! - `rgfc[crun + 1]` — u32 file-character positions bounding each paragraph
//! - `rgbx[crun]` — 13-byte BX descriptors; byte 0 is a word offset into the
//!   page where that paragraph's PAPX lives
//! - `crun` — u8 count of paragraphs, stored at byte 511
//!
//! A PAPX is `[cw:1][istd:2][grpprl:cb-3]` where `cb = cw * 2` (including the
//! `cw` byte). When `cw == 0`, the real `cw` is the next byte (the "Word8"
//! re-read) — without this, row-terminator paragraphs appear to have no TAP.
//!
//! The `grpprl` is decoded by [`super::sprm::extract_pap_props`].

use super::piece_table::{Piece, decode_cp_range, sanitize_text};
use super::sprm::PapProps;
use super::styles::{StyleDef, heading_level_for_istd};

/// A paragraph descriptor recovered from a PAPX FKP page.
#[derive(Debug, Clone)]
pub struct FkpParagraph {
    /// Start FC (file character position) of the paragraph, inclusive.
    pub fc_start: u32,
    /// End FC of the paragraph, exclusive.
    pub fc_end: u32,
    /// The PAP `grpprl` bytes (without the `cw`/`istd` header).
    pub grpprl: Vec<u8>,
    /// The paragraph style index (`istd`) from the PAPX header, before any
    /// `sprmPStyle` (0x640A) override carried in the `grpprl`.
    pub istd: u16,
}

/// Result of decoding one PAPX: its `grpprl` and the header `istd`.
struct PapxData {
    grpprl: Vec<u8>,
    istd: u16,
}

/// A fully-resolved main-text paragraph: its raw text, the terminating
/// character, and the distilled PAP properties.
#[derive(Debug, Clone)]
pub struct DocParagraph {
    /// Paragraph text WITHOUT the terminating mark character.
    pub text: String,
    /// The terminating character (`\r` paragraph mark, `\x07` cell/row mark,
    /// `\x0C` page break, …). Drives cell/row grouping in `doc_to_ir`.
    pub terminator: char,
    /// Distilled PAP flags (`fInTable`, row-mark, list, …).
    pub props: PapProps,
}

/// Parse every PAPX FKP page referenced by the PlcfBtePapx.
///
/// `word_doc` is the WordDocument stream (where FKP pages live);
/// `table_stream` holds the PlcfBtePapx itself. Returns one `FkpParagraph`
/// per paragraph across all pages, in no particular order — callers filter
/// to the main-text range and sort by `fc_start`.
pub fn parse_papx_paragraphs(
    word_doc: &[u8],
    table_stream: &[u8],
    fc_plcf_bte_papx: u32,
    lcb_plcf_bte_papx: u32,
) -> Vec<FkpParagraph> {
    let start = fc_plcf_bte_papx as usize;
    if lcb_plcf_bte_papx < 4 || start + 4 > table_stream.len() {
        return Vec::new();
    }
    let end = (start + lcb_plcf_bte_papx as usize).min(table_stream.len());
    let plc = &table_stream[start..end];

    // PlcfBtePapx: (n+1) u32 FCs, then n u32 BTEs. n = (lcb - 4) / 8.
    let n = (plc.len().saturating_sub(4)) / 8;
    if n == 0 {
        return Vec::new();
    }
    let cp_arr = (n + 1) * 4; // size of the FC array
    if cp_arr + n * 4 > plc.len() {
        return Vec::new();
    }

    // Bound the FKP walk against a malformed PlcfBtePapx. Every BTE names a
    // 512-byte page, so there can be at most `word_doc.len() / 512` distinct
    // physical pages; clamp `n` to that, and skip any page we have already
    // visited. Without this, a hostile document could list many BTEs pointing
    // at the same (or many) pages and force repeated/again-large parsing
    // without bound (AGENTS.md rule 6: no input may hang or run away).
    let max_pages = word_doc.len() / 512;
    let n = n.min(max_pages);

    let mut out = Vec::new();
    let mut visited = std::collections::HashSet::with_capacity(n.min(64));
    for i in 0..n {
        let bte = u32::from_le_bytes([
            plc[cp_arr + i * 4],
            plc[cp_arr + i * 4 + 1],
            plc[cp_arr + i * 4 + 2],
            plc[cp_arr + i * 4 + 3],
        ]);
        // Low 22 bits are the page number; high bits are reserved.
        let pn = (bte & 0x003F_FFFF) as usize;
        if !visited.insert(pn) {
            continue; // same page referenced again — parsed once already
        }
        if let Some(page) = word_doc.get(pn * 512..pn * 512 + 512) {
            parse_fkp_page(page, &mut out);
        }
    }
    out
}

/// Parse a single 512-byte PAPX FKP page, appending `FkpParagraph`s to `out`.
fn parse_fkp_page(page: &[u8], out: &mut Vec<FkpParagraph>) {
    let crun = page[511] as usize;
    if crun == 0 || crun >= 64 {
        return;
    }

    // rgfc: crun + 1 u32 file positions.
    let mut rgfc = Vec::with_capacity(crun + 1);
    let mut pos = 0usize;
    for _ in 0..=crun {
        if pos + 4 > page.len() {
            return;
        }
        rgfc.push(u32::from_le_bytes([page[pos], page[pos + 1], page[pos + 2], page[pos + 3]]));
        pos += 4;
    }

    // rgbx: crun 13-byte BX descriptors. Byte 0 of each is the word offset
    // into the page where the PAPX lives.
    for i in 0..crun {
        let bx_off = pos + i * 13;
        if bx_off >= page.len() {
            break;
        }
        let word_off = page[bx_off] as usize;
        let fc_start = rgfc[i];
        let fc_end = rgfc[i + 1];
        let papx = if word_off == 0 {
            PapxData {
                grpprl: Vec::new(),
                istd: 0,
            }
        } else {
            extract_grpprl(page, word_off)
        };
        out.push(FkpParagraph {
            fc_start,
            fc_end,
            grpprl: papx.grpprl,
            istd: papx.istd,
        });
    }
}

/// Extract the PAP `grpprl` (and header `istd`) from a page at the given word
/// offset.
///
/// Layout: `[cw:1][istd:2][grpprl: cb-3]`, `cb = cw * 2`. The Word8 re-read
/// (`cw == 0` → use the next byte) is applied so row-terminator paragraphs,
/// which carry the full TAP, are not mistaken for empty PAPXs.
fn extract_grpprl(page: &[u8], word_off: usize) -> PapxData {
    let mut p = word_off * 2;
    if p >= page.len() {
        return PapxData {
            grpprl: Vec::new(),
            istd: 0,
        };
    }
    let mut cw = page[p] as usize;
    let reread = cw == 0;
    if reread {
        // Word8 re-read: the real cw is the following byte.
        p += 1;
        if p >= page.len() {
            return PapxData {
                grpprl: Vec::new(),
                istd: 0,
            };
        }
        cw = page[p] as usize;
    }
    let cb = cw * 2; // total PAPX bytes for the istd+grpprl block
    if cb < 3 {
        return PapxData {
            grpprl: Vec::new(),
            istd: 0,
        }; // only cw + istd, no grpprl
    }
    // `istd` is the 2 bytes immediately after `cw` (post-reread), before grpprl.
    let istd = if p + 3 <= page.len() {
        u16::from_le_bytes([page[p + 1], page[p + 2]])
    } else {
        0
    };
    let grpprl_start = p + 3; // skip cw (1) + istd (2)
    // Per MS-DOC §2.9.175 (PapxInFkp) the grpprl length differs by form:
    //  • cw != 0: grpprlInPapx *is* a GrpPrlAndIstd of `2*cw - 1` bytes, so the
    //    grpprl itself is `2*cw - 3`.
    //  • cw == 0: grpprlInPapx is `[cb':1][GrpPrlAndIstd: 2*cb']`, so the
    //    grpprl is `2*cb' - 2` — one byte longer than the non-reread form for
    //    the same value. That extra byte is the `+1` below; dropping it would
    //    truncate the trailing SPRM (e.g. the row's TAP).
    let grpprl_end = (p + cb + if reread { 1 } else { 0 }).min(page.len());
    if grpprl_start >= grpprl_end {
        return PapxData {
            grpprl: Vec::new(),
            istd,
        };
    }
    PapxData {
        grpprl: page[grpprl_start..grpprl_end].to_vec(),
        istd,
    }
}

/// Convert a character position to a real byte offset in the WordDocument
/// stream (i.e. with the compressed-encoding bit 30 stripped).
#[allow(dead_code)] // kept for API symmetry with `fc_to_cp` and future list work
pub fn cp_to_fc(cp: u32, pieces: &[Piece]) -> Option<u32> {
    for p in pieces {
        // A malformed piece (non-monotonic CP range) must not drive an
        // underflow; skip it. parse_plc_pcd also rejects these up front.
        if p.cp_end < p.cp_start {
            continue;
        }
        if cp >= p.cp_start && cp < p.cp_end {
            let (base, stride) = piece_byte_base(p);
            let off = (cp - p.cp_start) as u64 * stride as u64;
            let byte = base as u64 + off;
            if byte <= u32::MAX as u64 {
                return Some(byte as u32);
            }
        }
    }
    // Allow cp == final cp_end for "one past the end" lookups.
    if let Some(p) = pieces.last() {
        if p.cp_end < p.cp_start {
            return None;
        }
        if cp == p.cp_end {
            let (base, stride) = piece_byte_base(p);
            let off = (cp - p.cp_start) as u64 * stride as u64;
            let byte = base as u64 + off;
            if byte <= u32::MAX as u64 {
                return Some(byte as u32);
            }
        }
    }
    None
}

/// Convert a file character position (FC) to a character position (CP).
///
/// The FC is normalised to a real byte offset (bit 30 stripped when set)
/// before being matched against each piece's real byte range, so the lookup
/// works whether or not a compressed piece's FKP FCs carry bit 30.
///
/// Arithmetic is performed in `u64` with saturating ops so a malformed piece
/// (non-monotonic or huge CP range) degrades to "not in this piece" rather
/// than overflowing (AGENTS.md rule 6).
pub fn fc_to_cp(fc: u32, pieces: &[Piece]) -> Option<u32> {
    let fbyte = if fc & 0x4000_0000 != 0 {
        (fc & !0x4000_0000) / 2
    } else {
        fc
    };
    for p in pieces {
        if p.cp_end < p.cp_start {
            continue;
        }
        let (base, stride) = piece_byte_base(p);
        let base_u = base as u64;
        let stride_u = stride as u64;
        let cp_start_u = p.cp_start as u64;
        let cp_end_u = p.cp_end as u64;
        let end_u = base_u.saturating_add((cp_end_u - cp_start_u).saturating_mul(stride_u));
        let fbyte_u = fbyte as u64;
        if fbyte_u >= base_u && fbyte_u <= end_u {
            // Safe: `fbyte >= base` established above, both fit in u32.
            let off = fbyte - base;
            return Some(p.cp_start + off / stride);
        }
    }
    None
}

/// Real byte offset and stride (bytes per character) of a piece's start.
fn piece_byte_base(p: &Piece) -> (u32, u32) {
    if p.is_compressed {
        ((p.fc & !0x4000_0000) / 2, 1)
    } else {
        (p.fc, 2)
    }
}

/// Resolve a paragraph's heading level (1–[`MAX_OUTLINE_LEVEL`]), or `None`
/// when it is body text.
///
/// The precedence is one documented rule. Direct formatting overrides the style
/// in Word, so `sprmPOutlineLvl` (0x6412) settles the question whenever the
/// grpprl carries a valid one:
///
/// - present with a level of 1–[`MAX_OUTLINE_LEVEL`] → that level, and the style
///   is never consulted;
/// - present with level `0` → an explicit body-text marker: `None`, and again
///   the style is never consulted (a paragraph styled `Heading 3` but demoted to
///   body text stays body text);
/// - absent, or carrying an operand that is not a valid outline level → fall
///   back to the paragraph's style, `istd`, which the caller has already
///   resolved from the PAPX header and any `sprmPStyle` (0x640A) override.
fn resolve_heading_level(props: &PapProps, istd: u16, styles: &[StyleDef]) -> Option<u8> {
    if props.outline_lvl_explicit {
        // Settled by direct formatting: either a real level or an explicit
        // body-text marker (`None`).
        props.heading_level
    } else {
        heading_level_for_istd(styles, istd)
    }
}

/// Build the main-text paragraph list for `doc_to_ir`.
///
/// Walks the PAPX FKP paragraphs, keeps only those whose start CP falls in
/// the main document text range `[0, text_len)`, decodes each paragraph's CP
/// range directly from `word_doc` via the piece table, and distils PAP flags.
///
/// Each paragraph is decoded on its own CP range (never by indexing a flat
/// `Vec<char>` of the whole document) so that a surrogate pair in a Unicode
/// piece decodes into exactly one `char` at the right position — indexing a
/// `Vec<char>` by UTF-16 CP counts desyncs the moment an astral character
/// appears. The inner text is run through [`sanitize_text`] (which strips
/// field codes `0x13/0x14/0x15` and maps control chars); the trailing
/// character is kept raw as the `terminator` that drives cell/row grouping.
pub fn build_paragraphs(
    word_doc: &[u8],
    pieces: &[Piece],
    fkp: &[FkpParagraph],
    text_len: u32,
    styles: &[StyleDef],
) -> Vec<DocParagraph> {
    let mut keyed: Vec<(u32, &FkpParagraph)> = fkp
        .iter()
        .filter_map(|fp| fc_to_cp(fp.fc_start, pieces).map(|cp| (cp, fp)))
        .filter(|(cp, _)| *cp < text_len)
        .collect();
    keyed.sort_by_key(|(cp, _)| *cp);

    let mut out = Vec::with_capacity(keyed.len());
    for (cp_start, fp) in keyed {
        let cp_end = fc_to_cp(fp.fc_end, pieces).unwrap_or(cp_start + 1);
        let cp_end = cp_end.min(text_len);
        if cp_end <= cp_start {
            continue;
        }
        // Decode this paragraph's CP range directly; sanitise the inner text
        // but keep the trailing terminator raw (it drives table grouping).
        let decoded = decode_cp_range(word_doc, pieces, cp_start, cp_end);
        let chars: Vec<char> = decoded.chars().collect();
        if chars.is_empty() {
            continue;
        }
        let terminator = chars[chars.len() - 1];
        let content: String = sanitize_text(&chars[..chars.len() - 1].iter().collect::<String>());
        // `extract_pap_props` walks the grpprl once, decoding every PAP SPRM —
        // including `sprmPStyle` (0x640A) into `style_istd` and
        // `sprmPOutlineLvl` (0x6412) into `heading_level`.
        let mut props = super::sprm::extract_pap_props(&fp.grpprl);
        let istd = props.style_istd.unwrap_or(fp.istd);
        props.heading_level = resolve_heading_level(&props, istd, styles);
        out.push(DocParagraph {
            text: content,
            terminator,
            props,
        });
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::MAX_OUTLINE_LEVEL;
    use crate::doc::fib::Fib;
    use crate::doc::piece_table::Piece;
    use crate::doc::sprm::extract_pap_props;
    use crate::doc::styles::{StyleDef, parse_style_sheet};

    fn unicode_piece(fc: u32, cp_end: u32) -> Piece {
        Piece {
            cp_start: 0,
            cp_end,
            fc,
            is_compressed: false,
        }
    }

    #[test]
    fn cp_to_fc_and_back_unicode() {
        // table.doc: one Unicode piece, fc = 0x800, text_len = 23.
        let pieces = [unicode_piece(0x800, 23)];
        // cp 1 → byte 0x800 + 1*2 = 0x802.
        assert_eq!(cp_to_fc(1, &pieces), Some(0x802));
        // cp 0 → 0x800.
        assert_eq!(cp_to_fc(0, &pieces), Some(0x800));
        // Round-trip.
        assert_eq!(fc_to_cp(0x802, &pieces), Some(1));
        assert_eq!(fc_to_cp(0x800, &pieces), Some(0));
    }

    #[test]
    fn fc_to_cp_strips_compressed_bit() {
        // A compressed piece with bit 30 set; FC carries the same bit.
        let pieces = [Piece {
            cp_start: 0,
            cp_end: 5,
            fc: 0x4000_0010, // compressed, real offset = 0x10/2 = 8
            is_compressed: true,
        }];
        // fc with bit 30 → real byte 8 → cp 0.
        assert_eq!(fc_to_cp(0x4000_0010, &pieces), Some(0));
        // fc 9 (real byte, no bit) → cp 1.
        assert_eq!(fc_to_cp(9, &pieces), Some(1));
    }

    #[test]
    fn extract_grpprl_handles_word8_reread() {
        // Build a 512-byte page with one paragraph PAPX at word offset 247.
        // cw=6, istd=0000, grpprl = [16 24 01 49 66 01 00 00 00] (9 bytes).
        let mut page = vec![0u8; 512];
        let pstart = 247 * 2;
        let papx = [
            0x06, 0x00, 0x00, 0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00,
        ];
        page[pstart..pstart + papx.len()].copy_from_slice(&papx);
        page[511] = 1; // crun
        // rgfc[0..2]
        page[0..4].copy_from_slice(&0x800u32.to_le_bytes());
        page[4..8].copy_from_slice(&0x806u32.to_le_bytes());
        // rgbx[0].offset at byte 8
        page[8] = 247;

        let mut out = Vec::new();
        parse_fkp_page(&page, &mut out);
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].fc_start, 0x800);
        assert_eq!(out[0].fc_end, 0x806);
        assert_eq!(out[0].grpprl, vec![0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn empty_papx_when_word_off_zero() {
        let mut page = vec![0u8; 512];
        page[511] = 1;
        page[0..4].copy_from_slice(&0x800u32.to_le_bytes());
        page[4..8].copy_from_slice(&0x802u32.to_le_bytes());
        page[8] = 0; // no PAPX
        let mut out = Vec::new();
        parse_fkp_page(&page, &mut out);
        assert_eq!(out.len(), 1);
        assert!(out[0].grpprl.is_empty());
    }

    #[test]
    fn build_paragraphs_slices_text_and_flags() {
        // cp0..6: a leading '\r' mark, cell "1", cell "2", then a row mark.
        // One Unicode piece whose bytes live at fc=0x800 in `word_doc`.
        let raw = "\r1\u{7}2\u{7}\u{7}";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, 6)];

        // FKP paragraphs with FC ranges (Unicode: cp*2 + 0x800) and grpprls.
        let mk = |cp0: u32, cp1: u32, grpprl: &[u8]| FkpParagraph {
            fc_start: 0x800 + cp0 * 2,
            fc_end: 0x800 + cp1 * 2,
            grpprl: grpprl.to_vec(),
            istd: 0,
        };
        let cell = [0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00];
        let rowmark = [
            0x16, 0x24, 0x01, 0x17, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00, 0x08, 0xd6,
            0x02, 0x00, 0x00, 0x00,
        ];
        let fkp = vec![
            mk(0, 1, &[]),      // leading '\r', default props
            mk(1, 3, &cell),    // "1\u{7}"
            mk(3, 5, &cell),    // "2\u{7}"
            mk(5, 6, &rowmark), // "\u{7}" row mark
        ];

        let paras = build_paragraphs(&word_doc, &pieces, &fkp, 6, &[]);
        assert_eq!(paras.len(), 4);
        // leading mark
        assert_eq!(paras[0].text, "");
        assert_eq!(paras[0].terminator, '\r');
        assert!(!paras[0].props.f_in_table);
        // cells
        assert_eq!(paras[1].text, "1");
        assert_eq!(paras[1].terminator, '\u{7}');
        assert!(paras[1].props.f_in_table);
        assert!(!paras[1].props.is_table_trailing_mark);
        assert_eq!(paras[2].text, "2");
        // row mark
        assert_eq!(paras[3].text, "");
        assert_eq!(paras[3].terminator, '\u{7}');
        assert!(paras[3].props.f_in_table);
        assert!(paras[3].props.is_table_trailing_mark);
    }

    /// Regression: an astral character (emoji) is a single `char` but two
    /// UTF-16 code units, so a `Vec<char>` indexed by CP desyncs every later
    /// paragraph. Decoding each CP range directly must keep alignment. Each
    /// paragraph ends with a terminator (`\r`) as a real `.doc` does.
    #[test]
    fn build_paragraphs_keeps_astral_alignment() {
        // "Hi 😀\r" (cp0..6) then "there\r" (cp6..12). The emoji occupies
        // cp3 and cp4 (two UTF-16 units) but is one char.
        let raw = "Hi 😀\rthere\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, 12)];

        let mk = |cp0: u32, cp1: u32| FkpParagraph {
            fc_start: 0x800 + cp0 * 2,
            fc_end: 0x800 + cp1 * 2,
            grpprl: Vec::new(),
            istd: 0,
        };
        let fkp = vec![mk(0, 6), mk(6, 12)];

        let paras = build_paragraphs(&word_doc, &pieces, &fkp, 12, &[]);
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text, "Hi 😀", "emoji must not desync the range");
        assert_eq!(paras[0].terminator, '\r');
        assert_eq!(paras[1].text, "there", "second paragraph must be intact");
        assert_eq!(paras[1].terminator, '\r');
    }

    /// Regression: field codes (0x13/0x14/0x15) in a paragraph must be
    /// stripped from the IR text, matching the sanitised plain-text path.
    #[test]
    fn build_paragraphs_strips_field_codes() {
        // A HYPERLINK field run inside one paragraph, terminated by '\r'.
        let raw = "See\x13HYPERLINK\x14result\x15here\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];

        let mk = |cp0: u32, cp1: u32| FkpParagraph {
            fc_start: 0x800 + cp0 * 2,
            fc_end: 0x800 + cp1 * 2,
            grpprl: Vec::new(),
            istd: 0,
        };
        let fkp = vec![mk(0, raw.chars().count() as u32)];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &[]);
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].terminator, '\r');
        let t = &paras[0].text;
        assert!(!t.contains('\u{13}'), "field begin must be stripped");
        assert!(!t.contains('\u{14}'), "field separator must be stripped");
        assert!(!t.contains('\u{15}'), "field end must be stripped");
        assert!(t.contains("HYPERLINK"), "field result text survives");
        assert_eq!(t, "SeeHYPERLINKresulthere");
    }

    /// Regression: a paragraph whose PAPX `istd` points at a built-in `Heading N`
    /// style (sti 1..9) must resolve to `heading_level == Some(N)` via the style
    /// sheet, even when no `sprmPOutlineLvl` (0x6412) is present. This is the
    /// style-sheet fallback path that replaces the line heuristic for styled
    /// headings.
    #[test]
    fn build_paragraphs_resolves_heading_from_style_istd() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef::default(), // istd 2
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 3: built-in Heading 3
        ];
        let raw = "Subsection\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl: Vec::new(),
            istd: 3,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(3),
            "built-in Heading style must resolve to its level"
        );
    }

    /// Regression: a `sprmPStyle` (0x640A) carried in the grpprl must override
    /// the PAPX `istd` when resolving the style sheet. A paragraph whose PAPX
    /// header says `Normal` (istd 0) but whose grpprl re-styles it to
    /// `Heading 3` must resolve to level 3 — not `None`, and not whatever the
    /// PAPX `istd` alone would have resolved to (which here is nothing).
    #[test]
    fn build_paragraphs_prefers_sprm_style_over_papx_istd() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef::default(), // istd 2
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 3: built-in Heading 3
        ];
        let raw = "Subsection\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        // PAPX `istd` = 0 (Normal), but the grpprl carries `sprmPStyle` (0x640A,
        // spra-3 => 4-byte operand) re-pointing at istd 3. The low two operand
        // bytes are the target istd.
        let grpprl = vec![0x0A, 0x64, 0x03, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 0,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(3),
            "`sprmPStyle` override must win over the PAPX `istd`"
        );
    }

    /// Regression: a user-defined heading style (sti 0x0FFE, name carries the
    /// level) must resolve via its name, case-insensitively, through the full
    /// `build_paragraphs` pipeline. Here the style name is lowercase
    /// `"heading 3"` to prove the lookup is not case-sensitive.
    #[test]
    fn build_paragraphs_resolves_user_heading_name_case_insensitive() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef {
                sti: 0x0FFE,
                name: "heading 3".into(),
            }, // istd 2: user-defined, name carries level
        ];
        let raw = "Subsection\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl: Vec::new(),
            istd: 2,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(3),
            "user-defined `heading N` name must resolve to its level"
        );
    }

    /// Regression: `sprmPOutlineLvl` (0x6412) carried in the grpprl must set the
    /// heading level directly, with no style sheet involved. The opcode is
    /// 0x6412 (LE) + a 4-byte operand whose low byte is the outline level; here
    /// `5` must surface as Heading 5. This is the headline "outline SPRM" path
    /// and was previously only covered at the `extract_pap_props` layer — this
    /// test proves it flows through `FKP → build_paragraphs → props.heading_level`.
    #[test]
    fn build_paragraphs_resolves_heading_from_sprm_outline_lvl() {
        let styles = vec![StyleDef::default()]; // istd 0: Normal
        let raw = "Section Five\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        // grpprl = 0x6412 (LE) + 4-byte operand, low byte = 5.
        let grpprl = vec![0x12, 0x64, 0x05, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 0, // PAPX says Normal; the SPRM overrides it to Heading 5
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(5),
            "`sprmPOutlineLvl` must set the heading level directly"
        );
    }

    /// Regression: when both a heading style and `sprmPOutlineLvl` are present,
    /// the SPRM is authoritative and wins. Here the style (istd 3, built-in
    /// Heading 3) would resolve to level 3, but the grpprl's `sprmPOutlineLvl`
    /// level 5 must take precedence — proving the documented precedence
    /// (`props.heading_level` is only filled from the style when the SPRM left
    /// it `None`).
    #[test]
    fn build_paragraphs_sprm_outline_lvl_overrides_style() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef::default(), // istd 2
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 3: built-in Heading 3
        ];
        let raw = "Section Five\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        let grpprl = vec![0x12, 0x64, 0x05, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 3, // would resolve to Heading 3 on its own
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(5),
            "`sprmPOutlineLvl` must override the style-derived level"
        );
    }

    /// Build one `LPStd` for a synthetic style sheet: `cbStd(u16)` + `STD`,
    /// where `STD` = `StdfBase` (10 bytes, `sti` in its low 12 bits) +
    /// `xstzName` (`cch` + code units + 2-byte null) — MS-DOC §2.9.135,
    /// §2.9.258, §2.9.354.
    /// With a style sheet present but no heading signal on the paragraph, the
    /// level must stay `None` so the `.doc` walk falls back to the line-based
    /// heuristic. The other tests in this file mostly pass `&[]` for `styles`;
    /// this one pins the behaviour with a *populated* style sheet, which is the
    /// case a real document with a style sheet hits for its body text.
    #[test]
    fn build_paragraphs_leaves_body_text_unheaded_with_styles_present() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 1: a heading style exists in the document
        ];
        let raw = "Ordinary body prose.\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl: Vec::new(),
            istd: 0, // Normal
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level, None,
            "body text stays unheaded even when the style sheet contains heading styles, \
             so the caller falls back to the heuristic"
        );
    }

    /// The `sprmPStyle` override must work in **both** directions. The existing
    /// test covers restyling a `Normal` paragraph up to `Heading 3`; this covers
    /// restyling a `Heading 3` paragraph down to `Normal`, which must *remove*
    /// the heading rather than leaving the PAPX `istd`'s level in place.
    #[test]
    fn build_paragraphs_sprm_style_override_can_remove_a_heading() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef::default(), // istd 2
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 3: built-in Heading 3
        ];
        let raw = "Restyled to body text\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        // PAPX `istd` = 3 (Heading 3) but `sprmPStyle` (0x640A) restyles it to
        // istd 0 (Normal), so the heading must go away.
        let grpprl = vec![0x0A, 0x64, 0x00, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 3,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level, None,
            "`sprmPStyle` restyling a heading down to Normal must drop the heading"
        );
    }

    /// An explicit `sprmPOutlineLvl` of 0 marks the paragraph as **body text**.
    /// Direct formatting overrides the style in Word, so such a paragraph must
    /// not fall back to its style: a paragraph whose style is `Heading 3` but
    /// which is explicitly marked body text is body text.
    ///
    /// This pins the difference between "the SPRM is absent" (consult the
    /// style) and "the SPRM is present with level 0" (settle it as body text).
    #[test]
    fn build_paragraphs_sprm_outline_lvl_zero_suppresses_styled_heading() {
        let styles = vec![
            StyleDef::default(), // istd 0: Normal
            StyleDef::default(), // istd 1
            StyleDef::default(), // istd 2
            StyleDef {
                sti: 3,
                name: "Heading 3".into(),
            }, // istd 3: built-in Heading 3
        ];
        let raw = "Not a heading\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        // 0x6412 (LE) + 4-byte operand whose low byte is 0 (explicit body text).
        let grpprl = vec![0x12, 0x64, 0x00, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 3, // the style alone would resolve to Heading 3
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level, None,
            "an explicit outline level 0 must suppress the style-derived heading"
        );
    }

    /// MS-DOC outline levels run to `MAX_OUTLINE_LEVEL` and `sprmPOutlineLvl`
    /// accepts that whole range; the clamp to the IR's 1..=MAX_HEADING_DEPTH
    /// depth happens later, at `emit_prose`. This pins the boundary: the SPRM
    /// path must neither reject the deepest level nor pre-clamp it.
    #[test]
    fn build_paragraphs_sprm_outline_lvl_accepts_deepest_level() {
        let styles = vec![StyleDef::default()]; // istd 0: Normal
        let raw = "Deep section\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        // 0x6412 (LE) + 4-byte operand whose low byte is the deepest level.
        let grpprl = vec![0x12, 0x64, MAX_OUTLINE_LEVEL, 0x00, 0x00, 0x00];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl,
            istd: 0,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(MAX_OUTLINE_LEVEL),
            "the outline SPRM accepts the full 1..=MAX_OUTLINE_LEVEL range; clamping to \
             MAX_HEADING_DEPTH is the IR boundary's job, not this one"
        );
    }

    fn lpstd(sti: u16, name: &str) -> Vec<u8> {
        let mut std = vec![0u8; 10]; // StdfBase
        std[0..2].copy_from_slice(&sti.to_le_bytes());
        let units: Vec<u16> = name.encode_utf16().collect();
        std.extend_from_slice(&(units.len() as u16).to_le_bytes()); // cch
        for u in &units {
            std.extend_from_slice(&u.to_le_bytes());
        }
        std.extend_from_slice(&0u16.to_le_bytes()); // chTerm
        let mut out = (std.len() as u16).to_le_bytes().to_vec(); // cbStd
        out.extend_from_slice(&std);
        out
    }

    /// A spec-conformant 15-style STSH (§2.9.271): `cbStshi`(18) then `Stshif`(18),
    /// followed directly by `rglpstd`, using the fixed-index table — istd 0 is
    /// Normal (sti 0), istd 1–9 are Heading 1–9 (sti 1–9), and istd 10–14 are
    /// empty. So `istd 3` is `Heading 3`. Built from the spec, not from our
    /// parser.
    fn stsh_with_heading_3() -> Vec<u8> {
        let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
        d.extend_from_slice(&15u16.to_le_bytes()); // cstd
        d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
        d.extend_from_slice(&[0u8; 14]); // remainder of the 18-byte Stshif
        d.extend_from_slice(&lpstd(0, "Normal"));
        for lvl in 1..=9u16 {
            d.extend_from_slice(&lpstd(lvl, &format!("Heading {lvl}")));
        }
        for _ in 0..5 {
            d.extend_from_slice(&[0u8; 2]); // empty LPStd (cbStd = 0)
        }
        d
    }

    fn fib_with_stsh(start: u32, len: u32) -> Fib {
        Fib {
            version: 0,
            use_table1: false,
            clx_offset: 0,
            clx_size: 0,
            text_len: 0,
            footnote_len: 0,
            header_len: 0,
            comment_len: 0,
            endnote_len: 0,
            textbox_len: 0,
            header_textbox_len: 0,
            fc_plcf_bte_papx: 0,
            lcb_plcf_bte_papx: 0,
            fc_plcf_lst: 0,
            lcb_plcf_lst: 0,
            fc_stshf: start,
            lcb_stshf: len,
        }
    }

    /// Integration: drive the FIB-aware `parse_style_sheet` (the call
    /// `document.rs` makes) and feed its result into `build_paragraphs`, exactly
    /// as the real `.doc` walk does — but with a synthetic Table stream instead
    /// of a parsed CFB. The POI corpus cannot exercise this (every file reports
    /// `fc_stshf == 0`), so this is the regression guard that stands in for a
    /// "real .doc with a style sheet": a paragraph styled `Heading 3` (istd 3)
    /// must resolve to level 3 through the full byte-to-IR pipeline.
    #[test]
    fn build_paragraphs_resolves_heading_from_parsed_style_sheet() {
        let stsh = stsh_with_heading_3();
        let start = 64usize;
        let mut table_stream = vec![0u8; start + stsh.len()];
        table_stream[start..start + stsh.len()].copy_from_slice(&stsh);
        let fib = fib_with_stsh(start as u32, stsh.len() as u32);
        let styles = parse_style_sheet(&table_stream, &fib);
        assert_eq!(styles.len(), 15, "the 15 fixed-index LPStd entries");
        assert_eq!(styles[3].sti, 3);
        assert_eq!(styles[3].name, "Heading 3");

        let raw = "Subsection\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let pieces = [unicode_piece(0x800, raw.chars().count() as u32)];
        let fkp = [FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + raw.chars().count() as u32 * 2,
            grpprl: Vec::new(),
            istd: 3,
        }];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, &styles);
        assert_eq!(paras.len(), 1);
        assert_eq!(
            paras[0].props.heading_level,
            Some(3),
            "heading must resolve from a style sheet obtained via parse_style_sheet"
        );
    }

    /// Regression: the `cw == 0` Word8 re-read branch must not drop the trailing
    /// byte of the grpprl. Per [MS-DOC] §2.9.175 (PapxInFkp), the
    /// re-read form is `[cb':1][GrpPrlAndIstd: 2*cb']`, so the grpprl is
    /// `2*cb' - 2` bytes — one byte longer than the `2*cb - 3` non-reread
    /// form. The grpprl therefore ends at `p + 1 + 2*cb'`; the `+1` in
    /// `extract_grpprl` is what keeps that byte (the buggy `p + 2*cb'` would
    /// drop it).
    /// When the last SPRM is `sprmTDefTable`, dropping that byte truncates the
    /// TAP and the cell-merge spans vanish. The existing
    /// `extract_grpprl_handles_word8_reread` test uses `cw = 6`, so it never
    /// exercises this branch.
    #[test]
    fn papx_cw_zero_reread_extracts_trailing_tdef_table() {
        // 512-byte FKP page with the PAPX at word offset 0.
        let mut page = vec![0u8; 512];
        // PAPX: 0x00 (cw re-read marker), cw' = 17, istd = 0000, grpprl (32 bytes).
        page[0] = 0x00; // cw == 0 -> re-read the next byte as the real cw
        page[1] = 0x11; // cw' = 17 -> grpprl should be 2*17 - 2 = 32 bytes
        page[2] = 0x00;
        page[3] = 0x00; // istd
        // grpprl = sprmPFInTable(0x2416)=1  (3 bytes)
        //        + sprmTDefTable (opcode 2 + 2-byte cb=26 + 25-byte operand)
        let mut grpprl: Vec<u8> = vec![
            0x16, 0x24, 0x01, // sprmPFInTable = 1
            0x08, 0xD6, 0x1A, 0x00, // sprmTDefTable (0xD608), 2-byte cb = 26
            0x01, // itcMac = 1
            0x00, 0x00, 0x88, 0x13, // rgdxaCenter: 0, 5000
        ];
        grpprl.resize(32, 0); // rgtc padding (20 zero bytes) -> 32-byte grpprl
        assert_eq!(grpprl.len(), 32);
        page[4..4 + 32].copy_from_slice(&grpprl);

        let extracted = extract_grpprl(&page, 0);
        let props = extract_pap_props(&extracted.grpprl);
        assert!(
            props.tap.is_some(),
            "cw==0 re-read must keep the full grpprl so the trailing TAP parses"
        );
        assert!(props.is_table_trailing_mark);
    }

    /// A piece whose CP range spans nearly the whole `u32` space must not
    /// trigger an arithmetic-overflow panic in `fc_to_cp` (the
    /// `cp_end - cp_start` / `* stride` computation, papx.rs:197). Wrapped in
    /// `catch_unwind` because the failure mode is a panic; debug builds have
    /// overflow-checks enabled (AGENTS.md rule 6: no input may panic).
    #[test]
    fn fc_to_cp_huge_range_does_not_panic() {
        let pieces = [Piece {
            cp_start: 0,
            cp_end: u32::MAX,
            fc: 0,
            is_compressed: false,
        }];
        let result =
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| fc_to_cp(0, &pieces)));
        assert!(result.is_ok(), "fc_to_cp must not overflow on a huge declared CP range");
    }

    /// Regression (AGENTS.md rule 6): a PlcfBtePapx listing many BTEs that all
    /// reference the same FKP page must not cause unbounded repeated parsing.
    /// The walk dedupes visited pages and clamps `n` to the number of physical
    /// pages, so the work stays bounded even with a hostile large `lcb`.
    #[test]
    fn papx_fkp_walk_is_bounded() {
        // One real 512-byte page holding a single empty PAPX.
        let mut word_doc = vec![0u8; 512];
        word_doc[511] = 1; // crun = 1
        word_doc[0..4].copy_from_slice(&0x800u32.to_le_bytes());
        word_doc[4..8].copy_from_slice(&0x802u32.to_le_bytes());
        word_doc[8] = 0; // no PAPX

        // PlcfBtePapx with n = 1000 BTEs, all pointing at page 0.
        let n: usize = 1000;
        let mut plc = Vec::new();
        for _ in 0..=n {
            plc.extend_from_slice(&0u32.to_le_bytes());
        }
        for _ in 0..n {
            plc.extend_from_slice(&0u32.to_le_bytes());
        }

        let out = parse_papx_paragraphs(&word_doc, &plc, 0, plc.len() as u32);
        // Despite 1000 BTEs, only the single physical page is parsed once.
        assert_eq!(out.len(), 1, "same page referenced 1000× must parse once, not 1000×");
    }
}
