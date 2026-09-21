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

use super::chpx::{FkpRun, resolve_chp_cp_runs, resolve_chp_segments};
use super::piece_table::{
    HyperlinkSpan, Piece, decode_cp_range, sanitize_text_with_hyperlinks_and_chp,
};
use super::sprm::{ChpProps, PapProps};

/// A paragraph descriptor recovered from a PAPX FKP page.
#[derive(Debug, Clone)]
pub struct FkpParagraph {
    /// Start FC (file character position) of the paragraph, inclusive.
    pub fc_start: u32,
    /// End FC of the paragraph, exclusive.
    pub fc_end: u32,
    /// The PAP `grpprl` bytes (without the `cw`/`istd` header).
    pub grpprl: Vec<u8>,
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
    /// `HYPERLINK` field display-text spans (byte ranges into `text`)
    /// paired with their target URLs.
    pub hyperlinks: Vec<HyperlinkSpan>,
    /// Character-property run boundaries (byte ranges into `text`) from
    /// CHPX: bold/italic/underline/color/font-size per run.
    /// Always covers the whole of `text` when `text` is non-empty — a
    /// document with no CHPX FKP at all naturally collapses to a single
    /// span with `ChpProps::default()` (see `resolve_chp_segments`), so
    /// callers never need a separate "no formatting info" case.
    pub chp_runs: Vec<(std::ops::Range<usize>, ChpProps)>,
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
        let grpprl = if word_off == 0 {
            Vec::new()
        } else {
            extract_grpprl(page, word_off)
        };
        out.push(FkpParagraph {
            fc_start,
            fc_end,
            grpprl,
        });
    }
}

/// Extract the PAP `grpprl` from a page at the given word offset.
///
/// Layout: `[cw:1][istd:2][grpprl: cb-3]`, `cb = cw * 2`. The Word8 re-read
/// (`cw == 0` → use the next byte) is applied so row-terminator paragraphs,
/// which carry the full TAP, are not mistaken for empty PAPXs.
fn extract_grpprl(page: &[u8], word_off: usize) -> Vec<u8> {
    let mut p = word_off * 2;
    if p >= page.len() {
        return Vec::new();
    }
    let mut cw = page[p] as usize;
    let reread = cw == 0;
    if reread {
        // Word8 re-read: the real cw is the following byte.
        p += 1;
        if p >= page.len() {
            return Vec::new();
        }
        cw = page[p] as usize;
    }
    let cb = cw * 2; // total PAPX bytes for the istd+grpprl block
    if cb < 3 {
        return Vec::new(); // only cw + istd, no grpprl
    }
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
        return Vec::new();
    }
    page[grpprl_start..grpprl_end].to_vec()
}

/// The CP ranges an FC run `[fc_start, fc_end)` covers, in CP order.
///
/// In a fast-saved (complex) file the text of one FC run is not one CP
/// range: the piece table scatters edits, so consecutive file bytes can
/// be far apart in the character space and vice versa. Mapping a run by
/// its two end points (CP of the start, CP of the end) silently took
/// whatever CPs lay *between* those two points — a paragraph gained the
/// text of its neighbours, a formatting run covered characters it did
/// not format, a start below every piece mapped to nothing. Each piece is
/// intersected with the run instead.
pub fn fc_run_to_cp_ranges(fc_start: u32, fc_end: u32, pieces: &[Piece]) -> Vec<(u32, u32)> {
    let norm = |fc: u32| {
        if fc & 0x4000_0000 != 0 {
            (fc & !0x4000_0000) / 2
        } else {
            fc
        }
    };
    let (run_start, run_end) = (norm(fc_start) as u64, norm(fc_end) as u64);
    let mut out = Vec::new();
    if run_end <= run_start {
        return out;
    }
    for p in pieces {
        if p.cp_end <= p.cp_start {
            continue;
        }
        let (base, stride) = piece_byte_base(p);
        let (base, stride) = (base as u64, stride as u64);
        let piece_end = base.saturating_add((p.cp_end - p.cp_start) as u64 * stride);
        let lo = run_start.max(base);
        let hi = run_end.min(piece_end);
        if hi <= lo {
            continue;
        }
        let cp_lo = p.cp_start as u64 + (lo - base) / stride;
        let cp_hi = p.cp_start as u64 + (hi - base).div_ceil(stride);
        if cp_hi > cp_lo {
            out.push((cp_lo.min(u32::MAX as u64) as u32, cp_hi.min(u32::MAX as u64) as u32));
        }
    }
    out.sort_unstable();
    out
}

/// Real byte offset and stride (bytes per character) of a piece's start.
fn piece_byte_base(p: &Piece) -> (u32, u32) {
    if p.is_compressed {
        ((p.fc & !0x4000_0000) / 2, 1)
    } else {
        (p.fc, 2)
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
    lid: u16,
    chp_runs: &[FkpRun],
) -> Vec<DocParagraph> {
    // PAPX runs in CP space: every FKP run intersected with every piece.
    // A paragraph is then the text up to and including the next paragraph
    // mark, and its properties are the run holding that mark ([MS-DOC]
    // §2.4.6.1 — the PAPX applies to the paragraph whose mark it covers).
    // Taking each FKP run as one paragraph instead, with its two FC end
    // points mapped to CPs, was right only for a never-fast-saved file.
    let mut pap: Vec<(u32, u32, &FkpParagraph)> = fkp
        .iter()
        .flat_map(|fp| {
            fc_run_to_cp_ranges(fp.fc_start, fp.fc_end, pieces)
                .into_iter()
                .map(move |(a, b)| (a.min(text_len), b.min(text_len), fp))
        })
        .filter(|(a, b, _)| b > a)
        .collect();
    pap.sort_by_key(|&(a, b, _)| (a, b));

    // FC→CP-convert and decode every CHPX run exactly once for the whole
    // document — `resolve_chp_segments` is called once per interval
    // below, and both the piece walk and `extract_chp_props` are real work;
    // doing either of them per paragraph instead of once here turned a
    // real corpus sweep into a multi-minute hang.
    let sorted_chp_runs = resolve_chp_cp_runs(chp_runs, pieces);

    let mut out = Vec::with_capacity(pap.len());
    let mut buf: Vec<char> = Vec::new();
    let mut buf_props: Vec<ChpProps> = Vec::new();
    let mut cursor = 0u32;
    let empty = FkpParagraph {
        fc_start: 0,
        fc_end: 0,
        grpprl: Vec::new(),
    };
    for idx in 0..pap.len() {
        let (a, b, fp) = pap[idx];
        // Text no PAPX run covers still belongs to a paragraph: decode it
        // with default properties rather than lose it from the structured
        // view (the flat text has it).
        let intervals: [(u32, u32, &FkpParagraph); 2] = [
            (cursor, a.max(cursor), &empty),
            (a.max(cursor), b.max(cursor), fp),
        ];
        for &(seg_a, seg_b, run) in &intervals {
            if seg_b <= seg_a {
                continue;
            }
            cursor = seg_b;
            let segments = resolve_chp_segments(&sorted_chp_runs, seg_a, seg_b);
            for (seg_start, seg_end, props) in &segments {
                let chunk = decode_cp_range(word_doc, pieces, *seg_start, *seg_end, lid);
                for ch in chunk.chars() {
                    let is_mark = matches!(ch, '\r' | '\u{7}');
                    // Deleted revision-mark text (`sprmCFRMarkDel`) is
                    // not part of the accepted view: the flat text already
                    // excludes it, and keeping it here made
                    // `to_ir()`/`to_html()` show sentences `plain_text()`
                    // did not. The paragraph mark itself stays, so a
                    // wholly deleted paragraph still terminates.
                    if props.f_rmark_del && !is_mark {
                        continue;
                    }
                    buf.push(ch);
                    buf_props.push(props.clone());
                    if is_mark {
                        emit_paragraph(&mut buf, &mut buf_props, run, &mut out);
                    }
                }
            }
        }
    }
    if !buf.is_empty() {
        // Text after the last mark (a truncated or mark-less file): keep
        // it as a final paragraph.
        buf.push('\r');
        buf_props.push(buf_props.last().cloned().unwrap_or_default());
        let last = pap.last().map(|&(_, _, fp)| fp).unwrap_or(&empty);
        emit_paragraph(&mut buf, &mut buf_props, last, &mut out);
    }
    out
}

/// Turn the buffered characters (ending with their paragraph mark) into a
/// `DocParagraph` with `run`'s properties, and clear the buffers.
fn emit_paragraph(
    buf: &mut Vec<char>,
    buf_props: &mut Vec<ChpProps>,
    run: &FkpParagraph,
    out: &mut Vec<DocParagraph>,
) {
    let terminator = buf[buf.len() - 1];
    let content_str: String = buf[..buf.len() - 1].iter().collect();
    let content_props = &buf_props[..buf_props.len() - 1];
    let (content, hyperlinks, chp_spans) =
        sanitize_text_with_hyperlinks_and_chp(&content_str, content_props);
    let props = super::sprm::extract_pap_props(&run.grpprl);
    out.push(DocParagraph {
        text: content,
        terminator,
        props,
        hyperlinks,
        chp_runs: chp_spans,
    });
    buf.clear();
    buf_props.clear();
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::piece_table::Piece;
    use crate::doc::sprm::extract_pap_props;

    fn unicode_piece(fc: u32, cp_end: u32) -> Piece {
        Piece {
            cp_start: 0,
            cp_end,
            fc,
            is_compressed: false,
        }
    }

    #[test]
    fn test_fc_run_maps_to_cp_ranges_in_a_unicode_piece() {
        // table.doc: one Unicode piece, fc = 0x800, text_len = 23.
        let pieces = [unicode_piece(0x800, 23)];
        // bytes 0x802..0x806 = cps 1..3.
        assert_eq!(fc_run_to_cp_ranges(0x802, 0x806, &pieces), vec![(1, 3)]);
        // A run reaching below and beyond the piece is clipped to it.
        assert_eq!(fc_run_to_cp_ranges(0x700, 0x900, &pieces), vec![(0, 23)]);
        assert!(fc_run_to_cp_ranges(0x900, 0x910, &pieces).is_empty());
    }

    #[test]
    fn test_fc_run_strips_the_compressed_bit() {
        // A compressed piece with bit 30 set; an FC may carry the same bit.
        let pieces = [Piece {
            cp_start: 0,
            cp_end: 5,
            fc: 0x4000_0010, // compressed, real offset = 0x10/2 = 8
            is_compressed: true,
        }];
        assert_eq!(fc_run_to_cp_ranges(0x4000_0010, 0x4000_0014, &pieces), vec![(0, 2)]);
        // fc 9 (real byte, no bit) → cp 1.
        assert_eq!(fc_run_to_cp_ranges(9, 11, &pieces), vec![(1, 3)]);
    }

    /// A fast-saved file scatters one paragraph's characters over the
    /// pieces: the run's CPs are the union of its intersections with each
    /// piece, in CP order, never "everything between the two end points".
    #[test]
    fn test_fc_run_across_scattered_pieces_yields_each_pieces_cps() {
        // cp 0..10 at bytes 0x800.., cp 10..20 at bytes 0x100.. (edited
        // later, stored earlier), both 8-bit.
        let pieces = [
            Piece {
                cp_start: 0,
                cp_end: 10,
                fc: 0x4000_0000 | (0x800 * 2),
                is_compressed: true,
            },
            Piece {
                cp_start: 10,
                cp_end: 20,
                fc: 0x4000_0000 | (0x100 * 2),
                is_compressed: true,
            },
        ];
        // One FC run covering bytes 0x100..0x900: both pieces, in CP order.
        assert_eq!(fc_run_to_cp_ranges(0x100, 0x900, &pieces), vec![(0, 10), (10, 20)]);
        // A run over the later-stored piece only.
        assert_eq!(fc_run_to_cp_ranges(0x105, 0x108, &pieces), vec![(15, 18)]);
    }

    #[test]
    fn test_extract_grpprl_handles_word8_reread() {
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
    fn test_empty_papx_when_word_off_zero() {
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
    fn test_build_paragraphs_slices_text_and_flags() {
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

        let paras = build_paragraphs(&word_doc, &pieces, &fkp, 6, 0, &[]);
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

    /// Text a tracked change deleted (`sprmCFRMarkDel`) is left out of
    /// the structured paragraphs the way it is left out of the flat text,
    /// so every surface agrees on what the document says.
    #[test]
    fn test_build_paragraphs_drops_deleted_revision_text() {
        let raw = "Keep this deleted text here.\r";
        let text_bytes: Vec<u8> = raw.encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut word_doc = vec![0u8; 0x800 + text_bytes.len()];
        word_doc[0x800..0x800 + text_bytes.len()].copy_from_slice(&text_bytes);
        let n = raw.encode_utf16().count() as u32;
        let pieces = [unicode_piece(0x800, n)];
        let fkp = vec![FkpParagraph {
            fc_start: 0x800,
            fc_end: 0x800 + n * 2,
            grpprl: Vec::new(),
        }];
        // "deleted text " (cp 10..23) carries sprmCFRMarkDel = 1.
        let del = FkpRun {
            fc_start: 0x800 + 10 * 2,
            fc_end: 0x800 + 23 * 2,
            grpprl: vec![0x00, 0x08, 0x01],
        };
        let plain = |a: u32, b: u32| FkpRun {
            fc_start: 0x800 + a * 2,
            fc_end: 0x800 + b * 2,
            grpprl: Vec::new(),
        };
        let runs = vec![plain(0, 10), del, plain(23, n)];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, n, 0, &runs);
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text, "Keep this here.");
    }

    /// Regression: an astral character (emoji) is a single `char` but two
    /// UTF-16 code units, so a `Vec<char>` indexed by CP desyncs every later
    /// paragraph. Decoding each CP range directly must keep alignment. Each
    /// paragraph ends with a terminator (`\r`) as a real `.doc` does.
    #[test]
    fn test_build_paragraphs_keeps_astral_alignment() {
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
        };
        let fkp = vec![mk(0, 6), mk(6, 12)];

        let paras = build_paragraphs(&word_doc, &pieces, &fkp, 12, 0, &[]);
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text, "Hi 😀", "emoji must not desync the range");
        assert_eq!(paras[0].terminator, '\r');
        assert_eq!(paras[1].text, "there", "second paragraph must be intact");
        assert_eq!(paras[1].terminator, '\r');
    }

    /// Regression: field codes (0x13/0x14/0x15) in a paragraph must be
    /// stripped from the IR text, matching the sanitised plain-text path.
    ///
    /// The instruction text between `0x13` and `0x14`
    /// ("HYPERLINK ...", the field's own code, never shown by Word) must
    /// be dropped entirely, not just its boundary markers; only the cached
    /// result (between `0x14` and `0x15`) is visible text. This test used
    /// to assert the opposite (`"SeeHYPERLINKresulthere"`, keeping both
    /// halves) — that was the bug, not the contract.
    #[test]
    fn test_build_paragraphs_strips_field_codes() {
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
        };
        let fkp = vec![mk(0, raw.chars().count() as u32)];
        let paras = build_paragraphs(&word_doc, &pieces, &fkp, raw.chars().count() as u32, 0, &[]);
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].terminator, '\r');
        let t = &paras[0].text;
        assert!(!t.contains('\u{13}'), "field begin must be stripped");
        assert!(!t.contains('\u{14}'), "field separator must be stripped");
        assert!(!t.contains('\u{15}'), "field end must be stripped");
        assert!(
            !t.contains("HYPERLINK"),
            "field instruction text must not leak into visible text"
        );
        assert_eq!(t, "Seeresulthere");
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
    /// `test_extract_grpprl_handles_word8_reread` test uses `cw = 6`, so it never
    /// exercises this branch.
    #[test]
    fn test_papx_cw_zero_reread_extracts_trailing_tdef_table() {
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
        let props = extract_pap_props(&extracted);
        assert!(
            props.tap.is_some(),
            "cw==0 re-read must keep the full grpprl so the trailing TAP parses"
        );
        assert!(props.is_table_trailing_mark);
    }

    /// A piece whose CP range spans nearly the whole `u32` space must not
    /// trigger an arithmetic-overflow panic in the FC→CP walk (the
    /// `cp_end - cp_start` / `* stride` computation). Wrapped in
    /// `catch_unwind` because the failure mode is a panic; debug builds have
    /// overflow-checks enabled (AGENTS.md rule 6: no input may panic).
    #[test]
    fn test_fc_run_huge_range_does_not_panic() {
        let pieces = [Piece {
            cp_start: 0,
            cp_end: u32::MAX,
            fc: 0,
            is_compressed: false,
        }];
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            fc_run_to_cp_ranges(0, u32::MAX, &pieces)
        }));
        assert!(result.is_ok(), "the FC walk must not overflow on a huge declared CP range");
    }

    /// Regression (AGENTS.md rule 6): a PlcfBtePapx listing many BTEs that all
    /// reference the same FKP page must not cause unbounded repeated parsing.
    /// The walk dedupes visited pages and clamps `n` to the number of physical
    /// pages, so the work stays bounded even with a hostile large `lcb`.
    #[test]
    fn test_papx_fkp_walk_is_bounded() {
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
