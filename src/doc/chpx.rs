//! CHPX (character properties) parsing for Word binary documents.
//!
//! Character properties live in CHPX FKP (Formatted CHpx Page) pages,
//! indexed by the PlcBteChpx in the Table stream — the character-property
//! analogue of [`super::papx`]'s PAPX walk, and structurally simpler:
//!
//! - `PlcBteChpx` (in the Table stream): `(n+1)` u32 FCs bounding each text
//!   range, then `n` `PnFkpChpx` u32s whose low 22 bits are a page number.
//!   Same shape as `PlcfBtePapx`.
//! - Each 512-byte `ChpxFkp` page sits at `pn * 512` in the WordDocument
//!   stream and holds `rgfc[crun + 1]` (u32 FCs bounding each run),
//!   `rgb[crun]` (1-byte word-offsets into the page; `rgb[i] * 2` is the
//!   byte offset of that run's `Chpx`, or 0 = no CHPX for that run), then
//!   `crun` (1 byte, the page's last byte).
//! - A `Chpx` is `[cb:1][grpprl: cb bytes]` — unlike a PAPX, there is no
//!   `cw`/`istd`/Word8-re-read framing to strip first.
//!
//! Verified against [MS-DOC] §3.4 "Example of a PlcBteChpx" (its own worked
//! byte-for-byte example), not derived from the PAPX layout by analogy.
//!
//! The `grpprl` is decoded by [`super::sprm::extract_chp_props`].

use super::papx::fc_to_cp;
use super::piece_table::Piece;
use super::sprm::{ChpProps, extract_chp_props};

/// A character-run descriptor recovered from a CHPX FKP page.
#[derive(Debug, Clone)]
pub struct FkpRun {
    /// Start FC (file character position) of the run, inclusive.
    pub fc_start: u32,
    /// End FC of the run, exclusive.
    pub fc_end: u32,
    /// The CHP `grpprl` bytes (without the `cb` length byte).
    pub grpprl: Vec<u8>,
}

/// Parse every CHPX FKP page referenced by the PlcBteChpx.
///
/// `word_doc` is the WordDocument stream (where FKP pages live);
/// `table_stream` holds the PlcBteChpx itself. Returns one `FkpRun` per
/// character run across all pages, in no particular order.
pub fn parse_chpx_runs(
    word_doc: &[u8],
    table_stream: &[u8],
    fc_plcf_bte_chpx: u32,
    lcb_plcf_bte_chpx: u32,
) -> Vec<FkpRun> {
    let start = fc_plcf_bte_chpx as usize;
    if lcb_plcf_bte_chpx < 4 || start + 4 > table_stream.len() {
        return Vec::new();
    }
    let end = (start + lcb_plcf_bte_chpx as usize).min(table_stream.len());
    let plc = &table_stream[start..end];

    // PlcBteChpx: (n+1) u32 FCs, then n u32 PnFkpChpx. n = (lcb - 4) / 8.
    let n = (plc.len().saturating_sub(4)) / 8;
    if n == 0 {
        return Vec::new();
    }
    let fc_arr = (n + 1) * 4;
    if fc_arr + n * 4 > plc.len() {
        return Vec::new();
    }

    // Bound the FKP walk the same way `parse_papx_paragraphs` does: a
    // malformed PlcBteChpx must not drive unbounded/repeated page parsing
    // (AGENTS.md rule 6).
    let max_pages = word_doc.len() / 512;
    let n = n.min(max_pages);

    let mut out = Vec::new();
    let mut visited = std::collections::HashSet::with_capacity(n.min(64));
    for i in 0..n {
        let pn_bte = u32::from_le_bytes([
            plc[fc_arr + i * 4],
            plc[fc_arr + i * 4 + 1],
            plc[fc_arr + i * 4 + 2],
            plc[fc_arr + i * 4 + 3],
        ]);
        let pn = (pn_bte & 0x003F_FFFF) as usize;
        if !visited.insert(pn) {
            continue;
        }
        if let Some(page) = word_doc.get(pn * 512..pn * 512 + 512) {
            parse_fkp_page(page, &mut out);
        }
    }
    out
}

/// Parse a single 512-byte CHPX FKP page, appending `FkpRun`s to `out`.
fn parse_fkp_page(page: &[u8], out: &mut Vec<FkpRun>) {
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

    // rgb: crun 1-byte word-offsets. rgb[i] * 2 is the byte offset (from the
    // page start) of that run's Chpx; 0 = no CHPX for that run.
    for i in 0..crun {
        let rgb_off = pos + i;
        if rgb_off >= page.len() {
            break;
        }
        let word_off = page[rgb_off] as usize;
        let fc_start = rgfc[i];
        let fc_end = rgfc[i + 1];
        let grpprl = if word_off == 0 {
            Vec::new()
        } else {
            extract_chpx_grpprl(page, word_off)
        };
        out.push(FkpRun {
            fc_start,
            fc_end,
            grpprl,
        });
    }
}

/// Extract the CHP `grpprl` from a page at the given word offset.
///
/// Layout: `[cb:1][grpprl: cb bytes]` — no `cw`/`istd` header, no Word8
/// re-read (that quirk is PAPX-specific).
fn extract_chpx_grpprl(page: &[u8], word_off: usize) -> Vec<u8> {
    let p = word_off * 2;
    if p >= page.len() {
        return Vec::new();
    }
    let cb = page[p] as usize;
    let start = p + 1;
    let end = (start + cb).min(page.len());
    if start >= end {
        return Vec::new();
    }
    page[start..end].to_vec()
}

/// Resolve the CP ranges within `[0, text_len)` whose CHP marks them as
/// deleted revision-mark text (`sprmCFRMarkDel`), merged and sorted.
///
/// This is the "at minimum" fix for issue #288: the accepted-view policy
/// already applied to DOCX's `w:del` extended to DOC, without attempting
/// the larger #287 (full character-formatting fidelity).
pub fn resolve_deleted_cp_ranges(
    word_doc: &[u8],
    table_stream: &[u8],
    pieces: &[Piece],
    fc_plcf_bte_chpx: u32,
    lcb_plcf_bte_chpx: u32,
    text_len: u32,
) -> Vec<(u32, u32)> {
    let runs = parse_chpx_runs(word_doc, table_stream, fc_plcf_bte_chpx, lcb_plcf_bte_chpx);
    let mut deleted: Vec<(u32, u32)> = runs
        .iter()
        .filter(|r| {
            let props: ChpProps = extract_chp_props(&r.grpprl);
            props.f_rmark_del
        })
        .filter_map(|r| {
            let cp_start = fc_to_cp(r.fc_start, pieces)?;
            let cp_end = fc_to_cp(r.fc_end, pieces).unwrap_or(cp_start);
            let cp_start = cp_start.min(text_len);
            let cp_end = cp_end.min(text_len);
            if cp_end > cp_start { Some((cp_start, cp_end)) } else { None }
        })
        .collect();

    deleted.sort_unstable_by_key(|&(s, _)| s);
    merge_ranges(deleted)
}

/// Merge overlapping/adjacent `(start, end)` ranges (already sorted by
/// start) into a minimal disjoint set.
fn merge_ranges(ranges: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
    let mut out: Vec<(u32, u32)> = Vec::with_capacity(ranges.len());
    for (start, end) in ranges {
        if let Some(last) = out.last_mut() {
            if start <= last.1 {
                last.1 = last.1.max(end);
                continue;
            }
        }
        out.push((start, end));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn unicode_piece(fc: u32, cp_end: u32) -> Piece {
        Piece {
            cp_start: 0,
            cp_end,
            fc,
            is_compressed: false,
        }
    }

    /// Regression, byte-for-byte from [MS-DOC] §3.4's own worked example:
    /// a page with 3 runs, two carrying a Chpx (word offsets 0xFA and
    /// 0xF8) and one with no CHPX (word offset 0).
    #[test]
    fn test_parse_fkp_page_matches_spec_worked_example() {
        let mut page = vec![0u8; 512];
        // rgfc[0..4] = 0x400, 0x407, 0x410, 0x411
        page[0..4].copy_from_slice(&0x400u32.to_le_bytes());
        page[4..8].copy_from_slice(&0x407u32.to_le_bytes());
        page[8..12].copy_from_slice(&0x410u32.to_le_bytes());
        page[12..16].copy_from_slice(&0x411u32.to_le_bytes());
        // rgb[0..3] = 0xFA, 0xF8, 0x00
        page[16] = 0xFA;
        page[17] = 0xF8;
        page[18] = 0x00;
        page[511] = 3; // crun

        // chpx[0] at word offset 0xFA -> byte 0x1F4: cb=9, 9-byte grpprl.
        let chpx0_grpprl = [0x42, 0x0A, 0x07, 0x70, 0x1C, 0xFF, 0x99, 0x00, 0x00];
        page[0x1F4] = chpx0_grpprl.len() as u8;
        page[0x1F5..0x1F5 + chpx0_grpprl.len()].copy_from_slice(&chpx0_grpprl);

        // chpx[1] at word offset 0xF8 -> byte 0x1F0: cb=3, 3-byte grpprl.
        let chpx1_grpprl = [0x3E, 0x08, 0x01];
        page[0x1F0] = chpx1_grpprl.len() as u8;
        page[0x1F1..0x1F1 + chpx1_grpprl.len()].copy_from_slice(&chpx1_grpprl);

        let mut out = Vec::new();
        parse_fkp_page(&page, &mut out);
        assert_eq!(out.len(), 3);
        assert_eq!(out[0].fc_start, 0x400);
        assert_eq!(out[0].fc_end, 0x407);
        assert_eq!(out[0].grpprl, chpx0_grpprl);
        assert_eq!(out[1].fc_start, 0x407);
        assert_eq!(out[1].fc_end, 0x410);
        assert_eq!(out[1].grpprl, chpx1_grpprl);
        assert_eq!(out[2].fc_start, 0x410);
        assert_eq!(out[2].fc_end, 0x411);
        assert!(out[2].grpprl.is_empty(), "rgb[2] == 0 means no CHPX for that run");
    }

    #[test]
    fn test_extract_chpx_grpprl_no_header_stripping() {
        let mut page = vec![0u8; 512];
        page[10] = 3; // cb = 3
        page[11..14].copy_from_slice(&[0xAA, 0xBB, 0xCC]);
        // word_off * 2 == 10 -> word_off == 5
        assert_eq!(extract_chpx_grpprl(&page, 5), vec![0xAA, 0xBB, 0xCC]);
    }

    #[test]
    fn test_resolve_deleted_cp_ranges_finds_deleted_run() {
        // Two runs: cp[0,3) normal, cp[3,6) deleted (sprmCFRMarkDel = 1).
        // `resolve_deleted_cp_ranges` only reads the FKP page out of
        // `word_doc` (via the page number); it never dereferences a
        // piece's `fc` against `word_doc`, so `word_doc` here IS the FKP
        // page (page 0) and the piece's `fc = 0x800` is a purely notional
        // text-storage offset, exactly as in a real file where the FKP
        // page and the text it describes live in unrelated regions of the
        // WordDocument stream.
        let pieces = [unicode_piece(0x800, 6)];

        let mut page = vec![0u8; 512];
        page[0..4].copy_from_slice(&(0x800u32).to_le_bytes());
        page[4..8].copy_from_slice(&(0x800 + 3 * 2u32).to_le_bytes());
        page[8..12].copy_from_slice(&(0x800 + 6 * 2u32).to_le_bytes());
        page[12] = 0; // rgb[0]: no CHPX -> not deleted
        page[13] = 200; // rgb[1]: word offset 200 -> byte 400
        page[511] = 2; // crun

        let del_grpprl = [0x00, 0x08, 0x01]; // sprmCFRMarkDel = 1
        page[400] = del_grpprl.len() as u8;
        page[401..401 + del_grpprl.len()].copy_from_slice(&del_grpprl);

        // PlcBteChpx: 2 FCs (unused by this function beyond sizing) + 1
        // PnFkpChpx naming page 0.
        let mut plc = Vec::new();
        plc.extend_from_slice(&0u32.to_le_bytes());
        plc.extend_from_slice(&1024u32.to_le_bytes());
        plc.extend_from_slice(&0u32.to_le_bytes()); // page 0

        let deleted = resolve_deleted_cp_ranges(&page, &plc, &pieces, 0, plc.len() as u32, 6);
        assert_eq!(deleted, vec![(3, 6)]);
    }

    #[test]
    fn test_merge_ranges_combines_overlaps_and_keeps_disjoint() {
        assert_eq!(merge_ranges(vec![(0, 3), (2, 5), (8, 10)]), vec![(0, 5), (8, 10)]);
        assert_eq!(merge_ranges(vec![]), Vec::<(u32, u32)>::new());
    }

    /// Regression (AGENTS.md rule 6): many PnFkpChpx entries pointing at the
    /// same page must parse that page once, not repeatedly.
    #[test]
    fn test_chpx_fkp_walk_is_bounded() {
        let mut word_doc = vec![0u8; 512];
        word_doc[511] = 1; // crun = 1
        word_doc[0..4].copy_from_slice(&0x800u32.to_le_bytes());
        word_doc[4..8].copy_from_slice(&0x802u32.to_le_bytes());
        word_doc[8] = 0; // no CHPX

        let n: usize = 1000;
        let mut plc = Vec::new();
        for _ in 0..=n {
            plc.extend_from_slice(&0u32.to_le_bytes());
        }
        for _ in 0..n {
            plc.extend_from_slice(&0u32.to_le_bytes());
        }

        let out = parse_chpx_runs(&word_doc, &plc, 0, plc.len() as u32);
        assert_eq!(out.len(), 1, "same page referenced 1000× must parse once, not 1000×");
    }
}
