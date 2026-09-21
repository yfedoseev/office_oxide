//! `.doc` list formatting: `PlfLst` (list definitions) and `PlfLfo` (list
//! format overrides), the legacy analog of DOCX's `numbering.xml`.
//!
//! A paragraph's `ilfo` (`sprmPIlfo`) is a 1-based index into `PlfLfo`'s
//! `LFO` array, not into `PlfLst` directly: `LFO.lsid` is what actually
//! selects the matching `LSTF` (list definition) in `PlfLst`, and that
//! `LSTF`'s `LVL` array (one entry per level, unless `fSimpleList`) carries
//! each level's declared start-at value and number-format code. Before this
//! module, `Fib::fc_plcf_lst`/`lcb_plcf_lst` were parsed and
//! then never read again anywhere in the crate, so `List::start_number` was
//! always `None` and every list was rendered unordered regardless of its
//! real number format.
//!
//! Byte layouts verified against the published [MS-DOC] "Example of a
//! List" worked example (`PlfLst`, `LSTF`, `LVL`/`LVLF`, `PlfLfo`, `LFO`)
//! and the `LFOLVL` structure page, rather than recalled from memory.
//!
//! `LFOLVL` per-document start-at overrides (`fStartAt == 1`) ARE parsed —
//! confirmed this is the real mechanism (not a second `LSTF`/`LVL` with a
//! different declared `iStartAt`) by checking a real corpus file named for
//! exactly this scenario, `apache_tika__testWORD_override_list_numbering.doc`
//! (it turned out to use only `fFormatting == 1` overrides — full per-level
//! style replacements, not start-at ones — so it doesn't itself exercise
//! `fStartAt`, but confirmed `LFOLVL` is where a real Word-authored file
//! puts its overrides at all, which is what mattered for getting the byte
//! layout and the override-precedence logic right). A full `fFormatting ==
//! 1` override (a whole replacement `LVL`, not just its start-at value) is
//! skipped over rather than parsed — this module only ever needed
//! `start_at`/`nfc`, and a formatting-only override with no `fStartAt`
//! doesn't change either.

/// One list level's declared numbering, from an `LVL`'s `LVLF`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ListLevel {
    /// `LVLF.iStartAt` — the number sequence's start-at value for this
    /// level (only meaningful when `nfc != 0xFF`).
    pub start_at: i32,
    /// `LVLF.nfc` (`MSONFC`) — the number-format code. `0xFF` means "no
    /// number style", i.e. this level is a bullet, not a numbered level.
    pub nfc: u8,
}

impl ListLevel {
    /// `true` when this level has a real number sequence (bullet lists use
    /// `nfc == 0xFF`, per [MS-DOC] §2.9.191 MSONFC).
    pub fn is_numbered(&self) -> bool {
        self.nfc != 0xFF
    }
}

/// One list definition: an `LSTF` and its `LVL` array.
#[derive(Debug, Clone)]
struct ListDef {
    lsid: i32,
    levels: Vec<ListLevel>,
}

/// One `LFO`'s own data: which `LSTF` it points at, plus any per-level
/// start-at overrides declared directly on the `LFO` (`LFOLVL` entries
/// with `fStartAt == 1`).
#[derive(Debug, Clone)]
struct LfoEntry {
    lsid: i32,
    /// `(ilvl, overridden start_at)` pairs, one per `LFOLVL` that set
    /// `fStartAt`. Usually at most one or two entries.
    start_overrides: Vec<(u8, i32)>,
}

/// Combined `PlfLst` + `PlfLfo` lookup: resolves a paragraph's `(ilfo,
/// ilvl)` to the `ListLevel` that actually governs its numbering.
#[derive(Debug, Clone, Default)]
pub struct ListFormatting {
    list_defs: Vec<ListDef>,
    /// Indexed by `ilfo - 1` (`ilfo` is 1-based).
    lfos: Vec<LfoEntry>,
}

impl ListFormatting {
    /// Parse `PlfLst` (at `fc_plf_lst`/`lcb_plf_lst`) and `PlfLfo` (at
    /// `fc_plf_lfo`/`lcb_plf_lfo`) from the Table stream. Any parse failure
    /// (truncated/malformed data) degrades to an empty formatting table —
    /// every lookup then returns `None`, matching the pre-list-format behaviour
    /// rather than erroring the whole document.
    pub fn parse(
        table_stream: &[u8],
        fc_plf_lst: u32,
        lcb_plf_lst: u32,
        fc_plf_lfo: u32,
        lcb_plf_lfo: u32,
    ) -> Self {
        let list_defs = if lcb_plf_lst > 0 {
            let start = fc_plf_lst as usize;
            let end = start
                .saturating_add(lcb_plf_lst as usize)
                .min(table_stream.len());
            if start < end {
                parse_plf_lst(&table_stream[start..])
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };

        // `PlfLfo`'s own bounds (`lcb_plf_lfo`) cover only the fixed `LFO`
        // array; `rgLfoData`, which carries the `LFOLVL` overrides this
        // needs, follows immediately after and is NOT included in that
        // length (mirroring the identical `PlfLst`/LVL-array relationship
        // above) — so, as with `PlfLst`, parse from `start` to the end of
        // the Table stream rather than bounding to `lcb_plf_lfo`.
        let lfos = if lcb_plf_lfo > 0 {
            let start = fc_plf_lfo as usize;
            if start < table_stream.len() {
                parse_plf_lfo(&table_stream[start..])
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };

        Self { list_defs, lfos }
    }

    /// Resolve a paragraph's declared list level, given its `ilfo`
    /// (`sprmPIlfo`, 1-based; `0`/negative/out-of-range means "not in a
    /// list", per [MS-DOC] §2.4.6.3) and `ilvl` (0-based nesting depth).
    pub fn level_for(&self, ilfo: i16, ilvl: u8) -> Option<ListLevel> {
        if ilfo <= 0 {
            return None;
        }
        let lfo = self.lfos.get((ilfo - 1) as usize)?;
        let def = self.list_defs.iter().find(|d| d.lsid == lfo.lsid)?;
        // A level deeper than the list defines (malformed input, or a
        // simple/1-level list referenced at ilvl > 0) falls back to the
        // list's own first (and, for a simple list, only) level rather
        // than returning nothing.
        let mut level = def
            .levels
            .get(ilvl as usize)
            .or_else(|| def.levels.first())
            .copied()?;
        if let Some(&(_, start_at)) = lfo.start_overrides.iter().find(|(l, _)| *l == ilvl) {
            level.start_at = start_at;
        }
        Some(level)
    }
}

/// Parse the `PlfLst` structure: `cLst: u16`, then `cLst` `LSTF` entries
/// (28 bytes each), then a flat array of `LVL`s — 9 per non-simple `LSTF`,
/// 1 per simple one, in the same order as the `LSTF` array.
fn parse_plf_lst(data: &[u8]) -> Vec<ListDef> {
    if data.len() < 2 {
        return Vec::new();
    }
    let c_lst = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut pos = 2usize;

    // First pass: read every LSTF (lsid, fSimpleList) and remember where
    // the LVL array for each one begins once we know each LSTF's level
    // count (needed to size the LVL array itself, per the spec's own
    // "size of PlfLst does not include the LVL array" note).
    struct LstfInfo {
        lsid: i32,
        levels_count: usize,
    }
    let mut lstfs = Vec::with_capacity(c_lst.min(4096));
    for _ in 0..c_lst {
        if pos + 28 > data.len() {
            break;
        }
        let lsid = i32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
        // Offset 24 (0-based within the 28-byte LSTF): the fSimpleList /
        // unused1 / fAutoNum / unused2 / fHybrid / reserved1 bitfield byte.
        let flags = data[pos + 24];
        let f_simple_list = (flags & 0x01) != 0;
        lstfs.push(LstfInfo {
            lsid,
            levels_count: if f_simple_list { 1 } else { 9 },
        });
        pos += 28;
    }

    // Second pass: walk the flat LVL array in the same order, consuming
    // each LSTF's declared level count.
    let mut defs = Vec::with_capacity(lstfs.len());
    for info in lstfs {
        let mut levels = Vec::with_capacity(info.levels_count);
        for _ in 0..info.levels_count {
            match parse_one_lvl(data, pos) {
                Some((level, consumed)) => {
                    levels.push(level);
                    pos += consumed;
                },
                None => break, // truncated — keep whatever levels parsed so far
            }
        }
        defs.push(ListDef {
            lsid: info.lsid,
            levels,
        });
    }
    defs
}

/// Parse one `LVL` starting at `pos`, returning the level's `(start_at,
/// nfc)` and the total byte length of this `LVL` (so the caller can
/// advance past its variable-length `grpprlPapx`/`grpprlChpx`/`Xst` tail).
fn parse_one_lvl(data: &[u8], pos: usize) -> Option<(ListLevel, usize)> {
    // LVLF is a fixed 28 bytes: iStartAt(4) is all we need, but
    // cbGrpprlChpx/cbGrpprlPapx (offsets 24/25 into it) are needed to skip
    // to the next LVL.
    if pos + 28 > data.len() {
        return None;
    }
    let start_at = i32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    let nfc = data[pos + 4];
    let cb_grpprl_chpx = data[pos + 24] as usize;
    let cb_grpprl_papx = data[pos + 25] as usize;

    let after_lvlf = pos + 28;
    let after_grpprls = after_lvlf
        .checked_add(cb_grpprl_papx)?
        .checked_add(cb_grpprl_chpx)?;
    if after_grpprls + 2 > data.len() {
        return None;
    }
    // Xst: cch: u16, then cch UTF-16 code units.
    let cch = u16::from_le_bytes([data[after_grpprls], data[after_grpprls + 1]]) as usize;
    let xst_len = 2 + cch.saturating_mul(2);
    let total = after_grpprls.checked_add(xst_len)?.checked_sub(pos)?;

    Some((ListLevel { start_at, nfc }, total))
}

/// Parse the `PlfLfo` structure: `lfoMac: u32`, then `lfoMac` `LFO` entries
/// (16 bytes each: `lsid`, 2 unused `u32`s, `clfolvl: u8`, + 3 more bytes),
/// then `lfoMac` `LFOData` entries (each `cp: u32` + `clfolvl` `LFOLVL`
/// overrides). Returns one `LfoEntry` per `LFO`, in array order (index `i`
/// == `ilfo - 1`).
fn parse_plf_lfo(data: &[u8]) -> Vec<LfoEntry> {
    if data.len() < 4 {
        return Vec::new();
    }
    let lfo_mac = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let mut pos = 4usize;

    struct LfoInfo {
        lsid: i32,
        clfolvl: usize,
    }
    let mut lfo_infos = Vec::with_capacity(lfo_mac.min(4096));
    for _ in 0..lfo_mac {
        if pos + 16 > data.len() {
            break;
        }
        let lsid = i32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
        let clfolvl = data[pos + 12] as usize;
        lfo_infos.push(LfoInfo { lsid, clfolvl });
        pos += 16;
    }

    // LFOData array: one per LFO — cp (4 bytes, ignored) + clfolvl LFOLVL
    // overrides.
    let mut result = Vec::with_capacity(lfo_infos.len());
    for info in lfo_infos {
        pos += 4; // cp
        let mut start_overrides = Vec::new();
        for _ in 0..info.clfolvl {
            match parse_one_lfolvl(data, pos) {
                Some((ilvl, start_at, consumed)) => {
                    if let Some(start_at) = start_at {
                        start_overrides.push((ilvl, start_at));
                    }
                    pos += consumed;
                },
                None => break, // truncated — keep whatever overrides parsed so far
            }
        }
        result.push(LfoEntry {
            lsid: info.lsid,
            start_overrides,
        });
    }
    result
}

/// Parse one `LFOLVL` at `pos`: `iStartAt: i32`, then a `u32` bitfield
/// (`iLvl`: 4 bits, `fStartAt`: 1 bit, `fFormatting`: 1 bit, `grfhic`: 8
/// bits, unused: 18 bits), then — only when `fFormatting == 1` — a
/// complete embedded `LVL` that fully replaces the corresponding one.
///
/// Returns `(iLvl, overridden start_at or None, total byte length)`. The
/// embedded `LVL` (when present) is skipped over via the same variable-
/// length walk `parse_one_lvl` uses, not parsed into a `ListLevel` — this
/// module only tracks start-at overrides, not full formatting overrides.
fn parse_one_lfolvl(data: &[u8], pos: usize) -> Option<(u8, Option<i32>, usize)> {
    if pos + 8 > data.len() {
        return None;
    }
    let start_at_raw = i32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    let flags = u32::from_le_bytes([data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]]);
    let i_lvl = (flags & 0x0F) as u8;
    let f_start_at = (flags >> 4) & 1 == 1;
    let f_formatting = (flags >> 5) & 1 == 1;

    let mut total = 8usize;
    if f_formatting {
        let (_, lvl_len) = parse_one_lvl(data, pos + 8)?;
        total += lvl_len;
    }

    let start_at = (f_start_at && !f_formatting).then_some(start_at_raw);
    Some((i_lvl, start_at, total))
}

#[cfg(test)]
impl ListFormatting {
    /// Build directly from list definitions, bypassing `PlfLst`/`PlfLfo`
    /// byte parsing — lets other modules' tests (e.g. `convert_doc.rs`)
    /// exercise `start_number`/`ordered` resolution without constructing
    /// raw Table-stream bytes; the byte-level parser has its own tests in
    /// this module.
    pub(crate) fn from_parts(defs: Vec<(i32, Vec<ListLevel>)>, lfo_lsids: Vec<i32>) -> Self {
        Self {
            list_defs: defs
                .into_iter()
                .map(|(lsid, levels)| ListDef { lsid, levels })
                .collect(),
            lfos: lfo_lsids
                .into_iter()
                .map(|lsid| LfoEntry {
                    lsid,
                    start_overrides: Vec::new(),
                })
                .collect(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The exact byte layout from [MS-DOC]'s own "Example of a List"
    /// worked example: one list (`lsid = 0x44F53D09`), 9 levels
    /// (`fSimpleList = 0`), and one `LFO` whose `lsid` matches it.
    /// `lvl[0].iStartAt = 1`, `lvl[1].iStartAt = 3`, `lvl[2]` is unnumbered
    /// (`nfc = 0xFF`).
    fn spec_example_table_stream() -> Vec<u8> {
        let mut data = vec![0u8; 0x1000];

        // PlfLst at 0x0536: cLst = 1, then one 28-byte LSTF.
        let plf_lst_start = 0x0536;
        data[plf_lst_start..plf_lst_start + 2].copy_from_slice(&1u16.to_le_bytes());
        let lstf_start = plf_lst_start + 2;
        data[lstf_start..lstf_start + 4].copy_from_slice(&0x44F53D09u32.to_le_bytes()); // lsid
        data[lstf_start + 4..lstf_start + 8].copy_from_slice(&0x31200A2Cu32.to_le_bytes()); // tplc
        // rgistdPara[9] left as zero; flags byte (offset 24) = 0 -> fSimpleList=0 (9 levels).
        data[lstf_start + 24] = 0x00;

        // LVL array begins at fcPlfLst + lcbPlfLst = 0x0536 + 0x1E = 0x0554.
        let mut pos = 0x0554usize;
        // lvl[0]: iStartAt=1, nfc=0x00, cbGrpprlChpx=0x0D, cbGrpprlPapx=0x18, Xst cch=2.
        pos = write_test_lvl(&mut data, pos, 1, 0x00, 0x0D, 0x18, 2);
        // lvl[1]: iStartAt=3, nfc=0x04, cbGrpprlChpx=0x0D, cbGrpprlPapx=0x18, Xst cch=4.
        pos = write_test_lvl(&mut data, pos, 3, 0x04, 0x0D, 0x18, 4);
        // lvl[2]: iStartAt=1 (ignored, unnumbered), nfc=0xFF, cbGrpprlChpx=0x0D, cbGrpprlPapx=0x10, Xst cch=8.
        pos = write_test_lvl(&mut data, pos, 1, 0xFF, 0x0D, 0x10, 8);
        // Remaining 6 levels: minimal (cch=0), just to keep the array well-formed.
        for _ in 0..6 {
            pos = write_test_lvl(&mut data, pos, 1, 0xFF, 0, 0, 0);
        }
        let _ = pos;

        // PlfLfo at 0x07E1: lfoMac=1, one 16-byte LFO (lsid matches the LSTF above).
        let plf_lfo_start = 0x07E1;
        data[plf_lfo_start..plf_lfo_start + 4].copy_from_slice(&1u32.to_le_bytes());
        let lfo_start = plf_lfo_start + 4;
        data[lfo_start..lfo_start + 4].copy_from_slice(&0x44F53D09u32.to_le_bytes()); // lsid
        data[lfo_start + 12] = 0; // clfolvl = 0

        data
    }

    /// Write one `LVL` at `pos`, return the offset just past it.
    fn write_test_lvl(
        data: &mut [u8],
        pos: usize,
        start_at: i32,
        nfc: u8,
        cb_chpx: u8,
        cb_papx: u8,
        xst_cch: u16,
    ) -> usize {
        data[pos..pos + 4].copy_from_slice(&start_at.to_le_bytes());
        data[pos + 4] = nfc;
        data[pos + 24] = cb_chpx;
        data[pos + 25] = cb_papx;
        let after_lvlf = pos + 28;
        let after_grpprls = after_lvlf + cb_papx as usize + cb_chpx as usize;
        data[after_grpprls..after_grpprls + 2].copy_from_slice(&xst_cch.to_le_bytes());
        after_grpprls + 2 + (xst_cch as usize) * 2
    }

    #[test]
    fn test_resolves_start_at_and_nfc_from_the_spec_worked_example() {
        let table = spec_example_table_stream();
        let fmt = ListFormatting::parse(&table, 0x0536, 0x001E, 0x07E1, 0x0018);

        let lvl0 = fmt.level_for(1, 0).expect("ilfo=1, ilvl=0 must resolve");
        assert_eq!(lvl0.start_at, 1);
        assert!(lvl0.is_numbered());

        let lvl1 = fmt.level_for(1, 1).expect("ilfo=1, ilvl=1 must resolve");
        assert_eq!(lvl1.start_at, 3);
        assert!(lvl1.is_numbered());

        let lvl2 = fmt.level_for(1, 2).expect("ilfo=1, ilvl=2 must resolve");
        assert!(!lvl2.is_numbered(), "nfc=0xFF must report as unnumbered (a bullet level)");
    }

    /// The real-world override shape: a corpus file literally
    /// named `apache_tika__testWORD_override_list_numbering.doc` turned out
    /// to express its start-at override via a `LFOLVL` entry on the `LFO`
    /// (`fStartAt == 1`), not via a second `LSTF`/`LVL` pair with a
    /// different declared `iStartAt`. Byte layout verified against the
    /// published `LFOLVL` structure page.
    #[test]
    fn test_lfolvl_start_at_override_takes_precedence_over_the_lstfs_own_value() {
        let mut data = spec_example_table_stream();

        // Rewrite the LFO's clfolvl from 0 to 1, and append one LFOLVL
        // overriding level 0's start-at to 42.
        let plf_lfo_start = 0x07E1;
        let lfo_start = plf_lfo_start + 4;
        data[lfo_start + 12] = 1; // clfolvl = 1

        // LFOData begins right after the (single) LFO: cp (4 bytes,
        // ignored) then the one LFOLVL.
        let lfo_data_start = lfo_start + 16;
        let lfolvl_start = lfo_data_start + 4;
        data[lfolvl_start..lfolvl_start + 4].copy_from_slice(&42i32.to_le_bytes()); // iStartAt
        // flags: iLvl=0 (bits 0-3), fStartAt=1 (bit 4), fFormatting=0 (bit 5).
        let flags: u32 = 1 << 4;
        data[lfolvl_start + 4..lfolvl_start + 8].copy_from_slice(&flags.to_le_bytes());

        let fmt = ListFormatting::parse(&data, 0x0536, 0x001E, 0x07E1, 0x0018);

        let overridden = fmt.level_for(1, 0).expect("ilfo=1, ilvl=0 must resolve");
        assert_eq!(
            overridden.start_at, 42,
            "the LFOLVL override must win over the LSTF's own iStartAt=1"
        );
        assert!(overridden.is_numbered(), "nfc is unaffected by a start-at-only override");

        // Level 1 has no override — must still be the LSTF's own value.
        let unaffected = fmt.level_for(1, 1).expect("ilfo=1, ilvl=1 must resolve");
        assert_eq!(unaffected.start_at, 3, "an override on level 0 must not leak into level 1");
    }

    #[test]
    fn test_ilfo_zero_or_negative_is_not_in_a_list() {
        let table = spec_example_table_stream();
        let fmt = ListFormatting::parse(&table, 0x0536, 0x001E, 0x07E1, 0x0018);
        assert!(fmt.level_for(0, 0).is_none());
        assert!(fmt.level_for(-1, 0).is_none());
    }

    #[test]
    fn test_unknown_ilfo_resolves_to_none_not_a_panic() {
        let table = spec_example_table_stream();
        let fmt = ListFormatting::parse(&table, 0x0536, 0x001E, 0x07E1, 0x0018);
        assert!(fmt.level_for(99, 0).is_none());
    }

    #[test]
    fn test_empty_table_stream_never_panics() {
        let fmt = ListFormatting::parse(&[], 0, 0, 0, 0);
        assert!(fmt.level_for(1, 0).is_none());
    }

    #[test]
    fn test_truncated_plf_lst_degrades_gracefully() {
        // Claims a large lcb but the actual data is short — must not panic
        // or read out of bounds.
        let table = vec![0xFFu8; 16];
        let fmt = ListFormatting::parse(&table, 0, 10_000, 0, 0);
        assert!(fmt.level_for(1, 0).is_none());
    }

    /// A level index deeper than the list actually defines (a simple list
    /// referenced at ilvl > 0) must fall back to the list's first level
    /// rather than returning `None`.
    #[test]
    fn test_ilvl_past_the_lists_own_depth_falls_back_to_the_first_level() {
        let table = spec_example_table_stream();
        let fmt = ListFormatting::parse(&table, 0x0536, 0x001E, 0x07E1, 0x0018);
        let deep = fmt
            .level_for(1, 50)
            .expect("out-of-range ilvl must fall back, not None");
        assert_eq!(deep.start_at, 1); // lvl[0]'s value
    }
}
