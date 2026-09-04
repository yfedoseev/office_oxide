//! SPRM (Single Property Modifier) decoding for Word binary documents.
//!
//! A `grpprl` (group of property modifiers) is a flat byte stream of
//! consecutive SPRMs. Each SPRM is a 2-byte opcode followed by an operand
//! whose size is implied by the opcode's `spra` field (bits 15..13):
//!
//! | spra | operand size                                  |
//! |------|------------------------------------------------|
//! | 0,1  | 1 byte                                         |
//! | 2,4,5| 2 bytes                                        |
//! | 3    | 4 bytes                                        |
//! | 7    | 3 bytes                                        |
//! | 6    | variable: 1-byte length prefix, then N bytes   |
//!
//! The `sgc` field (bits 12..10) names the property class — `1` = PAP
//! (paragraph), `5` = TAP (table). See [MS-DOC] §2.4.1.
//!
//! `spra == 6` SPRMs carry a 1-byte length prefix. The *only* exception that
//! uses a genuine **2-byte** `cb` prefix is `sprmTDefTable` (`0xD608`), whose
//! `TDefTableOperand.cb` is 2 bytes by [MS-DOC] §2.9.321 (operand length is
//! `cb - 1`). `parse_grpprl` special-cases that single opcode; every other
//! variable SPRM keeps the 1-byte prefix. `sprmPChgTabs` (`0xC615`) is a
//! *different* kind of exception — its `cb` is 1 byte but with a `255` escape
//! (operand length then derived from the payload); that handling lives in the
//! list/tab-stop PR, not here. The fixtures pass either way only because their
//! `cb < 256`, so a ≥12-column table is what exposes the difference.

use crate::ir::{TabAlignment, TabLeader, TabStop};

use super::MAX_OUTLINE_LEVEL;

/// A single decoded SPRM: opcode plus its operand bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sprm {
    /// The 2-byte SPRM opcode.
    pub opcode: u16,
    /// The operand bytes (may be empty for a zero-length variable SPRM).
    pub operand: Vec<u8>,
}

impl Sprm {
    /// `spra` — operand-size class (bits 15..13 of the opcode).
    #[allow(dead_code)] // diagnostic helper; used in tests
    pub fn spra(&self) -> u8 {
        ((self.opcode >> 13) & 0x7) as u8
    }

    /// `sgc` — property class (bits 12..10): `1` = PAP, `5` = TAP, …
    #[allow(dead_code)] // diagnostic helper; used in tests
    pub fn sgc(&self) -> u8 {
        ((self.opcode >> 10) & 0x7) as u8
    }
}

/// Classification of a SPRM's operand size, used only to walk the stream.
enum OperandSize {
    /// A fixed number of operand bytes.
    Fixed(u16),
    /// Variable: a 1-byte length prefix followed by that many operand bytes.
    Variable,
}

fn operand_size(opcode: u16) -> OperandSize {
    match (opcode >> 13) & 0x7 {
        0 | 1 => OperandSize::Fixed(1),
        2 | 4 | 5 => OperandSize::Fixed(2),
        3 => OperandSize::Fixed(4),
        7 => OperandSize::Fixed(3),
        6 => OperandSize::Variable,
        // spra is 3 bits wide so this arm is unreachable; treat as 0-length.
        _ => OperandSize::Fixed(0),
    }
}

/// `spra == 6` SPRMs that carry a **2-byte** `cb` length prefix instead of
/// the usual 1-byte prefix. Per [MS-DOC] §2.9.321 the only such opcode is
/// `sprmTDefTable` (`0xD608`), whose `TDefTableOperand.cb` is 2 bytes (operand
/// length is `cb - 1`). `sprmPChgTabs` (`0xC615`) is *not* in this set: it uses
/// a 1-byte `cb` with a `255` escape (handled separately), so treating it as
/// 2-byte would shift its operand by one byte and drop `cDel`.
fn is_two_byte_len_prefix(opcode: u16) -> bool {
    opcode == 0xD608
}

// NOTE: `sprmPChgTabsPapx` (`0xC60D`) carries a normal 1-byte `cb` length
// prefix (per [MS-DOC] its `PChgTabsPapxOperand.cb` is 2..=255), so it flows
// through the default variable-length branch below — it is *not* a no-prefix
// opcode. Its delete block is `PChgTabsDel` (`1 + 2·cDel`, one XAS per tab),
// unlike `sprmPChgTabs` (`0xC615`) whose `PchgTabsDelClose` is `1 + 4·cDel`;
// the stride is selected in `decode_pchg_tabs_operand` by opcode.

/// Length of the `sprmPChgTabs` (`0xC615`) `PChgTabsOperand` when its `cb`
/// byte is the `255` escape (the normal `cb != 255` case uses the literal
/// `cb` and never reaches here). The `PchgTabsDelClose` form is `cDel` (1
/// byte) + `4*cDel` (rgdxaDel + rgdxaClose) + `cAdd` (1 byte) + `2*cAdd`
/// (rgdxaAdd) + `cAdd` (rgtbdAdd) = `2 + 4*cDel + 3*cAdd`.
///
/// The quoted [MS-DOC] formula `4 × PChgTabsDelClose.cTabs + 3 ×
/// PChgTabsAdd.cTabs` omits the two `cTabs` count bytes; the `+2` here is a
/// deliberate correction — real parsers (and Word) store the `cDel`/`cAdd`
/// counts, so do not "simplify" this back to the spec text.
fn pchg_tabs_operand_len(grpprl: &[u8], start: usize) -> usize {
    if start >= grpprl.len() {
        return 0;
    }
    let c_del = grpprl[start] as usize;
    let add_pos = start + 1 + 4 * c_del;
    if add_pos >= grpprl.len() {
        return grpprl.len() - start;
    }
    let c_add = grpprl[add_pos] as usize;
    1 + 4 * c_del + 1 + 3 * c_add
}

/// Walk a `grpprl` and decode every SPRM it contains.
///
/// Truncated operands are returned with whatever bytes remain; a truncated
/// opcode (fewer than 2 bytes left) stops the walk.
pub fn parse_grpprl(grpprl: &[u8]) -> Vec<Sprm> {
    let mut out = Vec::new();
    let mut p = 0usize;
    let len = grpprl.len();

    while p + 2 <= len {
        let opcode = u16::from_le_bytes([grpprl[p], grpprl[p + 1]]);
        let (operand, next) = match operand_size(opcode) {
            OperandSize::Fixed(n) => {
                let start = p + 2;
                let end = (start + n as usize).min(len);
                (grpprl[start..end].to_vec(), start + n as usize)
            },
            OperandSize::Variable => {
                if is_two_byte_len_prefix(opcode) {
                    // 2-byte `cb` length prefix (MS-DOC §2.2.5.1). `cb` counts
                    // the rest of the structure + 1, so the operand is `cb - 1`
                    // bytes starting after the 2-byte `cb`. Total consumed:
                    // 2 (opcode) + 2 (cb) + (cb - 1) = cb + 3.
                    if p + 4 > len {
                        // `cb` itself is truncated — stop.
                        break;
                    }
                    let cb = u16::from_le_bytes([grpprl[p + 2], grpprl[p + 3]]) as usize;
                    let start = p + 4;
                    let end = (start + cb.saturating_sub(1)).min(len);
                    (grpprl[start..end].to_vec(), start + cb.saturating_sub(1))
                } else if opcode == 0xC615 {
                    // sprmPChgTabs: a **1-byte** `cb` length prefix (NOT the
                    // 2-byte `cb` of `sprmTDefTable`, and NOT no-prefix like
                    // `sprmPChgTabsPapx`). Per [MS-DOC] §2.9.182 the operand is a
                    // `PChgTabsOperand` whose byte length is normally the literal
                    // `cb` read at `p + 2`. The single escape value `cb == 255`
                    // does not mean 255 bytes; it means the length is instead
                    // derived from the payload's own `cDel` / `cAdd` counts (see
                    // `pchg_tabs_operand_len`). The 1-byte `cb` is always
                    // consumed; the operand starts at `p + 3` either way.
                    if p + 3 > len {
                        break;
                    }
                    let cb = grpprl[p + 2] as usize;
                    let start = p + 3;
                    let n = if cb == 255 {
                        pchg_tabs_operand_len(grpprl, start)
                    } else {
                        cb
                    };
                    let end = (start + n).min(len);
                    (grpprl[start..end].to_vec(), start + n)
                } else {
                    if p + 3 > len {
                        // Length prefix itself is truncated — stop.
                        break;
                    }
                    let n = grpprl[p + 2] as usize;
                    let start = p + 3;
                    let end = (start + n).min(len);
                    (grpprl[start..end].to_vec(), start + n)
                }
            },
        };
        out.push(Sprm { opcode, operand });
        p = next;
    }

    out
}

/// Paragraph-property flags distilled from a PAP grpprl.
///
/// Only the SPRMs needed for table / list reconstruction are tracked; every
/// other SPRM is walked over and discarded.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PapProps {
    /// `sprmPFInTable` (0x2416): this paragraph lives inside a table.
    pub f_in_table: bool,
    /// The paragraph carries a TAP (table row definition, `sprmTDefTable`
    /// 0xD608). Such a paragraph is the row-terminator ("table trailing
    /// paragraph") and is NOT itself a cell.
    pub is_table_trailing_mark: bool,
    /// `sprmPItap` (0x6649): table nesting depth (1 = top-level table).
    pub itap: u8,
    /// List level (0-based) when the paragraph is a list item. `None` when
    /// the paragraph carries no list SPRM. (Reserved for list support; the
    /// `table.doc` fixture has no in-body lists to verify this against.)
    pub ilvl: Option<u8>,
    /// List format override id (`ilfo`, `sprmPIlfo` `0x460B`), read as the
    /// *signed* `i16` [MS-DOC] specifies. `None` when the SPRM is absent
    /// (defaults to "not in a list"). Bands: `0`/`0xF801` = not in a list;
    /// `0x0001`–`0x07FE` = 1-based index; `0xF802`–`0xFFFF` = negated index
    /// (still a list item — see TODO(ilfo-negated) in `convert_doc.rs`).
    pub ilfo: Option<i16>,
    /// Parsed row definition (`sprmTDefTable` operand) for row-terminator
    /// paragraphs. `None` when the paragraph is not a row mark or the TAP
    /// is malformed.
    pub tap: Option<TapInfo>,
    /// Tab stops from `sprmPChgTabs` (0xC615) / `sprmPChgTabsPapx` (0xC60D).
    /// Empty when the paragraph carries no tab-stop SPRM. Populated by the
    /// list/tab-stop PR; in the tables-only PR this field is always empty, but
    /// it is cloned through the IR so the `Paragraph.tabs` shape stays uniform.
    pub tabs: Vec<TabStop>,
    /// The style index this paragraph is restyled to by a direct `sprmPStyle`
    /// (0x640A) in the grpprl, if any. It overrides the PAPX header's own
    /// `istd` when resolving the style (see `crate::doc::styles`); `None` means
    /// the PAPX header `istd` applies unchanged.
    pub style_istd: Option<u16>,
    /// True when the grpprl carries a `sprmPOutlineLvl` (0x6412) whose operand
    /// is a valid outline level (0–[`MAX_OUTLINE_LEVEL`]).
    ///
    /// This distinguishes "the SPRM is absent" from "the SPRM is present with
    /// level 0": level 0 is an explicit *body text* marker, so when this flag
    /// is set `heading_level` is final and the paragraph's style must not be
    /// consulted. An operand outside 0–`MAX_OUTLINE_LEVEL` (e.g. 0x68) is not a
    /// valid outline level at all, leaves this flag false, and is ignored.
    pub outline_lvl_explicit: bool,
    /// Resolved heading level (1–9) when this paragraph is a heading.
    ///
    /// This is a **derived** value, not a raw SPRM property, and it is filled in
    /// two stages: `extract_pap_props` sets it from `sprmPOutlineLvl` (0x6412)
    /// when that SPRM is present; otherwise it stays `None` until
    /// `build_paragraphs` falls back to the paragraph's style via
    /// `resolve_heading_level`. So the value returned by `extract_pap_props`
    /// alone is *not* necessarily final.
    ///
    /// `None` means "not a heading" (body text), and the `.doc` walk then falls
    /// back to the line-based heading heuristic.
    pub heading_level: Option<u8>,
}

/// One table cell descriptor (TKBKTAP, 20 bytes) distilled from a row's
/// `rgdxaCenter` array. Only the fields needed for merged-cell spans are kept.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TapCellInfo {
    /// `rgf` flags (MS-DOC `TCGRF`): the 2-bit `fVertMerge` field is bits 5-6,
    /// with `fvmClear = 0x00`, `fvmMerge = 0x0020` (continuation), and
    /// `fvmRestart = 0x0060` (first cell of a merge, both bits set).
    pub rgf: u16,
    /// Preferred cell width in twips (0 = derive from `rgdxaCenter`).
    /// Kept for completeness; spans are computed from `rgdxaCenter` alone.
    #[allow(dead_code)]
    pub w_width: u16,
}

/// A table row definition (`sprmTDefTable` = 0xD608) operand.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TapInfo {
    /// Number of cells in the row (`itcMac`).
    pub itc_mac: u8,
    /// Column boundary positions in twips, `itcMac + 1` entries
    /// (`rgdxaCenter`).
    pub centers: Vec<i16>,
    /// Per-cell descriptors (`rgtc`), `itcMac` entries.
    pub cells: Vec<TapCellInfo>,
}

/// Parse a `sprmTDefTable` operand into a row definition.
///
/// Layout: `[itcMac: 1 byte][rgdxaCenter: (itcMac+1) × int16
/// LE][rgtc: itcMac × 20-byte TKBKTAP]`. The 2-byte `cb` length prefix is
/// *not* part of `operand` — `parse_grpprl` strips it before calling this.
/// Returns `None` when the operand is truncated or malformed, in which case
/// callers fall back to `col_span = row_span = 1`.
pub fn parse_tdef_table(operand: &[u8]) -> Option<TapInfo> {
    if operand.len() < 3 {
        return None;
    }
    // itcMac is the first byte of the (cb-stripped) operand.
    let itc_mac = operand[0] as usize;
    let tcs_off = 1 + (itc_mac + 1) * 2;
    if tcs_off + itc_mac * 20 > operand.len() {
        return None;
    }

    let mut centers = Vec::with_capacity(itc_mac + 1);
    for i in 0..=itc_mac {
        let off = 1 + i * 2;
        centers.push(i16::from_le_bytes([operand[off], operand[off + 1]]));
    }
    let mut cells = Vec::with_capacity(itc_mac);
    for i in 0..itc_mac {
        let off = tcs_off + i * 20;
        cells.push(TapCellInfo {
            rgf: u16::from_le_bytes([operand[off], operand[off + 1]]),
            w_width: u16::from_le_bytes([operand[off + 2], operand[off + 3]]),
        });
    }
    Some(TapInfo {
        itc_mac: itc_mac as u8,
        centers,
        cells,
    })
}

/// Decode a `sprmPChgTabs` (`0xC615`) / `sprmPChgTabsPapx` (`0xC60D`) operand
/// into tab stops, returning the *effective* (added) stops.
///
/// Both carry a delete list (tabs to ignore) followed by an add list (tabs
/// to add); the delete block shape differs by opcode (see below). Positions are
/// 16-bit signed twips; the `TBD` descriptor
/// (§2.9.310) is 1 byte carrying `jc` (bits 0..2, justification) and `tlc`
/// (bits 3..5, leader). By the time `parse_grpprl` hands the operand here the
/// 1-byte `cb` prefix has already been stripped, so `operand` is always the
/// raw structure body.
///
/// The two opcodes differ only in their *delete* block: `0xC615` uses
/// `PchgTabsDelClose` (`1 + 4·cDel` — `rgdxaDel` + `rgdxaClose`, two XAS each),
/// while `0xC60D` uses `PchgTabsDel` (`1 + 2·cDel` — `rgdxaDel` only, one XAS).
/// The stride is selected by `opcode` in `decode_pchg_tabs_operand`.
///
/// Malformed input yields an empty vector rather than panicking — tab stops
/// are formatting metadata, so a bad operand degrades to "no tabs" instead of
/// corrupting the paragraph.
pub fn decode_pchg_tabs(opcode: u16, operand: &[u8]) -> Vec<TabStop> {
    match opcode {
        0xC615 | 0xC60D => decode_pchg_tabs_operand(opcode, operand),
        _ => Vec::new(),
    }
}

/// `PChgTabsOperand` delete list then `PchgTabsAdd`. The delete stride in bytes
/// per tab depends on the opcode: `4·cDel` for `0xC615` (`PchgTabsDelClose`),
/// `2·cDel` for `0xC60D` (`PchgTabsDel`).
fn decode_pchg_tabs_operand(opcode: u16, operand: &[u8]) -> Vec<TabStop> {
    let mut tabs = Vec::new();
    // Delete stride: rgdxaDel+rgdxaClose (2 XAS) for 0xC615, rgdxaDel only
    // (1 XAS) for 0xC60D.
    let del_stride = if opcode == 0xC615 { 4 } else { 2 };
    // Delete block (§2.9.181 / §2.9.178): cTabs (u8) then `del_stride` bytes
    // per tab. Skip the whole block to reach the add list.
    let Some(c_del) = operand.first().copied() else {
        return tabs;
    };
    let mut pos = 1 + (c_del as usize).saturating_mul(del_stride);
    // PchgTabsAdd (§2.9.180): cTabs (u8), rgdxaAdd (cTabs × 2-byte XAS),
    // rgtbdAdd (cTabs × 1-byte TBD).
    let Some(c_add) = operand.get(pos).copied() else {
        return tabs;
    };
    pos += 1;
    let positions_base = pos;
    let tbd_base = pos + (c_add as usize).saturating_mul(2);
    for i in 0..c_add {
        let xas_at = positions_base + 2 * i as usize;
        if xas_at + 2 > operand.len() {
            break;
        }
        let dxp = i16::from_le_bytes([operand[xas_at], operand[xas_at + 1]]) as i32;
        let tbd_at = tbd_base + i as usize;
        if tbd_at >= operand.len() {
            break;
        }
        tabs.push(tab_from_tbd(dxp, operand[tbd_at]));
    }
    tabs
}

/// Build a `TabStop` from a twips position and a 1-byte `TBD` descriptor
/// (`jc` in bits 0..2, `tlc` in bits 3..5).
fn tab_from_tbd(position_twips: i32, tbd: u8) -> TabStop {
    let jc = tbd & 0x7;
    let tlc = (tbd >> 3) & 0x7;
    TabStop {
        position_twips,
        alignment: match jc {
            0 => TabAlignment::Left,
            1 => TabAlignment::Center,
            2 => TabAlignment::Right,
            3 => TabAlignment::Decimal,
            4 => TabAlignment::Bar,
            _ => TabAlignment::Left,
        },
        leader: match tlc {
            0 => TabLeader::None,
            1 => TabLeader::Dot,
            2 => TabLeader::Hyphen,
            3 => TabLeader::Underscore,
            4 => TabLeader::Heavy,
            5 => TabLeader::MiddleDot,
            _ => TabLeader::None,
        },
    }
}

/// Decode a PAP `grpprl` into the paragraph flags we care about.
///
/// Unknown SPRMs are ignored. An empty `grpprl` yields the default
/// (all-false) `PapProps`, which classifies the paragraph as ordinary prose.
pub fn extract_pap_props(grpprl: &[u8]) -> PapProps {
    let mut props = PapProps::default();

    for sprm in parse_grpprl(grpprl) {
        match sprm.opcode {
            // sprmPFInTable — 1-byte operand, bit 0 = fInTable.
            0x2416 => {
                if let Some(&b) = sprm.operand.first() {
                    props.f_in_table = (b & 1) != 0;
                }
            },
            // sprmPItap — 4-byte operand, the table nesting depth.
            0x6649 => {
                if sprm.operand.len() >= 4 {
                    let v = u32::from_le_bytes([
                        sprm.operand[0],
                        sprm.operand[1],
                        sprm.operand[2],
                        sprm.operand[3],
                    ]);
                    props.itap = v as u8;
                }
            },
            // sprmTDefTable — presence marks a row-terminator paragraph;
            // the operand carries the row definition (cells, boundaries).
            0xD608 => {
                props.is_table_trailing_mark = true;
                props.tap = parse_tdef_table(&sprm.operand);
            },
            // sprmPIlfo (0x460B) — 2-byte operand read as *signed* `i16` per
            // [MS-DOC] §2.9.150: `0x0000`/`0xF801` mean "not in a list",
            // `0x0001`–`0x07FE` are 1-based indices into `PlfLfo.rgLfo`, and
            // `0xF802`–`0xFFFF` are the negation of a 1-based index (still in a
            // list). Storing it signed keeps the negation explicit.
            0x460B => {
                if sprm.operand.len() >= 2 {
                    props.ilfo = Some(i16::from_le_bytes([sprm.operand[0], sprm.operand[1]]));
                }
            },
            // sprmPIlvl (0x260A) — 1-byte operand, the list level (0-based).
            0x260A => {
                if let Some(&b) = sprm.operand.first() {
                    props.ilvl = Some(b);
                }
            },
            // sprmPStyle (0x640A) — the 4-byte operand's low word is the style
            // index (`istd`) this paragraph is restyled to. Read here rather
            // than in a separate pass so the grpprl is walked only once; it
            // overrides the PAPX header `istd` when the style is resolved.
            0x640A => {
                if sprm.operand.len() >= 2 {
                    props.style_istd = Some(u16::from_le_bytes([sprm.operand[0], sprm.operand[1]]));
                }
            },
            // sprmPOutlineLvl (0x6412) — the 4-byte operand's low byte is the
            // outline level: `0` = body text, `1`–`9` = Heading 1–9 (MS-DOC
            // §2.9.138). This is the authoritative heading marker and wins over
            // any style-derived level resolved later in `build_paragraphs`.
            //
            // Use the raw low byte and validate the `1..=9` range. A coincidental
            // byte pattern that merely *contains* the opcode must not be turned
            // into a heading: e.g. an operand low byte of 0x68 (104) is not a
            // valid outline level and must be rejected, not masked to `8`.
            0x6412 => {
                if let Some(&b) = sprm.operand.first() {
                    // A valid outline level is 0–MAX_OUTLINE_LEVEL; anything else
                    // is a coincidental byte pattern, not an outline level, and
                    // must leave the paragraph's status untouched.
                    if b <= MAX_OUTLINE_LEVEL {
                        // The level is now settled by direct formatting: 0 is an
                        // explicit body-text marker, 1–MAX_OUTLINE_LEVEL a heading.
                        props.outline_lvl_explicit = true;
                        if b >= 1 {
                            props.heading_level = Some(b);
                        }
                    }
                }
            },
            // sprmPChgTabs (0xC615) / sprmPChgTabsPapx (0xC60D): tab stops.
            0xC615 | 0xC60D => {
                if !sprm.operand.is_empty() {
                    props.tabs = decode_pchg_tabs(sprm.opcode, &sprm.operand);
                }
            },
            _ => {},
        }
    }

    props
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Cell-paragraph grpprl: `sprmPFInTable(0x2416)=1`, `sprmPItap(0x6649)=1`.
    fn cell_grpprl() -> Vec<u8> {
        vec![0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00]
    }

    /// Row-terminator grpprl: `sprmPFInTable(0x2416)=1`,
    /// `sprmPFInTableTtp(0x2417)=1`, `sprmPItap(0x6649)=1`, then the full TAP
    /// including `sprmTDefTable(0xD608)`. Only the head is needed to verify
    /// flag extraction; the rest is padding the walker must skip over.
    fn row_mark_grpprl() -> Vec<u8> {
        // 0x2416 op=01 | 0x2417 op=01 | 0x6649 op=01000000 | 0x563a op=1400
        // | 0xd634 1-byte len=06 ... | 0xd608 2-byte cb=04 payload=000000
        vec![
            0x16, 0x24, 0x01, // sprmPFInTable = 1
            0x17, 0x24, 0x01, // sprmPFInTableTtp = 1
            0x49, 0x66, 0x01, 0x00, 0x00, 0x00, // sprmPItap = 1
            0x3a, 0x56, 0x14, 0x00, // sprmTDefTableSpacing? 2-byte op
            0x34, 0xd6, 0x06, 0x00, 0x01, 0x02, 0x03, 0x6c, 0x00, // 1-byte len
            0x08, 0xd6, 0x04, 0x00, 0x00, 0x00, 0x00, // sprmTDefTable, 2-byte cb=4
        ]
    }

    #[test]
    fn walks_fixed_and_variable_sprms() {
        let sprms = parse_grpprl(&cell_grpprl());
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].opcode, 0x2416);
        assert_eq!(sprms[0].operand, vec![0x01]);
        assert_eq!(sprms[1].opcode, 0x6649);
        assert_eq!(sprms[1].operand, vec![0x01, 0x00, 0x00, 0x00]);
    }

    /// `sprmPOutlineLvl` (0x6412) is a spra-3 (4-byte) SPRM whose low byte is the
    /// outline level: `1`–`9` = Heading 1–9. The level must surface on
    /// `PapProps.heading_level`.
    #[test]
    fn sprm_p_outline_lvl_sets_heading_level() {
        // opcode 0x6412 (LE) + 4-byte operand, low byte = 2 (valid level).
        let grpprl = vec![0x12, 0x64, 0x02, 0x00, 0x00, 0x00];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.heading_level, Some(2), "outline level 2 -> Heading 2");
        assert!(props.outline_lvl_explicit, "a valid level settles the outline level");
    }

    #[test]
    fn sprm_p_outline_lvl_body_is_none() {
        // outline level 0 means body text, not a heading — and it is *explicit*,
        // so `resolve_heading_level` must not fall back to the style.
        let grpprl = vec![0x12, 0x64, 0x00, 0x00, 0x00, 0x00];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.heading_level, None, "outline level 0 -> not a heading");
        assert!(
            props.outline_lvl_explicit,
            "level 0 is an explicit body-text marker, not an absent SPRM"
        );
    }

    /// A `sprmPOutlineLvl` operand whose low byte is outside 1–9 (here 0x68 =
    /// 104) is not a valid outline level and must be ignored, even though the
    /// opcode byte pattern appears in the grpprl. This guards against
    /// coincidental byte sequences fabricating spurious headings.
    #[test]
    fn sprm_p_outline_lvl_rejects_invalid_byte() {
        let grpprl = vec![0x12, 0x64, 0x68, 0x01, 0x01, 0x00];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.heading_level, None, "0x68 is not a valid outline level");
        assert!(
            !props.outline_lvl_explicit,
            "an invalid operand is not an outline level at all, so the status stays unset"
        );
    }

    /// `sprmPStyle` (0x640A) overrides the PAPX `istd`; the 2-byte istd is the
    /// low word of the (spra-3, 4-byte) operand. It surfaces on
    /// `PapProps.style_istd` from the same grpprl walk that decodes the other
    /// paragraph properties.
    #[test]
    fn pap_props_read_sprm_p_style_override() {
        // opcode 0x640A (LE) + 4-byte operand, low word = istd 5.
        let grpprl = vec![0x0A, 0x64, 0x05, 0x00, 0x00, 0x00];
        assert_eq!(extract_pap_props(&grpprl).style_istd, Some(5));
    }

    #[test]
    fn pap_props_style_istd_absent_when_no_sprm() {
        assert_eq!(extract_pap_props(&[0x16, 0x24, 0x01]).style_istd, None);
    }

    #[test]
    fn row_mark_walk_consumes_every_byte() {
        let grpprl = row_mark_grpprl();
        let sprms = parse_grpprl(&grpprl);
        // No byte left behind: re-encoding the walked SPRMs — re-inserting the
        // correct length prefix (1-byte for ordinary spra=6 SPRMs, 2-byte `cb`
        // for the sole exception 0xD608) — reproduces the input exactly. This
        // guards against off-by-one skipping.
        let mut rebuilt = Vec::new();
        for s in &sprms {
            rebuilt.push((s.opcode & 0xFF) as u8);
            rebuilt.push((s.opcode >> 8) as u8);
            if s.opcode == 0xD608 {
                let cb = (s.operand.len() + 1) as u16;
                rebuilt.push(cb as u8);
                rebuilt.push((cb >> 8) as u8);
            } else if s.spra() == 6 {
                rebuilt.push(s.operand.len() as u8);
            }
            rebuilt.extend_from_slice(&s.operand);
        }
        assert_eq!(rebuilt, grpprl, "walker must consume every byte exactly");
        // The 0xD608 must be reached.
        assert!(sprms.iter().any(|s| s.opcode == 0xD608));
    }

    /// Regression for the `sprmTDefTable` 2-byte `cb` length prefix.
    ///
    /// For `cb >= 256` the 1-byte-prefix walk under-consumes by `256 × cb_high`
    /// and decodes the rest of the row's grpprl as fabricated SPRMs, which
    /// both drops the merged-cell spans (TAP fails to parse) and can collide
    /// with structure-driving opcodes. A 12-column table is the threshold:
    /// `cb = 4 + 22·itcMac = 268` (>= 256). The walker must read the 2-byte
    /// `cb`, decode the full TAP, and still reach the trailing SPRM.
    #[test]
    fn d608_two_byte_cb_decodes_wide_tables() {
        let itc: usize = 12;
        let mut tap = vec![itc as u8]; // itcMac
        for _ in 0..=itc {
            tap.extend_from_slice(&0i16.to_le_bytes()); // rgdxaCenter
        }
        for _ in 0..itc {
            tap.extend_from_slice(&[0u8; 20]); // TKBKTAP descriptors
        }
        let cb = (tap.len() + 1) as u16; // operand length is cb - 1
        let mut grpprl = vec![0x08, 0xd6, cb as u8, (cb >> 8) as u8];
        grpprl.extend_from_slice(&tap);
        // Trailing distinct SPRM proves the walker didn't under-consume.
        grpprl.extend_from_slice(&[0x16, 0x24, 0x01]); // sprmPFInTable = 1

        let sprms = parse_grpprl(&grpprl);
        let td = sprms
            .iter()
            .find(|s| s.opcode == 0xD608)
            .expect("0xD608 must be present");
        assert_eq!(
            td.operand.len(),
            tap.len(),
            "2-byte cb must decode the full {}-byte TAP",
            tap.len()
        );

        let props = extract_pap_props(&grpprl);
        assert_eq!(props.tap.as_ref().unwrap().itc_mac, 12, "12-column table TAP must parse");
        assert!(
            sprms.iter().any(|s| s.opcode == 0x2416),
            "trailing SPRM must be reached (no under-consumption)"
        );
    }

    #[test]
    fn extracts_cell_props() {
        let props = extract_pap_props(&cell_grpprl());
        assert!(props.f_in_table);
        assert!(!props.is_table_trailing_mark);
        assert_eq!(props.itap, 1);
    }

    #[test]
    fn extracts_row_mark_props() {
        let props = extract_pap_props(&row_mark_grpprl());
        assert!(props.f_in_table);
        assert!(props.is_table_trailing_mark);
        assert_eq!(props.itap, 1);
        assert!(props.tap.is_some(), "0xD608 operand must parse into a TapInfo");
    }

    #[test]
    fn empty_grpprl_is_ordinary_prose() {
        let props = extract_pap_props(&[]);
        assert!(!props.f_in_table);
        assert!(!props.is_table_trailing_mark);
        assert_eq!(props.itap, 0);
        assert!(props.ilvl.is_none());
        assert!(props.ilfo.is_none());
        assert!(props.tap.is_none());
    }

    #[test]
    fn decodes_pchg_tabs_new_stops() {
        // PChgTabsOperand (0xC615): PchgTabsDelClose (cDel=0) then PchgTabsAdd
        // (cAdd=2). Positions are 2-byte XAS (signed twips); each TBD is 1
        // byte with `jc` in bits 0..2. new[0]: jc=2 (Right), pos=2000;
        // new[1]: jc=1 (Center), pos=1000.
        let mut operand = vec![0x00]; // cDel = 0 (no deletes)
        operand.push(2); // cAdd = 2
        operand.extend_from_slice(&[0xD0, 0x07]); // rgdxaAdd[0] = 2000
        operand.extend_from_slice(&[0xE8, 0x03]); // rgdxaAdd[1] = 1000
        operand.push(0x02); // rgtbdAdd[0]: jc=2 (Right)
        operand.push(0x01); // rgtbdAdd[1]: jc=1 (Center)

        let tabs = decode_pchg_tabs(0xC615, &operand);
        assert_eq!(tabs.len(), 2);
        assert_eq!(tabs[0].position_twips, 2000);
        assert_eq!(tabs[0].alignment, TabAlignment::Right);
        assert_eq!(tabs[1].position_twips, 1000);
        assert_eq!(tabs[1].alignment, TabAlignment::Center);
    }

    #[test]
    fn pchg_tabs_malformed_operand_is_empty() {
        // Truncated operand: cDel=2 but no rgdxaDel/rgdxaClose bytes, so the
        // Add list is unreachable — must degrade to empty, not panic.
        assert!(decode_pchg_tabs(0xC615, &[0x02]).is_empty());
        // PchgTabsPapx: cDel=2 but no following bytes.
        assert!(decode_pchg_tabs(0xC60D, &[0x02, 0x00]).is_empty());
    }

    fn hex(s: &str) -> Vec<u8> {
        (0..s.len())
            .step_by(2)
            .map(|i| u8::from_str_radix(&s[i..i + 2], 16).unwrap())
            .collect()
    }

    #[test]
    fn parses_tdef_table_boundaries() {
        // 0xD608 operand bytes captured from a real Word document (the source
        // .doc is not distributed in this repo): itcMac=2, boundaries
        // [0, 6872, 9302]. (The 2-byte `cb` prefix is stripped by parse_grpprl,
        // so the operand here starts at itcMac.)
        let tap = parse_tdef_table(&hex(
            "020000d81a562400000000040101000401010004010100000000000000000004010100040101000401010004010100",
        ))
        .unwrap();
        assert_eq!(tap.itc_mac, 2);
        assert_eq!(tap.centers, vec![0, 6872, 9302]);
        assert_eq!(tap.cells.len(), 2);
        assert_eq!(tap.cells[0].rgf, 0);
        assert_eq!(tap.cells[0].w_width, 0);
    }

    #[test]
    fn parses_tdef_table_vertical_merge_flags() {
        // 0xD608 operand bytes captured from a real Word document (source .doc
        // not distributed): itcMac=4; the first cell descriptor carries
        // fVertMerge | fVertRestart (0x0060). (cb prefix stripped — operand
        // starts at itcMac.)
        let tap = parse_tdef_table(&hex(
            "04000026046a16d41f56246000000004010100040101000401010000000000000000000401010004010100040101000000000000000000040101000401010004010100000000000000000004010100040101000401010004010100",
        ))
        .unwrap();
        assert_eq!(tap.itc_mac, 4);
        assert_eq!(tap.centers, vec![0, 1062, 5738, 8148, 9302]);
        assert_eq!(tap.cells[0].rgf, 0x0060, "fVertMerge | fVertRestart");
        assert_eq!(tap.cells[1].rgf, 0);
        assert_eq!(tap.cells[3].rgf, 0);
    }

    #[test]
    fn rejects_truncated_tdef_table() {
        assert!(parse_tdef_table(&[]).is_none());
        // itcMac=4 but only 2 further bytes — needs far more for rgdxaCenter
        // + rgtc, so it must fail.
        assert!(parse_tdef_table(&[0x04, 0x00, 0x00]).is_none());
        // A TAP labelled as itcMac=4 needs 92 bytes and must fail when short.
        let row0 = hex(
            "020000d81a562400000000040101000401010004010100000000000000000004010100040101000401010004010100",
        );
        let mut malformed = vec![0x04];
        malformed.extend_from_slice(&row0[1..]);
        malformed.truncate(10); // far fewer than the 92 bytes required
        assert!(parse_tdef_table(&malformed).is_none());
    }

    #[test]
    fn spra_and_sgc_fields() {
        let cell = Sprm {
            opcode: 0x2416,
            operand: vec![1],
        };
        assert_eq!(cell.spra(), 1); // 1-byte operand
        assert_eq!(cell.sgc(), 1); // PAP

        let tdef = Sprm {
            opcode: 0xD608,
            operand: vec![],
        };
        assert_eq!(tdef.spra(), 6); // variable
        assert_eq!(tdef.sgc(), 5); // TAP
    }

    // --------------------------------------------------------------------
    // Regression tests for the opcode-identity defects in `extract_pap_props`
    // (PR #116 blind review). Every fixture uses the opcode *as Word writes it
    // per [MS-DOC]*, not the project's own constants, so the tests fail while
    // the decoder mislabels opcodes and turn green once dispatch is corrected.
    // --------------------------------------------------------------------

    /// `0x460B` is `sprmPIlfo` (2-byte operand). The decoder must populate
    /// `ilfo`. Today it is dispatched as `sprmPIlvl`, so `ilfo` stays `None`.
    #[test]
    fn sprm_pilfo_opcode_460b_populates_ilfo() {
        // sprmPIlfo (0x460B), operand = 0x0005 (ilfo index 5).
        let grpprl = [0x0B, 0x46, 0x05, 0x00];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.ilfo, Some(5), "0x460B is sprmPIlfo: must set ilfo");
        assert_eq!(props.ilvl, None, "0x460B must not be read as ilvl");
    }

    /// `0x260A` is `sprmPIlvl` (1-byte operand). The decoder must populate
    /// `ilvl`. Today it is never read (falls through to `_`).
    #[test]
    fn sprm_pilvl_opcode_260a_populates_ilvl() {
        // sprmPIlvl (0x260A), operand = 0x01 (level 1).
        let grpprl = [0x0A, 0x26, 0x01];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.ilvl, Some(1), "0x260A is sprmPIlvl: must set ilvl");
    }

    /// `0xC615` is `sprmPChgTabs`. The decoder must populate tab stops from it.
    #[test]
    fn sprm_pchg_tabs_opcode_c615_populates_tabs() {
        // Per [MS-DOC] §2.9.182 the operand is a `PChgTabsOperand`:
        // `PchgTabsDelClose` (cDel, then 4 bytes per delete) followed by
        // `PchgTabsAdd` (cAdd, then 2-byte positions + 1-byte TBD per add).
        // `0xC615` carries a **1-byte** `cb` (not the 2-byte `cb` of 0xD608), so
        // the grpprl layout is opcode(2) + cb(1) + body(cb bytes).
        //
        // The fixture uses cDel = 1 (one delete entry) on purpose: a 1-byte-cb
        // off-by-one that shifted the operand by one byte would misread `cDel`
        // and land the add list at the wrong offset, so only a non-zero `cDel`
        // exposes the bug. cAdd = 2; positions 2000 & 1000, TBD jc=2 / jc=1.
        let body: Vec<u8> = vec![
            1, // cDel = 1 (one delete entry — exercises the skip)
            0x00, 0x00, // rgdxaDel[0]
            0x00, 0x00, // rgdxaClose[0]
            2,    // cAdd = 2
            0xD0, 0x07, // rgdxaAdd[0] = 2000
            0xE8, 0x03, // rgdxaAdd[1] = 1000
            0x02, // rgtbdAdd[0]: jc=2 (Right)
            0x01, // rgtbdAdd[1]: jc=1 (Center)
        ];
        let cb = body.len() as u8; // 1-byte cb = body length
        let mut grpprl = vec![0x15, 0xC6]; // sprmPChgTabs (0xC615)
        grpprl.push(cb);
        grpprl.extend_from_slice(&body);

        let props = extract_pap_props(&grpprl);
        assert!(!props.tabs.is_empty(), "0xC615 is sprmPChgTabs: must populate tabs");
        assert_eq!(props.tabs.len(), 2, "cDel must be skipped, cAdd=2 adds remain");
        assert_eq!(props.tabs[0].position_twips, 2000);
        assert_eq!(props.tabs[0].alignment, TabAlignment::Right);
        assert_eq!(props.tabs[1].position_twips, 1000);
        assert_eq!(props.tabs[1].alignment, TabAlignment::Center);
    }

    /// `0xC615` in the `cb == 255` escape form: the literal byte count is not
    /// 255, the length is derived from the payload's own `cDel`/`cAdd`. A naive
    /// "255-byte operand" read would overrun and desync the rest of the grpprl;
    /// the byte-level round trip must reproduce the input exactly.
    #[test]
    fn sprm_pchg_tabs_c615_cb_255_escape_round_trips() {
        // cDel = 1 (delete block = 1 + 4*1 = 5 bytes), cAdd = 2 (add block =
        // 1 + 2*2 + 2 = 7 bytes). Total body = 12 bytes.
        let body: Vec<u8> = vec![
            1, 0x00, 0x00, 0x00, 0x00, // PchgTabsDelClose: cDel=1 + 4 bytes
            2, 0x64, 0x00, 0xC8, 0x00, 0x03, 0x01, // PchgTabsAdd: cAdd=2 + positions + TBDs
        ];
        let mut grpprl = vec![0x15, 0xC6, 0xFF]; // opcode + cb == 255 escape
        grpprl.extend_from_slice(&body);

        let sprms = parse_grpprl(&grpprl);
        assert_eq!(sprms.len(), 1, "must decode exactly one SPRM");
        let s = &sprms[0];
        assert_eq!(s.opcode, 0xC615);
        // The 1-byte cb (0xFF) is consumed; operand is the raw body.
        assert_eq!(s.operand, body, "operand must be the body without the cb");

        // Round-trip: rebuild the grpprl from the walked SPRM.
        let rebuilt = {
            let mut v = vec![(s.opcode & 0xFF) as u8, (s.opcode >> 8) as u8];
            v.push(0xFF); // cb escape
            v.extend_from_slice(&s.operand);
            v
        };
        assert_eq!(rebuilt, grpprl, "walker must consume every byte exactly");

        let props = extract_pap_props(&grpprl);
        assert_eq!(props.tabs.len(), 2);
        assert_eq!(props.tabs[0].position_twips, 100); // 0x64
        assert_eq!(props.tabs[1].position_twips, 200); // 0xC8
    }

    /// `0xC60D` (`sprmPChgTabsPapx`) carries a normal 1-byte `cb` prefix (not a
    /// no-prefix opcode), and its delete block is `PChgTabsDel` (`1 + 2·cDel`,
    /// one XAS per tab) rather than `PchgTabsDelClose` (`1 + 4·cDel`). Build the
    /// grpprl straight from [MS-DOC]: opcode + `cb = 7` + `PChgTabsDel{cTabs=1,
    /// rgdxaDel=[16]}` + `PChgTabsAdd{cTabs=1, pos=2000, TBD jc=2}`, followed by
    /// a `sprmPFInTable` so we can prove the walker does NOT swallow it.
    #[test]
    fn sprm_pchg_tabs_papx_c60d_populates_tabs() {
        // PchgTabsDel: cTabs=1, rgdxaDel=[16] (2-byte XAS) -> 3 bytes.
        // PchgTabsAdd: cTabs=1, rgdxaAdd=[2000], rgtbdAdd=[jc=2] -> 4 bytes.
        // Body = 7 bytes, so the 1-byte cb prefix is 7.
        let body: Vec<u8> = vec![
            1, 0x10, 0x00, // PchgTabsDel: cTabs=1, rgdxaDel=[16]
            1, 0xD0, 0x07, 0x02, // PchgTabsAdd: cTabs=1, rgdxaAdd=[2000], TBD jc=2
        ];
        let mut grpprl = vec![0x0D, 0xC6, 7]; // opcode + 1-byte cb
        grpprl.extend_from_slice(&body);
        grpprl.extend_from_slice(&[0x16, 0x24, 0x01]); // sprmPFInTable, operand 0x01

        // The trailing SPRM must be decoded, not swallowed by a bad length.
        let sprms = parse_grpprl(&grpprl);
        assert_eq!(sprms.len(), 2, "0xC60D must decode AND leave the trailing SPRM intact");
        assert_eq!(sprms[0].opcode, 0xC60D);
        assert_eq!(sprms[0].operand, body);
        assert_eq!(sprms[1].opcode, 0x2416);

        let props = extract_pap_props(&grpprl);
        assert_eq!(props.tabs.len(), 1, "0xC60D PChgTabsDel is 1 + 2·cDel");
        assert_eq!(props.tabs[0].position_twips, 2000);
        assert_eq!(props.tabs[0].alignment, TabAlignment::Right);
        assert!(props.f_in_table, "trailing sprmPFInTable must be reached and applied");
    }

    /// `0xD632` is `sprmTCellPadding` (NOT `sprmPChgTabs`). It must not be read
    /// as tab stops. Today it is dispatched as `sprmPChgTabs`, so `tabs` is
    /// populated — the inverse of the correct behaviour.
    #[test]
    fn sprm_tcell_padding_opcode_d632_does_not_populate_tabs() {
        // PChgTabsOperand-style bytes tagged with the TCellPadding opcode
        // (0xD632): 1-byte length prefix = 8, then cDel=0, cAdd=2, two
        // positions, two TBDs.
        let grpprl = [
            0x32, 0xD6, 8, 0x00, 0x02, 0xD0, 0x07, 0xE8, 0x03, 0x02, 0x01,
        ];
        let props = extract_pap_props(&grpprl);
        assert!(
            props.tabs.is_empty(),
            "0xD632 is sprmTCellPadding, not sprmPChgTabs: must not populate tabs"
        );
    }

    /// `0xC615` with a truncated length prefix must stop cleanly, never panic
    /// (AGENTS.md rule 6). A grpprl holding only the opcode (no cb byte) and one
    /// holding the cb but no body are both malformed inputs.
    #[test]
    fn sprm_pchg_tabs_c615_truncated_cb_is_empty() {
        // Opcode only, no cb byte: the 1-byte cb read is out of bounds -> stop.
        let sprms = parse_grpprl(&[0x15, 0xC6]);
        assert!(sprms.is_empty(), "truncated 0xC615 (no cb) must yield no SPRM, not panic");
        // cb present but body absent: cb == 4 claims 4 body bytes that do not
        // exist; the operand must clamp to empty, not read past the buffer.
        let sprms = parse_grpprl(&[0x15, 0xC6, 0x04]);
        assert_eq!(sprms.len(), 1, "opcode is present so one SPRM is produced");
        assert!(
            sprms[0].operand.is_empty(),
            "0xC615 cb with no body must clamp the operand to empty"
        );
    }

    /// `0xC615` in the `cb == 255` escape form must be consumed exactly so the
    /// following SPRM is reached. This is the exact off-by-one the 2-byte-cb
    /// misreading caused: a one-byte shift would desync the rest of the grpprl
    /// and either drop or mis-parse the trailing SPRM.
    #[test]
    fn sprm_pchg_tabs_c615_255_escape_followed_by_sprm() {
        // cDel=1, cAdd=2 (12-byte body), then a trailing sprmPFInTable
        // (0x2416, 1-byte operand 0x01).
        let body: Vec<u8> = vec![
            1, 0x00, 0x00, 0x00, 0x00, // PchgTabsDelClose: cDel=1 + 4 bytes
            2, 0x64, 0x00, 0xC8, 0x00, 0x03, 0x01, // PchgTabsAdd: cAdd=2 + positions + TBDs
        ];
        let mut grpprl = vec![0x15, 0xC6, 0xFF]; // opcode + cb == 255 escape
        grpprl.extend_from_slice(&body);
        grpprl.extend_from_slice(&[0x16, 0x24, 0x01]); // sprmPFInTable, operand 0x01

        let sprms = parse_grpprl(&grpprl);
        assert_eq!(sprms.len(), 2, "must decode both the 0xC615 and the trailing SPRM (no desync)");
        assert_eq!(sprms[0].opcode, 0xC615);
        assert_eq!(sprms[0].operand, body, "0xC615 operand must be the raw body");
        assert_eq!(sprms[1].opcode, 0x2416);
        assert_eq!(sprms[1].operand, vec![0x01]);

        let props = extract_pap_props(&grpprl);
        assert_eq!(props.tabs.len(), 2, "tabs from 0xC615 must be present");
        assert!(props.f_in_table, "trailing sprmPFInTable must be reached and applied");
    }

    /// A variable-length SPRM whose length prefix / body is truncated must stop
    /// the walk cleanly, never panic (AGENTS.md rule 6). Covers the truncation
    /// `break` for each special variable encoding: 0xD608 (2-byte cb), 0xC615
    /// (1-byte cb), and 0xC60D (1-byte cb).
    #[test]
    fn parse_grpprl_truncated_variable_sprm_prefixes() {
        assert!(
            parse_grpprl(&[0x08, 0xD6]).is_empty(),
            "0xD608 with no 2-byte cb must stop, not panic"
        );
        assert!(
            parse_grpprl(&[0x15, 0xC6]).is_empty(),
            "0xC615 with no 1-byte cb must stop, not panic"
        );
        assert!(
            parse_grpprl(&[0x0D, 0xC6]).is_empty(),
            "0xC60D with no body must stop, not panic"
        );
    }

    /// `pchg_tabs_operand_len` must bound itself against a short buffer instead
    /// of indexing out of range (AGENTS.md rule 6).
    #[test]
    fn pchg_tabs_operand_len_truncated() {
        // start beyond the buffer -> 0.
        assert_eq!(pchg_tabs_operand_len(&[], 0), 0);
        // cDel present but its rgdxa/rgdxaClose block runs past the end -> the
        // remaining bytes are returned, not a panic.
        assert_eq!(pchg_tabs_operand_len(&[0x02], 0), 1);
    }

    /// Consolidated opcode-conformance gate: every dispatched opcode must name
    /// the property [MS-DOC] assigns it. Fails today because the decoder
    /// misroutes `0x460B` (as `ilvl`) and `0xD632` (as tabs) and never reads
    /// `0x260A` / `0xC615`.
    #[test]
    fn opcode_conformance_gate() {
        // 0x460B = sprmPIlfo -> ilfo
        assert_eq!(
            extract_pap_props(&[0x0B, 0x46, 0x03, 0x00]).ilfo,
            Some(3),
            "0x460B = sprmPIlfo"
        );
        // 0x260A = sprmPIlvl -> ilvl
        assert_eq!(extract_pap_props(&[0x0A, 0x26, 0x02]).ilvl, Some(2), "0x260A = sprmPIlvl");
        // 0xD632 = sprmTCellPadding -> no tabs
        assert!(
            extract_pap_props(&[0x32, 0xD6, 4, 0x00, 0x01, 0x64, 0x00])
                .tabs
                .is_empty(),
            "0xD632 = sprmTCellPadding"
        );
    }
}
