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

use crate::ir::{
    LineSpacing, ParagraphAlignment, TabAlignment, TabLeader, TabStop, UnderlineStyle,
    VerticalAlign,
};

/// The `sprmPOutLvl` (0x2640) operand value meaning "this paragraph is body
/// text". The SPRM encodes the level zero-based, so 0x00–0x08 are Heading 1–9
/// and this value — not 0 — is the body-text marker ([MS-DOC] §2.6.2).
const OUTLVL_BODY_TEXT: u8 = 9;

/// A single decoded SPRM: opcode plus its operand bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sprm {
    /// The 2-byte SPRM opcode.
    pub opcode: u16,
    /// The operand bytes (may be empty for a zero-length variable SPRM).
    pub operand: Vec<u8>,
}

#[cfg(test)]
impl Sprm {
    /// `spra` — operand-size class (bits 15..13 of the opcode).
    pub fn spra(&self) -> u8 {
        ((self.opcode >> 13) & 0x7) as u8
    }

    /// `sgc` — property class (bits 12..10): `1` = PAP, `5` = TAP, …
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
    while let Some((sprm, next)) = next_sprm(grpprl, p) {
        out.push(sprm);
        p = next;
    }
    out
}

/// Decode the SPRM starting at `p`, returning it and the offset of the next
/// one (which may lie past the end when the operand is truncated). `None`
/// when fewer than two bytes remain or a length prefix is cut off.
fn next_sprm(grpprl: &[u8], p: usize) -> Option<(Sprm, usize)> {
    let len = grpprl.len();
    if p + 2 > len {
        return None;
    }
    {
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
                        return None;
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
                        return None;
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
                        return None;
                    }
                    let n = grpprl[p + 2] as usize;
                    let start = p + 3;
                    let end = (start + n).min(len);
                    (grpprl[start..end].to_vec(), start + n)
                }
            },
        };
        Some((Sprm { opcode, operand }, next))
    }
}

/// `sprmPHugePapx` ([MS-DOC] §2.6.2, `0x6646`): the paragraph's property
/// list is too large for the PAPX and lives in the `Data` stream. The
/// 4-byte operand is the `fc` of a `PrcData` ([MS-DOC] §2.9.210) there.
pub const SPRM_P_HUGE_PAPX: u16 = 0x6646;

/// The same sprm under its Word 97 number. [MS-DOC] §2.6.2 lists only
/// `0x6646`; Apache POI (`PAPX.findHuge`) and LibreOffice (`LN_PHugePapx`)
/// both also accept `0x6645` with the same 4-byte `fc` operand.
const SPRM_P_HUGE_PAPX_WORD97: u16 = 0x6645;

/// `PrcData.cbGrpprl` is at most `0x3FA2` ([MS-DOC] §2.9.210); anything
/// larger or negative is not a valid `PrcData`.
const PRC_DATA_MAX_CB: usize = 0x3FA2;

/// The allowance for copying `PrcData` out of the `Data` stream, shared by all
/// the paragraphs of one document.
///
/// Every copy is charged its full length. Paragraphs whose property lists are
/// byte for byte the same (the row ends of a long table, as a rule) are
/// expanded once and share the result, so they spend the allowance once; that
/// sharing is done by the caller, not here. Two property lists that merely name
/// the same `PrcData` each get a buffer of their own, and each is charged: a
/// budget that skipped repeated offsets would let a few bytes per paragraph
/// copy out 16 KB each.
#[derive(Debug, Default)]
pub struct HugeBudget {
    remaining: usize,
}

impl HugeBudget {
    /// A budget of `remaining` bytes.
    pub fn new(remaining: usize) -> Self {
        Self { remaining }
    }

    /// Whether a `PrcData` of `len` bytes may be copied, charging it.
    fn admit(&mut self, len: usize) -> bool {
        if len > self.remaining {
            return false;
        }
        self.remaining -= len;
        true
    }
}

/// Replace each `sprmPHugePapx` in `grpprl` with the property list its
/// `PrcData` holds in the `Data` stream, in place so SPRM order is kept.
///
/// Returns `None` when `grpprl` has no `sprmPHugePapx`. Otherwise returns the
/// expanded list and how many references could not be resolved (an operand
/// that is not 4 bytes, an `fc` outside `data`, a negative or oversized
/// `cbGrpprl`, or no `budget` left); an unresolved reference is dropped from
/// the list.
///
/// `budget` bounds the bytes copied out of `data` across all calls, charging
/// every copy (see [`HugeBudget`]). A 6-byte reference
/// expands to up to 16 KB, so without it a small file multiplies into
/// gigabytes.
///
/// The expansion is one level deep: a `sprmPHugePapx` inside a `PrcData` is not
/// followed, so a self-referencing `Data` stream cannot loop.
pub fn expand_huge_papx(
    grpprl: &[u8],
    data: &[u8],
    budget: &mut HugeBudget,
) -> Option<(Vec<u8>, usize)> {
    // Almost no paragraph has one; skip the SPRM walk unless an opcode's
    // bytes appear at all.
    let has = |op: u16| grpprl.windows(2).any(|w| w == op.to_le_bytes());
    if !has(SPRM_P_HUGE_PAPX) && !has(SPRM_P_HUGE_PAPX_WORD97) {
        return None;
    }
    let mut out: Option<Vec<u8>> = None;
    let mut unresolved = 0usize;
    let mut copied = 0usize;
    let mut p = 0usize;
    while let Some((sprm, next)) = next_sprm(grpprl, p) {
        if sprm.opcode == SPRM_P_HUGE_PAPX || sprm.opcode == SPRM_P_HUGE_PAPX_WORD97 {
            let buf = out.get_or_insert_with(Vec::new);
            buf.extend_from_slice(&grpprl[copied..p]);
            copied = next.min(grpprl.len());
            match prc_data_grpprl(&sprm.operand, data).filter(|&(_, g)| budget.admit(g.len())) {
                Some((_, g)) => buf.extend_from_slice(g),
                None => unresolved += 1,
            }
        }
        p = next;
    }
    out.map(|mut buf| {
        buf.extend_from_slice(&grpprl[copied..]);
        (buf, unresolved)
    })
}

/// The `fc` and `GrpPrl` of the `PrcData` that a `sprmPHugePapx` operand
/// points at.
fn prc_data_grpprl<'a>(operand: &[u8], data: &'a [u8]) -> Option<(usize, &'a [u8])> {
    let fc = u32::from_le_bytes(operand.get(..4)?.try_into().ok()?) as usize;
    let cb = i16::from_le_bytes(data.get(fc..fc.checked_add(2)?)?.try_into().ok()?);
    let cb = usize::try_from(cb).ok().filter(|&c| c <= PRC_DATA_MAX_CB)?;
    let start = fc + 2;
    Some((fc, data.get(start..start.checked_add(cb)?)?))
}

/// Paragraph-property flags distilled from a PAP grpprl.
///
/// Table/list reconstruction fields plus
/// alignment/indentation/spacing.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct PapProps {
    /// `sprmPFInTable` (0x2416): this paragraph lives inside a table.
    pub f_in_table: bool,
    /// The paragraph carries a TAP (table row definition, `sprmTDefTable`
    /// 0xD608). Such a paragraph is the row-terminator ("table trailing
    /// paragraph") and is NOT itself a cell.
    pub is_table_trailing_mark: bool,
    /// `sprmTTableHeader` (0x3404): the row this paragraph terminates is a
    /// header row, repeated at the top of each page.
    pub f_table_header: bool,
    /// `sprmPItap` (0x6649): table nesting depth (1 = top-level table).
    pub itap: u8,
    /// Display name of the paragraph's style, resolved by `build_paragraphs`
    /// from its `istd` (so `extract_pap_props` alone leaves it `None`). `None`
    /// also for the default style (`istd` 0) and for a style with no name.
    pub style_name: Option<String>,
    /// `sprmPFInnerTableCell` (0x244B): the last paragraph of a cell of a
    /// nested table, which ends in a paragraph mark and not in `\x07`.
    pub f_inner_table_cell: bool,
    /// List level (0-based) when the paragraph is a list item. `None` when
    /// the paragraph carries no list SPRM. (Reserved for list support; the
    /// `table.doc` fixture has no in-body lists to verify this against.)
    pub ilvl: Option<u8>,
    /// List format override id (`ilfo`, `sprmPIlfo` `0x460B`), read as the
    /// *signed* `i16` [MS-DOC] specifies. `None` when the SPRM is absent
    /// (defaults to "not in a list"). Bands: `0`/`0xF801` = not in a list;
    /// `0x0001`–`0x07FE` = 1-based index; `0xF802`–`0xFFFF` = negated index
    /// (still a list item, resolved by `ListFormatting::level_for`).
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
    /// Outline level, in [MS-DOC]'s value space: `0` = Heading 1 … `8` =
    /// Heading 9, and `9` = body text. `OutlineLevel::BodyText` now carries the
    /// explicit "not a heading" case, so this field is `None` only when no
    /// level was resolved at all.
    ///
    /// This is the evidence that a document has *real* heading structure, which
    /// is what gates the line-shape heading guess: a paragraph that resolved a
    /// level from either `sprmPOutLvl` or its style sets it; one marked
    /// `BodyText` explicitly suppresses a heading; `None` means "ask the guess".
    ///
    /// It comes from either source, and is filled in two stages:
    /// `extract_pap_props` sets it from `sprmPOutLvl` (0x2640) when that SPRM
    /// is present; otherwise it stays `None` until `build_paragraphs` resolves
    /// the paragraph's style (`istd`) into a level. So the value returned by
    /// `extract_pap_props` alone is *not* necessarily final.
    pub outline_level: Option<OutlineLevel>,
    /// The style index this paragraph is restyled to by a direct `sprmPIstd`
    /// (0x4600) in the grpprl, if any. It overrides the PAPX header's own
    /// `istd` when resolving the style (see `crate::doc::styles`); `None` means
    /// the PAPX header `istd` applies unchanged.
    pub style_istd: Option<u16>,
    /// `sprmPJc` (0x2461) / legacy `sprmPJc80` (0x2403): paragraph
    /// alignment. Values beyond the four/five `ParagraphAlignment`
    /// variants (Arabic Kashida justification, Thai distribute, …) map to
    /// the closest fit (`Justify`/`Distribute`) rather than being dropped.
    pub alignment: Option<ParagraphAlignment>,
    /// `sprmPDxaLeft` (0x845E) / legacy `sprmPDxaLeft80` (0x840F), in
    /// twips.
    pub indent_left_twips: Option<i32>,
    /// `sprmPDxaRight` (0x845D) / legacy `sprmPDxaRight80` (0x840E), in
    /// twips.
    pub indent_right_twips: Option<i32>,
    /// `sprmPDxaLeft1` (0x8460) / legacy `sprmPDxaLeft180` (0x8411), in
    /// twips. Negative = hanging indent.
    pub first_line_indent_twips: Option<i32>,
    /// `sprmPDyaBefore` (0xA413), in twips.
    pub space_before_twips: Option<u32>,
    /// `sprmPDyaAfter` (0xA414), in twips.
    pub space_after_twips: Option<u32>,
    /// `sprmPDyaLine` (0x6412): line spacing, in the IR's DOCX-shaped
    /// units (240ths of a line for `Auto`, twips otherwise).
    pub line_spacing: Option<LineSpacing>,
}

/// What a paragraph's outline level is, once direct formatting and the style
/// have been consulted.
///
/// `sprmPOutLvl` (0x2640) states a level directly, and it also states *body
/// text* (operand `0x09`) — which is not the same as saying nothing at all:
/// a paragraph styled `Heading 3` but marked body text must stay body text
/// rather than fall through to its style. That difference is why this is one
/// value with three states and not an `Option<u8>` plus a bool: two fields
/// that must agree is a state you can get wrong.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutlineLevel {
    /// A real heading level, in [MS-DOC]'s zero-based space: `0` = Heading 1
    /// … `8` = Heading 9.
    ///
    /// `source` is carried in the value rather than in a parallel bool: whether
    /// the level came from direct formatting or was resolved from a style
    /// changes what the document-level gate may conclude, and a second field
    /// that has to agree with the first is a state someone will set halfway.
    Heading {
        /// The zero-based outline level.
        level: u8,
        /// Where the level came from.
        source: LevelSource,
    },
    /// `sprmPOutLvl` operand `0x09`: explicitly *not* a heading.
    BodyText,
}

/// Where a paragraph's outline level came from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LevelSource {
    /// Stated by the paragraph itself, via `sprmPOutLvl` (0x2640).
    Sprm,
    /// Resolved from the paragraph's style (`istd` → `sti` / `Heading N`).
    Style,
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

/// Declare the paragraph SPRM dispatch and its identity registry from one
/// source.
///
/// The recurring defect this closes is "the constant is right but names a
/// different property": `0x6412` is `sprmPDyaLine` (line spacing, 758
/// occurrences in a 246-file corpus) and was once read as an outline level.
/// Per-opcode inline assertions cannot catch a *new* wrong arm, because
/// nobody writes an assertion for an arm they believe is correct.
///
/// Generating both the `match` and [`PAP_SPRM_REGISTRY`] from the same
/// invocation makes the check fail closed: an arm added without naming the
/// property and its [MS-DOC] section does not compile, and one that names
/// the wrong property fails `test_registry_matches_the_spec_table`.
macro_rules! pap_sprm_dispatch {
    (
        $(
            $(#[$meta:meta])*
            $spec_name:literal @ $section:literal => [$($opcode:literal),+ $(,)?]
                ($props:ident, $operand:ident) $body:block
        )*
    ) => {
        /// Every paragraph SPRM this crate decodes: `(opcode, spec name,
        /// [MS-DOC] section)`. Derived from the dispatch below, never
        /// maintained alongside it.
        ///
        /// Consumed by the identity tests; the value of enumerating it is
        /// that the enumeration cannot drift from the dispatch.
        #[cfg(test)]
        pub const PAP_SPRM_REGISTRY: &[(u16, &str, &str)] = &[
            $($(($opcode, $spec_name, $section),)+)*
        ];

        fn dispatch_pap_sprm(props: &mut PapProps, opcode: u16, operand: &[u8]) {
            match opcode {
                $(
                    $($opcode)|+ => {
                        let $props = props;
                        let $operand = operand;
                        $body
                    },
                )*
                _ => {},
            }
        }
    };
}

pap_sprm_dispatch! {
    /// 1-byte operand: 0..=8 are Heading 1..9 and 9 is body text.
    /// The opcode is 0x2640 — *not* the 0x6412 a byte-swapped reading
    /// suggests, which is `sprmPDyaLine`.
    ///
    /// `OutlineLevel::BodyText` records that the SPRM is present *and* says
    /// body text, which is different from the SPRM being absent: only the
    /// former suppresses a heading the paragraph's style would otherwise give
    /// it. An operand above 9 is not an outline level at all, so it is ignored
    /// and `outline_level` is left as `None` — "ask the style".
    "sprmPOutLvl" @ "2.6.2" => [0x2640] (props, operand) {
        if let Some(&lvl) = operand.first() {
            if lvl < OUTLVL_BODY_TEXT {
                props.outline_level = Some(OutlineLevel::Heading {
                    level: lvl,
                    source: LevelSource::Sprm,
                });
            } else if lvl == OUTLVL_BODY_TEXT {
                // Explicitly body text: settle it, and never consult the style.
                props.outline_level = Some(OutlineLevel::BodyText);
            }
        }
    }

    /// 2-byte operand: the style index (`istd`) this paragraph is restyled to.
    /// It overrides the PAPX header `istd` when the style is resolved.
    ///
    /// Not 0x640A: `ispmd` 0x0A is `sprmPIlvl` (0x260A), and 0x640A is not a
    /// paragraph SPRM at all.
    "sprmPIstd" @ "2.6.2" => [0x4600] (props, operand) {
        if operand.len() >= 2 {
            props.style_istd = Some(u16::from_le_bytes([operand[0], operand[1]]));
        }
    }

    /// 1-byte operand, bit 0 = fInTable.
    "sprmPFInTable" @ "2.6.2" => [0x2416] (props, operand) {
        if let Some(&b) = operand.first() {
            props.f_in_table = (b & 1) != 0;
        }
    }

    /// 4-byte signed operand: the table nesting depth. An out-of-range value
    /// saturates to the top of a `u8` instead of being truncated (257 would
    /// become 1) or clamped to the bottom (a negative depth would become 0): it
    /// must not pass for an ordinary paragraph, and the top is where the table
    /// builder reports it.
    "sprmPItap" @ "2.6.2" => [0x6649] (props, operand) {
        if operand.len() >= 4 {
            let v = i32::from_le_bytes([operand[0], operand[1], operand[2], operand[3]]);
            props.itap = u8::try_from(v).unwrap_or(u8::MAX);
        }
    }

    /// 1-byte `Bool8` operand, bit 0: this paragraph is the last one of a cell
    /// of a **nested** table (`itap` >= 2). Word ends such a cell with an
    /// ordinary paragraph mark carrying this flag; only an outermost cell ends
    /// with the `\x07` cell mark.
    "sprmPFInnerTableCell" @ "2.6.2" => [0x244B] (props, operand) {
        props.f_inner_table_cell = operand.first().is_some_and(|&b| b & 1 != 0);
    }

    /// 1-byte `Bool8` operand, bit 0: this paragraph ends a row of a
    /// **nested** table. It is the nested counterpart of `sprmPFTtp`, carried
    /// by an ordinary paragraph mark together with the row definition.
    "sprmPFInnerTtp" @ "2.6.2" => [0x244C] (props, operand) {
        if operand.first().is_some_and(|&b| b & 1 != 0) {
            props.is_table_trailing_mark = true;
        }
    }

    /// 1-byte `Bool8` operand: the row is a header row (`fTableHeader`).
    /// Carried by the row-terminator paragraph, like the rest of the TAP.
    "sprmTTableHeader" @ "2.6.3" => [0x3404] (props, operand) {
        props.f_table_header = operand.first().is_some_and(|&b| b & 1 != 0);
    }

    /// Presence marks a row-terminator paragraph; the operand carries the
    /// row definition (cells, boundaries).
    "sprmTDefTable" @ "2.6.3" => [0xD608] (props, operand) {
        props.is_table_trailing_mark = true;
        props.tap = parse_tdef_table(operand);
    }

    /// 2-byte operand read as a *signed* `i16`: `0x0000`/`0xF801` mean "not
    /// in a list", `0x0001`–`0x07FE` are 1-based indices into
    /// `PlfLfo.rgLfo`, and `0xF802`–`0xFFFF` are the negation of a 1-based
    /// index (still in a list). Storing it signed keeps the negation
    /// explicit.
    "sprmPIlfo" @ "2.6.2" => [0x460B] (props, operand) {
        if operand.len() >= 2 {
            props.ilfo = Some(i16::from_le_bytes([operand[0], operand[1]]));
        }
    }

    /// 1-byte operand: the list level (0-based).
    "sprmPIlvl" @ "2.6.2" => [0x260A] (props, operand) {
        if let Some(&b) = operand.first() {
            props.ilvl = Some(b);
        }
    }

    /// Tab stops. Decoding needs the opcode itself — `sprmPChgTabsPapx`
    /// (0xC60D) frames the same payload differently — so
    /// `extract_pap_props` handles both directly. The entries are declared
    /// here so the registry still enumerates every opcode the crate claims.
    "sprmPChgTabs" @ "2.6.2" => [0xC615] (props, operand) {
        let _ = (props, operand);
    }

    /// See `sprmPChgTabs`.
    "sprmPChgTabsPapx" @ "2.6.2" => [0xC60D] (props, operand) {
        let _ = (props, operand);
    }

    /// Unsigned 8-bit enum (1 byte, spra 0/1): paragraph justification.
    /// `sprmPJc` (modern, values 0-9) and `sprmPJc80` (legacy, values
    /// 0-5) share the same 0-3 meanings (left/center/right/justify);
    /// `sprmPJc`'s extended values beyond 3 are folded into the closest
    /// fit rather than dropped.
    "sprmPJc" @ "2.6.2" => [0x2461] (props, operand) {
        if let Some(&jc) = operand.first() {
            props.alignment = match jc {
                0 => Some(ParagraphAlignment::Left),
                1 => Some(ParagraphAlignment::Center),
                2 => Some(ParagraphAlignment::Right),
                4 => Some(ParagraphAlignment::Distribute),
                _ => Some(ParagraphAlignment::Justify),
            };
        }
    }

    /// See `sprmPJc`.
    "sprmPJc80" @ "2.6.2" => [0x2403] (props, operand) {
        if let Some(&jc) = operand.first() {
            props.alignment = match jc {
                0 => Some(ParagraphAlignment::Left),
                1 => Some(ParagraphAlignment::Center),
                2 => Some(ParagraphAlignment::Right),
                _ => Some(ParagraphAlignment::Justify),
            };
        }
    }

    /// XAS (signed 16-bit, twips, spra 4): logical left indent.
    "sprmPDxaLeft" @ "2.6.2" => [0x845E] (props, operand) {
        if operand.len() >= 2 {
            props.indent_left_twips = Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// See `sprmPDxaLeft` (legacy form, same unit and meaning).
    "sprmPDxaLeft80" @ "2.6.2" => [0x840F] (props, operand) {
        if operand.len() >= 2 {
            props.indent_left_twips = Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// XAS (signed 16-bit, twips, spra 4): logical right indent.
    "sprmPDxaRight" @ "2.6.2" => [0x845D] (props, operand) {
        if operand.len() >= 2 {
            props.indent_right_twips = Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// See `sprmPDxaRight` (legacy form, same unit and meaning).
    "sprmPDxaRight80" @ "2.6.2" => [0x840E] (props, operand) {
        if operand.len() >= 2 {
            props.indent_right_twips = Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// XAS (signed 16-bit, twips, spra 4): first-line indent relative to
    /// the rest of the paragraph. Negative = hanging indent.
    "sprmPDxaLeft1" @ "2.6.2" => [0x8460] (props, operand) {
        if operand.len() >= 2 {
            props.first_line_indent_twips =
                Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// See `sprmPDxaLeft1` (legacy form, same unit and meaning).
    "sprmPDxaLeft180" @ "2.6.2" => [0x8411] (props, operand) {
        if operand.len() >= 2 {
            props.first_line_indent_twips =
                Some(i32::from(i16::from_le_bytes([operand[0], operand[1]])));
        }
    }

    /// Unsigned 16-bit integer (spra 5), twips: spacing before the
    /// paragraph. Per [MS-DOC] the value MUST be 0x0000-0x7BC0; an
    /// out-of-range operand is ignored.
    "sprmPDyaBefore" @ "2.6.2" => [0xA413] (props, operand) {
        if operand.len() >= 2 {
            let v = u16::from_le_bytes([operand[0], operand[1]]);
            if v <= 0x7BC0 {
                props.space_before_twips = Some(u32::from(v));
            }
        }
    }

    /// See `sprmPDyaBefore`.
    "sprmPDyaAfter" @ "2.6.2" => [0xA414] (props, operand) {
        if operand.len() >= 2 {
            let v = u16::from_le_bytes([operand[0], operand[1]]);
            if v <= 0x7BC0 {
                props.space_after_twips = Some(u32::from(v));
            }
        }
    }

    /// `LSPD` (4 bytes, spra 3; [MS-DOC] §2.9.150): `dyaLine` (signed
    /// 16-bit) then `fMultLinespace` (16-bit). With `fMultLinespace` = 1,
    /// `dyaLine` is in 240ths of a line (single = 240) — DOCX's `auto`
    /// rule. With 0, a non-negative `dyaLine` is an at-least height in
    /// twips and a negative one an exact height of `|dyaLine|` twips.
    "sprmPDyaLine" @ "2.6.2" => [0x6412] (props, operand) {
        if operand.len() >= 4 {
            let dya = i16::from_le_bytes([operand[0], operand[1]]);
            let mult = u16::from_le_bytes([operand[2], operand[3]]);
            props.line_spacing = Some(if mult == 1 {
                LineSpacing::Auto(dya.max(0) as u32)
            } else if dya < 0 {
                LineSpacing::Exact(u32::from(dya.unsigned_abs()))
            } else {
                LineSpacing::AtLeast(dya as u32)
            });
        }
    }
}

/// Decode a PAP `grpprl` into the paragraph flags we care about.
///
/// Unknown SPRMs are ignored. An empty `grpprl` yields the default
/// (all-false) `PapProps`, which classifies the paragraph as ordinary prose.
pub fn extract_pap_props(grpprl: &[u8]) -> PapProps {
    let mut props = PapProps::default();

    for sprm in parse_grpprl(grpprl) {
        // `sprmPChgTabs` needs the opcode itself to know how its operand is
        // framed, so it is handled here rather than in the dispatch table.
        if matches!(sprm.opcode, 0xC615 | 0xC60D) {
            if !sprm.operand.is_empty() {
                props.tabs = decode_pchg_tabs(sprm.opcode, &sprm.operand);
            }
            continue;
        }
        dispatch_pap_sprm(&mut props, sprm.opcode, &sprm.operand);
    }

    props
}

/// Decode a ToggleOperand ([MS-DOC] §2.9.339) against the value in force:
/// `0x00` off, `0x01` on, `0x80` the value in force (the style's), `0x81`
/// its negation — POI's `getFlag`. Reading every non-zero operand as *on*
/// turned `0x80` ("same as the style") runs bold, italic or hidden.
fn toggle(operand: u8, current: bool) -> bool {
    match operand {
        0x00 => false,
        0x80 => current,
        0x81 => !current,
        _ => true,
    }
}

/// Character-property flags distilled from a CHP grpprl (`sgc` == 2).
///
/// Revision-mark flags and
/// bold/italic/underline/color/font-size. The font is kept as an index
/// only (`sprmCRgFtc2`, see [`ChpProps::ftc_other`]); resolving it to a name
/// is `super::fonts`' job, because it needs the `SttbfFfn` font table.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChpProps {
    /// `sprmCFRMarkDel` (0x0800): the run is formatted as deleted
    /// revision-mark text. The "accepted" view (what `plain_text()` and
    /// the IR both show) excludes it, mirroring the policy already
    /// applied to DOCX's `w:del`.
    pub f_rmark_del: bool,
    /// `sprmCFRMarkIns` (0x0801): the run is formatted as inserted
    /// revision-mark text. Inserted text is part of the accepted view
    /// (kept, not filtered); tracked here only so the registry enumerates
    /// every CHP opcode this crate recognizes, not just the ones that
    /// currently change extraction behavior.
    pub f_rmark_ins: bool,
    /// `sprmCFBold` (0x0835).
    pub bold: bool,
    /// `sprmCFItalic` (0x0836).
    pub italic: bool,
    /// `sprmCKul` (0x2A3E): underline style. `None` (Kul value `0`, no
    /// underline) is the default, not `Some(UnderlineStyle::None)` —
    /// there's no source signal that distinguishes "never set" from an
    /// explicit "turned off", so the simpler of the two is used.
    pub underline: Option<UnderlineStyle>,
    /// `sprmCCv` (0x6870): true RGB text color. `None` when the color is
    /// `cvAuto` (`fAuto` byte set) or the SPRM is absent.
    pub color: Option<[u8; 3]>,
    /// `sprmCHps` (0x4A43): font size in half-points.
    pub font_size_half_pt: Option<u32>,
    /// `sprmCRgFtc2` (0x4A51): the font-table index of the font used for text
    /// outside U+0020..U+007F, which is where the private-use characters of a
    /// symbol font are. `None` when the run does not set it.
    pub ftc_other: Option<u16>,
    /// `sprmCFStrike` (0x0837) / `sprmCFDStrike` (0x2A53): single or
    /// double strikethrough.
    pub strike: bool,
    /// `sprmCFVanish` (0x083C): hidden text. Word does not display it, so
    /// it is excluded from extracted text the way the DOCX reader excludes
    /// `w:vanish` runs.
    pub vanish: bool,
    /// `sprmCFCaps` (0x083B): displayed in capitals.
    pub all_caps: bool,
    /// `sprmCFSmallCaps` (0x083A): displayed in small capitals.
    pub small_caps: bool,
    /// `sprmCIss` (0x2A48): superscript/subscript.
    pub iss: Option<VerticalAlign>,
    /// `sprmCHpsPos` (0x4845): baseline offset in half-points, positive
    /// raised, negative lowered.
    pub hps_pos: Option<i16>,
    /// `sprmCHighlight` (0x2A0C): highlight colour, resolved from its `Ico`
    /// index. `None` for `0` (no highlight).
    pub highlight: Option<[u8; 3]>,
}

impl ChpProps {
    /// The run's vertical alignment: `sprmCIss` when present; otherwise a
    /// non-zero `sprmCHpsPos` raise/lower, the closest thing the IR has to
    /// a baseline offset.
    pub fn vertical_align(&self) -> Option<VerticalAlign> {
        if let Some(v) = &self.iss {
            return Some(v.clone());
        }
        match self.hps_pos {
            Some(p) if p > 0 => Some(VerticalAlign::Superscript),
            Some(p) if p < 0 => Some(VerticalAlign::Subscript),
            _ => None,
        }
    }

    /// Text Word does not show in the accepted view: deleted revision
    /// text or hidden (`sprmCFVanish`) text.
    pub fn is_excluded(&self) -> bool {
        (self.f_rmark_del && !crate::read_options::include_deleted_text())
            || (self.vanish && !crate::read_options::include_hidden_text())
    }
}

/// The `Ico` colour table ([MS-DOC] §2.9.119), used by `sprmCIco` and
/// `sprmCHighlight`. `0` is "auto" (no colour); indices past 16 are
/// undefined.
fn ico_to_rgb(ico: u8) -> Option<[u8; 3]> {
    Some(match ico {
        0x01 => [0x00, 0x00, 0x00], // black
        0x02 => [0x00, 0x00, 0xFF], // blue
        0x03 => [0x00, 0xFF, 0xFF], // cyan
        0x04 => [0x00, 0xFF, 0x00], // green
        0x05 => [0xFF, 0x00, 0xFF], // magenta
        0x06 => [0xFF, 0x00, 0x00], // red
        0x07 => [0xFF, 0xFF, 0x00], // yellow
        0x08 => [0xFF, 0xFF, 0xFF], // white
        0x09 => [0x00, 0x00, 0x80], // dark blue
        0x0A => [0x00, 0x80, 0x80], // dark cyan
        0x0B => [0x00, 0x80, 0x00], // dark green
        0x0C => [0x80, 0x00, 0x80], // dark magenta
        0x0D => [0x80, 0x00, 0x00], // dark red
        0x0E => [0x80, 0x80, 0x00], // dark yellow
        0x0F => [0x80, 0x80, 0x80], // dark gray
        0x10 => [0xC0, 0xC0, 0xC0], // light gray
        _ => return None,
    })
}

/// Declare the character SPRM dispatch and its identity registry from one
/// source, mirroring [`pap_sprm_dispatch`] for `sgc == 2` (CHP) opcodes.
macro_rules! chp_sprm_dispatch {
    (
        $(
            $(#[$meta:meta])*
            $spec_name:literal @ $section:literal => [$($opcode:literal),+ $(,)?]
                ($props:ident, $operand:ident) $body:block
        )*
    ) => {
        /// Every character SPRM this crate decodes: `(opcode, spec name,
        /// [MS-DOC] section)`. Derived from the dispatch below, never
        /// maintained alongside it.
        #[cfg(test)]
        pub const CHP_SPRM_REGISTRY: &[(u16, &str, &str)] = &[
            $($(($opcode, $spec_name, $section),)+)*
        ];

        fn dispatch_chp_sprm(props: &mut ChpProps, opcode: u16, operand: &[u8]) {
            match opcode {
                $(
                    $($opcode)|+ => {
                        let $props = props;
                        let $operand = operand;
                        $body
                    },
                )*
                _ => {},
            }
        }
    };
}

chp_sprm_dispatch! {
    /// ToggleOperand (1 byte, spra 0): non-zero means deleted revision-mark
    /// text.
    "sprmCFRMarkDel" @ "2.6.1" => [0x0800] (props, operand) {
        if let Some(&b) = operand.first() {
            props.f_rmark_del = toggle(b, props.f_rmark_del);
        }
    }

    /// ToggleOperand (1 byte, spra 0): non-zero means inserted
    /// revision-mark text.
    "sprmCFRMarkIns" @ "2.6.1" => [0x0801] (props, operand) {
        if let Some(&b) = operand.first() {
            props.f_rmark_ins = toggle(b, props.f_rmark_ins);
        }
    }

    /// ToggleOperand (1 byte, spra 0): whether the text is bold.
    "sprmCFBold" @ "2.6.1" => [0x0835] (props, operand) {
        if let Some(&b) = operand.first() {
            props.bold = toggle(b, props.bold);
        }
    }

    /// ToggleOperand (1 byte, spra 0): whether the text is italicized.
    "sprmCFItalic" @ "2.6.1" => [0x0836] (props, operand) {
        if let Some(&b) = operand.first() {
            props.italic = toggle(b, props.italic);
        }
    }

    /// Kul value (1 byte, spra 1): underlining style. `0` = none (leaves
    /// `underline` at its default `None`); other values map to the
    /// closest `UnderlineStyle` variant, with any value this crate does
    /// not otherwise recognize (e.g. `Hidden`, or a reserved value)
    /// falling back to `Single` rather than being silently dropped —
    /// some underline is a closer approximation of the source than none.
    "sprmCKul" @ "2.6.1" => [0x2A3E] (props, operand) {
        if let Some(&kul) = operand.first() {
            props.underline = match kul {
                0x00 => None,
                0x03 => Some(UnderlineStyle::Double),
                0x04 => Some(UnderlineStyle::Dotted),
                0x06 => Some(UnderlineStyle::Thick),
                0x07 => Some(UnderlineStyle::Dash),
                0x09 => Some(UnderlineStyle::DotDash),
                0x0A => Some(UnderlineStyle::DotDotDash),
                0x0B => Some(UnderlineStyle::Wave),
                0x02 => Some(UnderlineStyle::Words),
                _ => Some(UnderlineStyle::Single),
            };
        }
    }

    /// COLORREF (4 bytes, spra 3): `[red, green, blue, fAuto]`. `fAuto !=
    /// 0` means "use the automatic/default color" — treated as "no
    /// override", not black.
    "sprmCCv" @ "2.6.1" => [0x6870] (props, operand) {
        if operand.len() >= 4 {
            props.color =
                if operand[3] != 0 { None } else { Some([operand[0], operand[1], operand[2]]) };
        }
    }

    /// Signed 2-byte integer (spra 2): an index into the font table
    /// (`SttbfFfn`) of the font used for a character outside U+0020..U+007F
    /// ([MS-DOC] §2.6.1). A negative operand cannot be an index and is ignored.
    "sprmCRgFtc2" @ "2.6.1" => [0x4A51] (props, operand) {
        if operand.len() >= 2 {
            let ftc = i16::from_le_bytes([operand[0], operand[1]]);
            props.ftc_other = u16::try_from(ftc).ok();
        }
    }

    /// Unsigned 2-byte integer (spra 2), half-points. Per [MS-DOC], the
    /// value MUST be between 2 and 3276; an out-of-range operand is
    /// ignored rather than stored, since it cannot be a real font size.
    "sprmCHps" @ "2.6.1" => [0x4A43] (props, operand) {
        if operand.len() >= 2 {
            let hps = u16::from_le_bytes([operand[0], operand[1]]);
            if (2..=3276).contains(&hps) {
                props.font_size_half_pt = Some(u32::from(hps));
            }
        }
    }

    /// ToggleOperand (1 byte, spra 0): single strikethrough.
    "sprmCFStrike" @ "2.6.1" => [0x0837] (props, operand) {
        if let Some(&b) = operand.first() {
            props.strike = toggle(b, props.strike);
        }
    }

    /// ToggleOperand (1 byte, spra 1): double strikethrough — surfaced as
    /// the IR's single strikethrough flag, as the DOCX reader does for
    /// `w:dstrike`.
    "sprmCFDStrike" @ "2.6.1" => [0x2A53] (props, operand) {
        if let Some(&b) = operand.first() {
            props.strike = toggle(b, props.strike);
        }
    }

    /// ToggleOperand (1 byte, spra 0): hidden text.
    "sprmCFVanish" @ "2.6.1" => [0x083C] (props, operand) {
        if let Some(&b) = operand.first() {
            // Hiding deletes content, so it takes certainty: `0x80`/`0x81`
            // are relative to the style's value, which is not resolved
            // here, and leave the value in force (see the test).
            props.vanish = match b {
                0x00 => false,
                0x80 | 0x81 => props.vanish,
                _ => true,
            };
        }
    }

    /// ToggleOperand (1 byte, spra 0): small capitals.
    "sprmCFSmallCaps" @ "2.6.1" => [0x083A] (props, operand) {
        if let Some(&b) = operand.first() {
            props.small_caps = toggle(b, props.small_caps);
        }
    }

    /// ToggleOperand (1 byte, spra 0): all capitals.
    "sprmCFCaps" @ "2.6.1" => [0x083B] (props, operand) {
        if let Some(&b) = operand.first() {
            props.all_caps = toggle(b, props.all_caps);
        }
    }

    /// 1-byte operand (spra 1): `0` normal, `1` superscript, `2`
    /// subscript.
    "sprmCIss" @ "2.6.1" => [0x2A48] (props, operand) {
        if let Some(&b) = operand.first() {
            props.iss = match b {
                1 => Some(VerticalAlign::Superscript),
                2 => Some(VerticalAlign::Subscript),
                _ => None,
            };
        }
    }

    /// Signed 2-byte integer (spra 2): baseline offset in half-points.
    "sprmCHpsPos" @ "2.6.1" => [0x4845] (props, operand) {
        if operand.len() >= 2 {
            let v = i16::from_le_bytes([operand[0], operand[1]]);
            props.hps_pos = (v != 0).then_some(v);
        }
    }

    /// `Ico` (1 byte, spra 1): highlight colour; `0` is no highlight.
    "sprmCHighlight" @ "2.6.1" => [0x2A0C] (props, operand) {
        if let Some(&b) = operand.first() {
            props.highlight = ico_to_rgb(b);
        }
    }

    /// `Ico` (1 byte, spra 1): text colour from the 16-colour table. `0`
    /// (auto) clears it. Word writes `sprmCCv` after this when it has a
    /// true colour, so the later SPRM wins in grpprl order.
    "sprmCIco" @ "2.6.1" => [0x2A42] (props, operand) {
        if let Some(&b) = operand.first() {
            props.color = ico_to_rgb(b);
        }
    }
}

/// `Prm0.isprm` → SPRM opcode, for the character properties this crate
/// decodes. Transcribed from the `Prm0` table in [MS-DOC] §2.9.177 (the
/// `isprm` values are Word 6's SPRM numbers); every `Prm0` SPRM takes a
/// 1-byte operand. Paragraph and other modifiers are not listed and are
/// not applied to characters.
const PRM0_CHP_SPRMS: &[(u8, u16)] = &[
    (0x41, 0x0800), // sprmCFRMarkDel
    (0x42, 0x0801), // sprmCFRMarkIns
    (0x4D, 0x2A0C), // sprmCHighlight
    (0x55, 0x0835), // sprmCFBold
    (0x56, 0x0836), // sprmCFItalic
    (0x57, 0x0837), // sprmCFStrike
    (0x5A, 0x083A), // sprmCFSmallCaps
    (0x5B, 0x083B), // sprmCFCaps
    (0x5C, 0x083C), // sprmCFVanish
    (0x5E, 0x2A3E), // sprmCKul
    (0x62, 0x2A42), // sprmCIco
    (0x68, 0x2A48), // sprmCIss
    (0x73, 0x2A53), // sprmCFDStrike
];

/// The character SPRM a `Prm0.isprm` names, if it is one this crate
/// decodes.
pub(crate) fn prm0_sprm(isprm: u8) -> Option<u16> {
    PRM0_CHP_SPRMS
        .iter()
        .find(|(i, _)| *i == isprm)
        .map(|&(_, op)| op)
}

/// Apply a CHP `grpprl` on top of already-resolved properties — a piece
/// modifier over the run's own CHPX.
pub fn apply_chp_grpprl(props: &mut ChpProps, grpprl: &[u8]) {
    for sprm in parse_grpprl(grpprl) {
        dispatch_chp_sprm(props, sprm.opcode, &sprm.operand);
    }
}

/// Decode a CHP `grpprl` into the character flags we care about.
///
/// Unknown SPRMs are ignored. An empty `grpprl` yields the default
/// (all-false) `ChpProps`, which classifies the run as ordinary,
/// non-revision-marked text.
pub fn extract_chp_props(grpprl: &[u8]) -> ChpProps {
    let mut props = ChpProps::default();
    for sprm in parse_grpprl(grpprl) {
        dispatch_chp_sprm(&mut props, sprm.opcode, &sprm.operand);
    }
    props
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Opcode → property name, transcribed from [MS-DOC] §2.6.2 (paragraph
    /// SPRMs) and §2.6.3 (table SPRMs).
    ///
    /// This exists to be *disagreed with*. The dispatch registry is
    /// generated from the match arms, so an arm whose constant names a
    /// different property than its comment claims shows up here as a
    /// mismatch — which per-opcode inline assertions cannot do, because
    /// nobody writes an assertion for an arm they believe is correct.
    ///
    /// Entries beyond what the crate decodes are deliberate: they are the
    /// opcodes that have been mistaken for the ones above.
    const MS_DOC_SPRM_TABLE: &[(u16, &str)] = &[
        (0x2416, "sprmPFInTable"),
        (0x2417, "sprmPFTtp"),
        (0x260A, "sprmPIlvl"),
        (0x2640, "sprmPOutLvl"),
        (0x4600, "sprmPIstd"),
        (0x460B, "sprmPIlfo"),
        (0x6412, "sprmPDyaLine"),
        (0x6649, "sprmPItap"),
        (0x244B, "sprmPFInnerTableCell"),
        (0x244C, "sprmPFInnerTtp"),
        (0xC60D, "sprmPChgTabsPapx"),
        (0xC615, "sprmPChgTabs"),
        (0xD608, "sprmTDefTable"),
        (0x3404, "sprmTTableHeader"),
        (0x6646, "sprmPHugePapx"),
        (0x2461, "sprmPJc"),
        (0x2403, "sprmPJc80"),
        (0x845E, "sprmPDxaLeft"),
        (0x840F, "sprmPDxaLeft80"),
        (0x845D, "sprmPDxaRight"),
        (0x840E, "sprmPDxaRight80"),
        (0x8460, "sprmPDxaLeft1"),
        (0x8411, "sprmPDxaLeft180"),
        (0xA413, "sprmPDyaBefore"),
        (0xA414, "sprmPDyaAfter"),
    ];

    /// Every opcode the dispatch claims must name the property [MS-DOC]
    /// gives it, and must cite a section.
    #[test]
    fn test_registry_matches_the_spec_table() {
        for &(opcode, name, section) in PAP_SPRM_REGISTRY {
            let expected = MS_DOC_SPRM_TABLE
                .iter()
                .find(|(o, _)| *o == opcode)
                .map(|(_, n)| *n);
            assert_eq!(
                expected,
                Some(name),
                "opcode 0x{opcode:04X} is decoded as {name}, but [MS-DOC] calls it {expected:?}"
            );
            assert!(
                section.starts_with("2.6"),
                "0x{opcode:04X} ({name}) must cite its [MS-DOC] §2.6.x section, got {section:?}"
            );
        }
    }

    /// No opcode may be dispatched twice — two arms claiming the same
    /// constant means one of them never runs.
    #[test]
    fn test_registry_has_no_duplicate_opcodes() {
        let mut seen: Vec<u16> = PAP_SPRM_REGISTRY.iter().map(|(o, _, _)| *o).collect();
        seen.sort_unstable();
        let before = seen.len();
        seen.dedup();
        assert_eq!(before, seen.len(), "duplicate opcode in the dispatch");
    }

    /// The opcodes that have actually been confused for the ones we decode
    /// must not be decoded as something else. `0x6412` is line spacing and
    /// occurs 758 times in a 246-file corpus; reading it as an outline
    /// level marked most of those paragraphs as headings. It is decoded now
    /// — as line spacing, and as nothing else.
    #[test]
    fn test_known_confusable_opcodes_are_not_claimed() {
        let claims: Vec<&str> = PAP_SPRM_REGISTRY
            .iter()
            .filter(|(o, _, _)| *o == 0x6412)
            .map(|(_, n, _)| *n)
            .collect();
        assert_eq!(claims, ["sprmPDyaLine"], "0x6412 is sprmPDyaLine only");
        assert!(
            !PAP_SPRM_REGISTRY.iter().any(|(o, _, _)| *o == 0x640A),
            "0x640A is not a paragraph property this crate decodes"
        );
        // And line spacing never leaks into the outline level.
        let props = extract_pap_props(&[0x12, 0x64, 0x01, 0x00, 0x00, 0x00]);
        assert_eq!(props.outline_level, None);
    }

    /// `sprmPDyaLine`'s three `LSPD` forms map onto the IR's line-spacing
    /// rules.
    #[test]
    fn test_sprm_pdya_line_decodes_lspd() {
        // 1.5 lines: dyaLine = 360, fMultLinespace = 1.
        let p = extract_pap_props(&[0x12, 0x64, 0x68, 0x01, 0x01, 0x00]);
        assert_eq!(p.line_spacing, Some(LineSpacing::Auto(360)));
        // At least 12pt: dyaLine = 240 twips, fMultLinespace = 0.
        let p = extract_pap_props(&[0x12, 0x64, 0xF0, 0x00, 0x00, 0x00]);
        assert_eq!(p.line_spacing, Some(LineSpacing::AtLeast(240)));
        // Exactly 18pt: dyaLine = -360, fMultLinespace = 0.
        let p = extract_pap_props(&[0x12, 0x64, 0x98, 0xFE, 0x00, 0x00]);
        assert_eq!(p.line_spacing, Some(LineSpacing::Exact(360)));
        // Truncated operand: ignored.
        let p = extract_pap_props(&[0x12, 0x64, 0x98]);
        assert_eq!(p.line_spacing, None);
    }

    /// The registry must actually be reachable from the decoder, so a
    /// registry that drifts away from the dispatch cannot pass silently.
    #[test]
    fn test_every_registered_opcode_changes_the_decoded_props() {
        // A one-byte operand is enough for the flag/level SPRMs; the
        // multi-byte ones get four bytes.
        for &(opcode, name, _) in PAP_SPRM_REGISTRY {
            if matches!(opcode, 0xC615 | 0xC60D | 0xD608) {
                continue; // variable-length payloads, covered by their own tests
            }
            let mut grpprl = opcode.to_le_bytes().to_vec();
            grpprl.extend_from_slice(&[1u8, 0, 0, 0]);
            let props = extract_pap_props(&grpprl);
            assert_ne!(
                format!("{props:?}"),
                format!("{:?}", PapProps::default()),
                "0x{opcode:04X} ({name}) is registered but decoding it changes nothing"
            );
        }
    }

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
    fn test_walks_fixed_and_variable_sprms() {
        let sprms = parse_grpprl(&cell_grpprl());
        assert_eq!(sprms.len(), 2);
        assert_eq!(sprms[0].opcode, 0x2416);
        assert_eq!(sprms[0].operand, vec![0x01]);
        assert_eq!(sprms[1].opcode, 0x6649);
        assert_eq!(sprms[1].operand, vec![0x01, 0x00, 0x00, 0x00]);
    }

    /// Regression: 0x6412 is **sprmPDyaLine** (line spacing), not an outline
    /// level. An earlier revision read its low byte as one, so ordinary line
    /// spacing turned body paragraphs into headings. 0x6412 occurs 758 times in
    /// the POI corpus, all of it line spacing, so this must stay inert.
    #[test]
    fn test_sprm_p_dya_line_is_not_read_as_an_outline_level() {
        // LSPD: dyaLine = 0x0005, fMultLinespace = 0x0000.
        let props = extract_pap_props(&[0x12, 0x64, 0x05, 0x00, 0x00, 0x00]);
        assert_eq!(props.outline_level, None, "line spacing must never become an outline level");
    }

    /// The `sprmPOutLvl` operand is **zero-based**: `0x00`–`0x08` are Heading
    /// 1–9 and `0x09` is body text. Each edge is pinned separately, because an
    /// off-by-one here is invisible in aggregate counts — it silently re-labels
    /// every heading in every document by one level.
    #[test]
    fn test_sprm_p_out_lvl_zero_is_heading_one() {
        let props = extract_pap_props(&[0x40, 0x26, 0x00]);
        assert_eq!(
            props.outline_level,
            Some(OutlineLevel::Heading {
                level: 0,
                source: LevelSource::Sprm
            }),
            "0x00 is Heading 1, not body text"
        );
    }

    #[test]
    fn test_sprm_p_out_lvl_eight_is_heading_nine() {
        let props = extract_pap_props(&[0x40, 0x26, 0x08]);
        assert_eq!(
            props.outline_level,
            Some(OutlineLevel::Heading {
                level: 8,
                source: LevelSource::Sprm
            }),
            "0x08 is Heading 9"
        );
    }

    #[test]
    fn test_sprm_p_out_lvl_nine_is_body_text() {
        let props = extract_pap_props(&[0x40, 0x26, 0x09]);
        assert_eq!(
            props.outline_level,
            Some(OutlineLevel::BodyText),
            "0x09 means body text, not Heading 9, and must still settle the question"
        );
    }

    #[test]
    fn test_sprm_p_out_lvl_above_nine_is_not_an_outline_level() {
        let props = extract_pap_props(&[0x40, 0x26, 0x0A]);
        assert_eq!(
            props.outline_level, None,
            "an operand above 0x09 is not an outline level and must not settle anything"
        );
    }

    /// `sprmPIstd` (0x4600) overrides the PAPX `istd`; its operand is the 2-byte
    /// istd. It surfaces on `PapProps.style_istd` from the same grpprl walk that
    /// decodes the other paragraph properties.
    #[test]
    fn test_pap_props_read_sprm_p_istd_override() {
        // opcode 0x4600 (LE) + 2-byte operand = istd 5.
        let grpprl = vec![0x00, 0x46, 0x05, 0x00];
        assert_eq!(extract_pap_props(&grpprl).style_istd, Some(5));
    }

    #[test]
    fn test_pap_props_style_istd_absent_when_no_sprm() {
        assert_eq!(extract_pap_props(&[0x16, 0x24, 0x01]).style_istd, None);
    }

    #[test]
    fn test_row_mark_walk_consumes_every_byte() {
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
    fn test_d608_two_byte_cb_decodes_wide_tables() {
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
    fn test_extracts_cell_props() {
        let props = extract_pap_props(&cell_grpprl());
        assert!(props.f_in_table);
        assert!(!props.is_table_trailing_mark);
        assert_eq!(props.itap, 1);
    }

    #[test]
    fn test_extracts_row_mark_props() {
        let props = extract_pap_props(&row_mark_grpprl());
        assert!(props.f_in_table);
        assert!(props.is_table_trailing_mark);
        assert_eq!(props.itap, 1);
        assert!(props.tap.is_some(), "0xD608 operand must parse into a TapInfo");
    }

    /// Regression: alignment/indentation/spacing SPRMs decode
    /// to their real values, not the always-`None`/`0` they were before.
    #[test]
    fn test_extract_pap_props_alignment_indent_and_spacing() {
        // sprmPJc(0x2461) = 1 (center)
        let props = extract_pap_props(&[0x61, 0x24, 0x01]);
        assert_eq!(props.alignment, Some(ParagraphAlignment::Center));

        // sprmPJc80(0x2403) = 2 (right) — legacy opcode, same meaning.
        let props = extract_pap_props(&[0x03, 0x24, 0x02]);
        assert_eq!(props.alignment, Some(ParagraphAlignment::Right));

        // sprmPJc(0x2461) = 3 (both/justify)
        let props = extract_pap_props(&[0x61, 0x24, 0x03]);
        assert_eq!(props.alignment, Some(ParagraphAlignment::Justify));

        // sprmPDxaLeft(0x845E) = 720 twips (0.5in), sprmPDxaLeft1(0x8460) =
        // -360 (hanging indent), sprmPDxaRight(0x845D) = 0.
        let mut grpprl = Vec::new();
        grpprl.extend_from_slice(&[0x5E, 0x84]);
        grpprl.extend_from_slice(&720i16.to_le_bytes());
        grpprl.extend_from_slice(&[0x60, 0x84]);
        grpprl.extend_from_slice(&(-360i16).to_le_bytes());
        grpprl.extend_from_slice(&[0x5D, 0x84]);
        grpprl.extend_from_slice(&0i16.to_le_bytes());
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.indent_left_twips, Some(720));
        assert_eq!(props.first_line_indent_twips, Some(-360));
        assert_eq!(props.indent_right_twips, Some(0));

        // sprmPDyaBefore(0xA413) = 200, sprmPDyaAfter(0xA414) = 100 twips.
        let mut grpprl = Vec::new();
        grpprl.extend_from_slice(&[0x13, 0xA4]);
        grpprl.extend_from_slice(&200u16.to_le_bytes());
        grpprl.extend_from_slice(&[0x14, 0xA4]);
        grpprl.extend_from_slice(&100u16.to_le_bytes());
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.space_before_twips, Some(200));
        assert_eq!(props.space_after_twips, Some(100));

        // Out-of-range spacing (> 0x7BC0) is ignored, not stored.
        let mut grpprl = vec![0x13, 0xA4];
        grpprl.extend_from_slice(&0x7BC1u16.to_le_bytes());
        assert_eq!(extract_pap_props(&grpprl).space_before_twips, None);
    }

    #[test]
    fn test_a_negative_table_depth_stays_out_of_range_instead_of_becoming_outermost() {
        // 0x6649 is a signed 4-byte operand. A negative depth is as invalid as
        // one past 255: both saturate to the top, where the table builder
        // reports them, and neither passes for an ordinary outermost paragraph.
        for v in [-1i32, i32::MIN, 256, i32::MAX] {
            let mut g = vec![0x49, 0x66];
            g.extend_from_slice(&v.to_le_bytes());
            assert_eq!(extract_pap_props(&g).itap, u8::MAX, "operand {v}");
        }
        let mut g = vec![0x49, 0x66];
        g.extend_from_slice(&2i32.to_le_bytes());
        assert_eq!(extract_pap_props(&g).itap, 2);
    }

    #[test]
    fn test_empty_grpprl_is_ordinary_prose() {
        let props = extract_pap_props(&[]);
        assert!(!props.f_in_table);
        assert!(!props.is_table_trailing_mark);
        assert_eq!(props.itap, 0);
        assert!(props.ilvl.is_none());
        assert!(props.ilfo.is_none());
        assert!(props.tap.is_none());
    }

    #[test]
    fn test_decodes_pchg_tabs_new_stops() {
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
    fn test_pchg_tabs_malformed_operand_is_empty() {
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
    fn test_parses_tdef_table_boundaries() {
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
    fn test_parses_tdef_table_vertical_merge_flags() {
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
    fn test_rejects_truncated_tdef_table() {
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
    fn test_spra_and_sgc_fields() {
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
    // (from a blind review). Every fixture uses the opcode *as Word writes it
    // per [MS-DOC]*, not the project's own constants, so the tests fail while
    // the decoder mislabels opcodes and turn green once dispatch is corrected.
    // --------------------------------------------------------------------

    /// `0x460B` is `sprmPIlfo` (2-byte operand). The decoder must populate
    /// `ilfo`. Today it is dispatched as `sprmPIlvl`, so `ilfo` stays `None`.
    #[test]
    fn test_sprm_pilfo_opcode_460b_populates_ilfo() {
        // sprmPIlfo (0x460B), operand = 0x0005 (ilfo index 5).
        let grpprl = [0x0B, 0x46, 0x05, 0x00];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.ilfo, Some(5), "0x460B is sprmPIlfo: must set ilfo");
        assert_eq!(props.ilvl, None, "0x460B must not be read as ilvl");
    }

    /// `0x260A` is `sprmPIlvl` (1-byte operand). The decoder must populate
    /// `ilvl`. Today it is never read (falls through to `_`).
    #[test]
    fn test_sprm_pilvl_opcode_260a_populates_ilvl() {
        // sprmPIlvl (0x260A), operand = 0x01 (level 1).
        let grpprl = [0x0A, 0x26, 0x01];
        let props = extract_pap_props(&grpprl);
        assert_eq!(props.ilvl, Some(1), "0x260A is sprmPIlvl: must set ilvl");
    }

    /// `0xC615` is `sprmPChgTabs`. The decoder must populate tab stops from it.
    #[test]
    fn test_sprm_pchg_tabs_opcode_c615_populates_tabs() {
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
    fn test_sprm_pchg_tabs_c615_cb_255_escape_round_trips() {
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
    fn test_sprm_pchg_tabs_papx_c60d_populates_tabs() {
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
    fn test_sprm_tcell_padding_opcode_d632_does_not_populate_tabs() {
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
    fn test_sprm_pchg_tabs_c615_truncated_cb_is_empty() {
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
    fn test_sprm_pchg_tabs_c615_255_escape_followed_by_sprm() {
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
    fn test_parse_grpprl_truncated_variable_sprm_prefixes() {
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
    fn test_pchg_tabs_operand_len_truncated() {
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
    fn test_opcode_conformance_gate() {
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

    /// [MS-DOC] transcription for the CHP registry, mirroring
    /// `MS_DOC_SPRM_TABLE` above.
    const MS_DOC_CHP_SPRM_TABLE: &[(u16, &str)] = &[
        (0x0800, "sprmCFRMarkDel"),
        (0x0801, "sprmCFRMarkIns"),
        (0x0835, "sprmCFBold"),
        (0x0836, "sprmCFItalic"),
        (0x2A3E, "sprmCKul"),
        (0x6870, "sprmCCv"),
        (0x4A43, "sprmCHps"),
        (0x4A51, "sprmCRgFtc2"),
        (0x0837, "sprmCFStrike"),
        (0x2A53, "sprmCFDStrike"),
        (0x083C, "sprmCFVanish"),
        (0x083A, "sprmCFSmallCaps"),
        (0x083B, "sprmCFCaps"),
        (0x2A48, "sprmCIss"),
        (0x4845, "sprmCHpsPos"),
        (0x2A0C, "sprmCHighlight"),
        (0x2A42, "sprmCIco"),
        // Confusable neighbours, not decoded: sprmCFOutline, sprmCFShadow,
        // sprmCKcd, sprmCFSpec.
        (0x0838, "sprmCFOutline"),
        (0x0839, "sprmCFShadow"),
        (0x2A34, "sprmCKcd"),
        (0x0855, "sprmCFSpec"),
    ];

    #[test]
    fn test_chp_registry_matches_the_spec_table() {
        for &(opcode, name, section) in CHP_SPRM_REGISTRY {
            let expected = MS_DOC_CHP_SPRM_TABLE
                .iter()
                .find(|(o, _)| *o == opcode)
                .map(|(_, n)| *n);
            assert_eq!(
                expected,
                Some(name),
                "opcode 0x{opcode:04X} is decoded as {name}, but [MS-DOC] calls it {expected:?}"
            );
            assert!(
                section.starts_with("2.6"),
                "0x{opcode:04X} ({name}) must cite its [MS-DOC] §2.6.x section, got {section:?}"
            );
        }
    }

    /// Mirrors `test_every_registered_opcode_changes_the_decoded_props`, but for
    /// the CHP registry. `sprmCHps` needs its own payload since a generic
    /// `[1, 0, 0, 0]` operand (font size `1` half-point) falls outside its
    /// valid `2..=3276` range and would be silently ignored, not stored.
    #[test]
    fn test_every_registered_chp_opcode_changes_the_decoded_props() {
        for &(opcode, name, _) in CHP_SPRM_REGISTRY {
            let mut grpprl = opcode.to_le_bytes().to_vec();
            if opcode == 0x4A43 {
                grpprl.extend_from_slice(&[48u8, 0]); // 24pt, in range
            } else {
                grpprl.extend_from_slice(&[1u8, 0, 0, 0]);
            }
            let props = extract_chp_props(&grpprl);
            assert_ne!(
                format!("{props:?}"),
                format!("{:?}", ChpProps::default()),
                "0x{opcode:04X} ({name}) is registered but decoding it changes nothing"
            );
        }
    }

    /// `sprmCRgFtc2` is a signed 2-byte index into the font table; a negative
    /// operand cannot be one, and a short operand carries none.
    #[test]
    fn test_font_index_for_text_outside_ascii_is_read() {
        assert_eq!(extract_chp_props(&[0x51, 0x4A, 0x03, 0x00]).ftc_other, Some(3));
        assert_eq!(extract_chp_props(&[0x51, 0x4A, 0xFF, 0x7F]).ftc_other, Some(0x7FFF));
        assert_eq!(extract_chp_props(&[0x51, 0x4A, 0xFF, 0xFF]).ftc_other, None);
        assert_eq!(extract_chp_props(&[0x4F, 0x4A, 0x01, 0x00]).ftc_other, None); // sprmCRgFtc0 is not it
        assert_eq!(ChpProps::default().ftc_other, None);
    }

    #[test]
    fn test_chp_registry_has_no_duplicate_opcodes() {
        let mut seen: Vec<u16> = CHP_SPRM_REGISTRY.iter().map(|(o, _, _)| *o).collect();
        seen.sort_unstable();
        let before = seen.len();
        seen.dedup();
        assert_eq!(before, seen.len(), "duplicate opcode in the CHP dispatch");
    }

    #[test]
    fn test_extract_chp_props_marks_deleted_and_inserted_runs() {
        // sprmCFRMarkDel = 1
        let props = extract_chp_props(&[0x00, 0x08, 0x01]);
        assert!(props.f_rmark_del);
        assert!(!props.f_rmark_ins);

        // sprmCFRMarkIns = 1
        let props = extract_chp_props(&[0x01, 0x08, 0x01]);
        assert!(!props.f_rmark_del);
        assert!(props.f_rmark_ins);

        // A zero operand toggles the flag off, not on.
        let props = extract_chp_props(&[0x00, 0x08, 0x00]);
        assert!(!props.f_rmark_del);
    }

    #[test]
    fn test_extract_chp_props_empty_grpprl_is_default() {
        assert_eq!(extract_chp_props(&[]), ChpProps::default());
    }

    #[test]
    fn test_extract_chp_props_bold_and_italic() {
        let props = extract_chp_props(&[0x35, 0x08, 0x01, 0x36, 0x08, 0x01]);
        assert!(props.bold);
        assert!(props.italic);

        // Zero operand toggles off, not on.
        let props = extract_chp_props(&[0x35, 0x08, 0x00]);
        assert!(!props.bold);
    }

    #[test]
    fn test_extract_chp_props_underline_maps_kul_values() {
        assert_eq!(extract_chp_props(&[0x3E, 0x2A, 0x00]).underline, None);
        assert_eq!(extract_chp_props(&[0x3E, 0x2A, 0x01]).underline, Some(UnderlineStyle::Single));
        assert_eq!(extract_chp_props(&[0x3E, 0x2A, 0x03]).underline, Some(UnderlineStyle::Double));
        assert_eq!(extract_chp_props(&[0x3E, 0x2A, 0x0B]).underline, Some(UnderlineStyle::Wave));
        // An unrecognized/reserved value still yields "some underline",
        // not silence.
        assert_eq!(extract_chp_props(&[0x3E, 0x2A, 0x63]).underline, Some(UnderlineStyle::Single));
    }

    #[test]
    fn test_extract_chp_props_color_ignores_auto() {
        // fAuto = 0: real RGB color.
        let props = extract_chp_props(&[0x70, 0x68, 0xFF, 0x99, 0x00, 0x00]);
        assert_eq!(props.color, Some([0xFF, 0x99, 0x00]));

        // fAuto = 1: "use the automatic color" — no override.
        let props = extract_chp_props(&[0x70, 0x68, 0xFF, 0x99, 0x00, 0x01]);
        assert_eq!(props.color, None);
    }

    #[test]
    fn test_extract_chp_props_font_size_rejects_out_of_range() {
        // 24pt = 48 half-points.
        let props = extract_chp_props(&[0x43, 0x4A, 48, 0]);
        assert_eq!(props.font_size_half_pt, Some(48));

        // 0 is below the spec's minimum of 2 — ignored, not stored.
        let props = extract_chp_props(&[0x43, 0x4A, 0, 0]);
        assert_eq!(props.font_size_half_pt, None);
    }

    /// A ToggleOperand ([MS-DOC] §2.9.339) is `0x00` off, `0x01` on, `0x80`
    /// "the style's value" and `0x81` "the opposite of the style's value".
    /// Decoding every non-zero operand as *on* read `0x80` as on: visible
    /// text Word shows (its style is not hidden) was treated as hidden and
    /// dropped, and `0x80` bold/italic runs came out bold/italic. As in
    /// POI's `getFlag`, `0x80` keeps the value in force and `0x81` negates
    /// it (with no style value applied here, that value is "off").
    #[test]
    fn test_toggle_operands_0x80_and_0x81_are_relative_to_the_style() {
        for opcode in [
            [0x35, 0x08], // sprmCFBold
            [0x36, 0x08], // sprmCFItalic
            [0x37, 0x08], // sprmCFStrike
            [0x53, 0x2A], // sprmCFDStrike
            [0x3B, 0x08], // sprmCFCaps
            [0x3A, 0x08], // sprmCFSmallCaps
            [0x00, 0x08], // sprmCFRMarkDel
        ] {
            let with = |operand: u8| extract_chp_props(&[opcode[0], opcode[1], operand]);
            let on = |p: ChpProps| {
                p.bold || p.italic || p.strike || p.all_caps || p.small_caps || p.f_rmark_del
            };
            assert!(!on(with(0x80)), "{opcode:02X?}: 0x80 over an off style is off");
            assert!(on(with(0x81)), "{opcode:02X?}: 0x81 over an off style is on");
            // 0x81 after an explicit on turns it back off; 0x80 keeps it.
            let on_then = |second: u8| {
                extract_chp_props(&[opcode[0], opcode[1], 0x01, opcode[0], opcode[1], second])
            };
            assert!(!on(on_then(0x81)), "{opcode:02X?}: 0x81 negates the value in force");
            assert!(on(on_then(0x80)), "{opcode:02X?}: 0x80 keeps the value in force");
        }
    }

    /// `sprmCFVanish` hides text, so a wrong *on* deletes content. Without
    /// the style sheet the value a `0x80`/`0x81` operand is relative to is
    /// unknown, and real documents carry `0x81` on ordinary visible form
    /// text (which earlier releases showed and Word displays). Only an explicit
    /// `0x01` hides; `0x80`/`0x81` leave the value in force unchanged.
    #[test]
    fn test_vanish_hides_only_on_an_explicit_operand() {
        assert!(extract_chp_props(&[0x3C, 0x08, 0x01]).vanish);
        assert!(!extract_chp_props(&[0x3C, 0x08, 0x81]).vanish);
        assert!(!extract_chp_props(&[0x3C, 0x08, 0x80]).vanish);
        assert!(extract_chp_props(&[0x3C, 0x08, 0x01, 0x3C, 0x08, 0x80]).vanish);
        assert!(!extract_chp_props(&[0x3C, 0x08, 0x01, 0x3C, 0x08, 0x00]).vanish);
    }

    /// Hidden text, strikethrough, super/subscript, highlight and the
    /// 16-colour text colour were never decoded: hidden text surfaced as
    /// visible and the rest of the formatting was lost.
    #[test]
    fn test_extract_chp_props_decodes_visibility_and_decoration_sprms() {
        assert!(extract_chp_props(&[0x3C, 0x08, 0x01]).vanish);
        assert!(extract_chp_props(&[0x3C, 0x08, 0x01]).is_excluded());
        assert!(!extract_chp_props(&[0x3C, 0x08, 0x00]).vanish);
        assert!(extract_chp_props(&[0x37, 0x08, 0x01]).strike);
        assert!(extract_chp_props(&[0x53, 0x2A, 0x01]).strike);
        assert!(extract_chp_props(&[0x3B, 0x08, 0x01]).all_caps);
        assert!(extract_chp_props(&[0x3A, 0x08, 0x01]).small_caps);
        assert_eq!(
            extract_chp_props(&[0x48, 0x2A, 0x01]).vertical_align(),
            Some(VerticalAlign::Superscript)
        );
        assert_eq!(
            extract_chp_props(&[0x48, 0x2A, 0x02]).vertical_align(),
            Some(VerticalAlign::Subscript)
        );
        assert_eq!(extract_chp_props(&[0x48, 0x2A, 0x00]).vertical_align(), None);
        // sprmCHpsPos: +6 half-points raised, -6 lowered.
        assert_eq!(
            extract_chp_props(&[0x45, 0x48, 0x06, 0x00]).vertical_align(),
            Some(VerticalAlign::Superscript)
        );
        assert_eq!(
            extract_chp_props(&[0x45, 0x48, 0xFA, 0xFF]).vertical_align(),
            Some(VerticalAlign::Subscript)
        );
        // sprmCIss wins over sprmCHpsPos.
        assert_eq!(
            extract_chp_props(&[0x45, 0x48, 0x06, 0x00, 0x48, 0x2A, 0x02]).vertical_align(),
            Some(VerticalAlign::Subscript)
        );
        assert_eq!(extract_chp_props(&[0x0C, 0x2A, 0x07]).highlight, Some([0xFF, 0xFF, 0x00]));
        assert_eq!(extract_chp_props(&[0x0C, 0x2A, 0x00]).highlight, None);
        assert_eq!(extract_chp_props(&[0x42, 0x2A, 0x06]).color, Some([0xFF, 0x00, 0x00]));
        // A later sprmCCv overrides the Ico colour.
        assert_eq!(
            extract_chp_props(&[0x42, 0x2A, 0x06, 0x70, 0x68, 0x12, 0x34, 0x56, 0x00]).color,
            Some([0x12, 0x34, 0x56])
        );
    }

    /// Every `Prm0` entry maps to a CHP opcode this crate decodes, with a
    /// 1-byte operand (spra 0 or 1), as the `Prm0` table requires.
    #[test]
    fn test_prm0_table_names_decoded_one_byte_chp_sprms() {
        for &(isprm, opcode) in PRM0_CHP_SPRMS {
            assert!(
                CHP_SPRM_REGISTRY.iter().any(|(o, _, _)| *o == opcode),
                "isprm 0x{isprm:02X} maps to undecoded opcode 0x{opcode:04X}"
            );
            assert!((opcode >> 13) <= 1, "0x{opcode:04X} does not take a 1-byte operand");
        }
        assert_eq!(prm0_sprm(0x55), Some(0x0835));
        assert_eq!(prm0_sprm(0x00), None);
    }

    #[test]
    fn test_ico_table_matches_the_spec() {
        assert_eq!(ico_to_rgb(0), None);
        assert_eq!(ico_to_rgb(0x02), Some([0, 0, 0xFF]));
        assert_eq!(ico_to_rgb(0x10), Some([0xC0, 0xC0, 0xC0]));
        assert_eq!(ico_to_rgb(0x11), None);
    }

    /// Every `Ico` index ([MS-DOC] §2.9.119) decodes to its own colour
    /// through both SPRMs that carry one — text colour (`sprmCIco`) and
    /// highlight (`sprmCHighlight`) — so a dropped or swapped table entry
    /// fails here rather than colouring runs wrongly.
    #[test]
    fn test_every_ico_index_decodes_to_its_spec_colour() {
        let table: [(u8, [u8; 3]); 16] = [
            (0x01, [0x00, 0x00, 0x00]),
            (0x02, [0x00, 0x00, 0xFF]),
            (0x03, [0x00, 0xFF, 0xFF]),
            (0x04, [0x00, 0xFF, 0x00]),
            (0x05, [0xFF, 0x00, 0xFF]),
            (0x06, [0xFF, 0x00, 0x00]),
            (0x07, [0xFF, 0xFF, 0x00]),
            (0x08, [0xFF, 0xFF, 0xFF]),
            (0x09, [0x00, 0x00, 0x80]),
            (0x0A, [0x00, 0x80, 0x80]),
            (0x0B, [0x00, 0x80, 0x00]),
            (0x0C, [0x80, 0x00, 0x80]),
            (0x0D, [0x80, 0x00, 0x00]),
            (0x0E, [0x80, 0x80, 0x00]),
            (0x0F, [0x80, 0x80, 0x80]),
            (0x10, [0xC0, 0xC0, 0xC0]),
        ];
        for (ico, rgb) in table {
            assert_eq!(extract_chp_props(&[0x42, 0x2A, ico]).color, Some(rgb), "sprmCIco {ico}");
            assert_eq!(
                extract_chp_props(&[0x0C, 0x2A, ico]).highlight,
                Some(rgb),
                "sprmCHighlight {ico}"
            );
        }
        for ico in [0x00, 0x11, 0xFF] {
            assert_eq!(extract_chp_props(&[0x42, 0x2A, ico]).color, None, "ico {ico}");
        }
    }

    /// A zero baseline offset is neither raised nor lowered: only a
    /// strictly positive `hps_pos` is superscript and only a strictly
    /// negative one subscript.
    #[test]
    fn test_zero_baseline_offset_has_no_vertical_align() {
        let at = |p: i16| {
            ChpProps {
                hps_pos: Some(p),
                ..Default::default()
            }
            .vertical_align()
        };
        assert_eq!(at(0), None);
        assert_eq!(at(1), Some(VerticalAlign::Superscript));
        assert_eq!(at(-1), Some(VerticalAlign::Subscript));
    }

    // ── Operand framing and tab-stop decoding, behaviourally ──

    /// `spra` 7 ([MS-DOC] §2.2.5.1) is a 3-byte operand: the walker must
    /// consume exactly three bytes and land on the next SPRM.
    #[test]
    fn test_spra_seven_operand_is_three_bytes() {
        let grpprl = [0x01, 0xE6, 0xAA, 0xBB, 0xCC, 0x16, 0x24, 0x01];
        let sprms = parse_grpprl(&grpprl);
        assert_eq!(sprms.len(), 2, "{sprms:?}");
        assert_eq!(sprms[0].opcode, 0xE601);
        assert_eq!(sprms[0].operand, [0xAA, 0xBB, 0xCC]);
        assert_eq!(sprms[1].opcode, 0x2416);
        assert_eq!(sprms[1].operand, [0x01]);
    }

    /// Each fixed `spra` class consumes its own width: 1, 1, 2, 4, 2, 2.
    #[test]
    fn test_every_fixed_spra_class_consumes_its_width() {
        for (spra, width) in [
            (0u16, 1usize),
            (1, 1),
            (2, 2),
            (3, 4),
            (4, 2),
            (5, 2),
            (7, 3),
        ] {
            let opcode = (spra << 13) | 0x0001;
            let mut g = opcode.to_le_bytes().to_vec();
            g.extend(std::iter::repeat_n(0x5A, width));
            g.extend_from_slice(&[0x16, 0x24, 0x01]);
            let sprms = parse_grpprl(&g);
            assert_eq!(sprms[0].operand.len(), width, "spra {spra}");
            assert_eq!(sprms[1].opcode, 0x2416, "spra {spra} must land on the next SPRM");
        }
    }

    /// A variable SPRM whose length prefix is the last byte (a zero-length
    /// operand exactly at the end) is still a SPRM, not a truncation.
    #[test]
    fn test_variable_sprm_ending_exactly_at_the_grpprl_end_is_kept() {
        // 1-byte prefix (sprmPChgTabsPapx, cb = 0).
        let sprms = parse_grpprl(&[0x0D, 0xC6, 0x00]);
        assert_eq!(
            sprms,
            [Sprm {
                opcode: 0xC60D,
                operand: vec![]
            }]
        );
        // sprmPChgTabs' own 1-byte cb path.
        let sprms = parse_grpprl(&[0x15, 0xC6, 0x00]);
        assert_eq!(
            sprms,
            [Sprm {
                opcode: 0xC615,
                operand: vec![]
            }]
        );
        // sprmTDefTable's 2-byte cb (cb = 1: empty operand).
        let sprms = parse_grpprl(&[0x08, 0xD6, 0x01, 0x00]);
        assert_eq!(
            sprms,
            [Sprm {
                opcode: 0xD608,
                operand: vec![]
            }]
        );
        // One byte short of the prefix: nothing.
        assert!(parse_grpprl(&[0x08, 0xD6, 0x01]).is_empty());
        assert!(parse_grpprl(&[0x0D, 0xC6]).is_empty());
    }

    /// `sprmPChgTabs` with the `cb` = 255 escape: the operand length is
    /// `1 + 4·cDel + 1 + 3·cAdd`, derived from the payload, and the next
    /// SPRM follows it exactly.
    #[test]
    fn test_pchg_tabs_escape_length_counts_both_lists() {
        let mut op = vec![2u8]; // cDel
        op.extend_from_slice(&[0x10, 0, 0x20, 0, 0x30, 0, 0x40, 0]); // rgdxaDel + rgdxaClose
        op.push(3); // cAdd
        op.extend_from_slice(&[0x00, 0x01, 0x00, 0x02, 0x00, 0x03]); // rgdxaAdd
        op.extend_from_slice(&[0x01, 0x02, 0x03]); // rgtbdAdd
        assert_eq!(op.len(), 1 + 4 * 2 + 1 + 3 * 3);
        assert_eq!(pchg_tabs_operand_len(&op, 0), op.len());
        let mut g = vec![0x15, 0xC6, 255];
        g.extend_from_slice(&op);
        g.extend_from_slice(&[0x16, 0x24, 0x01]);
        let sprms = parse_grpprl(&g);
        assert_eq!(sprms[0].operand, op);
        assert_eq!(sprms[1].opcode, 0x2416);
        // cAdd past the end: the rest of the grpprl.
        assert_eq!(pchg_tabs_operand_len(&[5, 0, 0], 0), 3);
        assert_eq!(pchg_tabs_operand_len(&[9, 9, 9], 1), 2);
    }

    /// Several added tabs decode to their own positions and descriptors,
    /// in order, for both delete-list framings.
    #[test]
    fn test_pchg_tabs_decodes_every_added_tab_in_order() {
        for (opcode, del_stride) in [(0xC615u16, 4usize), (0xC60D, 2)] {
            let mut op = vec![1u8];
            op.extend(std::iter::repeat_n(0x77, del_stride));
            op.push(3);
            for pos in [720i16, 1440, -360] {
                op.extend_from_slice(&pos.to_le_bytes());
            }
            op.extend_from_slice(&[0x01, 0x0A, 0x13]); // center; right+dot; decimal+hyphen
            let tabs = decode_pchg_tabs(opcode, &op);
            let got: Vec<(i32, TabAlignment, TabLeader)> = tabs
                .iter()
                .map(|t| (t.position_twips, t.alignment, t.leader))
                .collect();
            assert_eq!(
                got,
                [
                    (720, TabAlignment::Center, TabLeader::None),
                    (1440, TabAlignment::Right, TabLeader::Dot),
                    (-360, TabAlignment::Decimal, TabLeader::Hyphen),
                ],
                "opcode {opcode:#06x}"
            );
        }
    }

    /// A truncated add list stops cleanly — no read past the operand.
    #[test]
    fn test_pchg_tabs_truncated_positions_do_not_overrun() {
        // cDel 0, cAdd 2, one byte of the first position.
        assert!(decode_pchg_tabs(0xC60D, &[0, 2, 0x10]).is_empty());
        // Positions present, descriptors missing.
        assert!(decode_pchg_tabs(0xC60D, &[0, 1, 0x10, 0x00]).is_empty());
        // Second position truncated: only the first tab.
        let tabs = decode_pchg_tabs(0xC60D, &[0, 2, 0x10, 0x00, 0x20]);
        assert!(tabs.len() <= 1);
    }

    /// Every `jc` (bits 0..2) and `tlc` (bits 3..5) value of a `TBD`
    /// ([MS-DOC] §2.9.310), with the reserved high bits set to prove they
    /// are masked off.
    #[test]
    fn test_tab_descriptor_decodes_every_alignment_and_leader() {
        let jcs = [
            TabAlignment::Left,
            TabAlignment::Center,
            TabAlignment::Right,
            TabAlignment::Decimal,
            TabAlignment::Bar,
            TabAlignment::Left,
            TabAlignment::Left,
            TabAlignment::Left,
        ];
        let tlcs = [
            TabLeader::None,
            TabLeader::Dot,
            TabLeader::Hyphen,
            TabLeader::Underscore,
            TabLeader::Heavy,
            TabLeader::MiddleDot,
            TabLeader::None,
            TabLeader::None,
        ];
        for (jc, want_align) in jcs.iter().enumerate() {
            for (tlc, want_leader) in tlcs.iter().enumerate() {
                let tbd = 0xC0 | (tlc as u8) << 3 | jc as u8;
                let tab = tab_from_tbd(100, tbd);
                assert_eq!(&tab.alignment, want_align, "jc {jc} tlc {tlc}");
                assert_eq!(&tab.leader, want_leader, "jc {jc} tlc {tlc}");
                assert_eq!(tab.position_twips, 100);
            }
        }
    }

    /// `sprmTDefTable`: `rgdxaCenter` has `itcMac + 1` entries and the
    /// 20-byte cell descriptors start right after it; each cell's `rgf`
    /// and width come from its own descriptor.
    #[test]
    fn test_tdef_table_reads_each_cell_descriptor_at_its_offset() {
        let mut op = vec![3u8];
        for c in [0i16, 1000, 2500, 4000] {
            op.extend_from_slice(&c.to_le_bytes());
        }
        for (rgf, w) in [(0x0060u16, 1000u16), (0x0020, 1500), (0x0000, 1500)] {
            op.extend_from_slice(&rgf.to_le_bytes());
            op.extend_from_slice(&w.to_le_bytes());
            op.extend_from_slice(&[0u8; 16]);
        }
        let tap = parse_tdef_table(&op).expect("parses");
        assert_eq!(tap.itc_mac, 3);
        assert_eq!(tap.centers, [0, 1000, 2500, 4000]);
        let cells: Vec<(u16, u16)> = tap.cells.iter().map(|c| (c.rgf, c.w_width)).collect();
        assert_eq!(cells, [(0x0060, 1000), (0x0020, 1500), (0x0000, 1500)]);
        // One byte short of the last descriptor: rejected, not misread.
        assert!(parse_tdef_table(&op[..op.len() - 1]).is_none());
    }

    #[test]
    fn test_extract_chp_props_ignores_unknown_opcodes() {
        // 0x2416 = sprmPFInTable — a PAP opcode, must not affect CHP props.
        let props = extract_chp_props(&[0x16, 0x24, 0x01]);
        assert_eq!(props, ChpProps::default());
    }

    /// A `Data` stream with `PrcData` (`cbGrpprl: i16` + grpprl) at `fc`.
    fn data_with(fc: usize, grpprl: &[u8]) -> Vec<u8> {
        let mut d = vec![0xEE; fc];
        d.extend_from_slice(&(grpprl.len() as i16).to_le_bytes());
        d.extend_from_slice(grpprl);
        d
    }

    fn huge(fc: u32) -> Vec<u8> {
        let mut g = SPRM_P_HUGE_PAPX.to_le_bytes().to_vec();
        g.extend_from_slice(&fc.to_le_bytes());
        g
    }

    #[test]
    fn test_expand_huge_papx_absent_returns_none() {
        assert_eq!(
            expand_huge_papx(&[0x16, 0x24, 0x01], &[0u8; 8], &mut HugeBudget::new(1_000_000)),
            None
        );
        assert_eq!(expand_huge_papx(&[], &[], &mut HugeBudget::new(1_000_000)), None);
    }

    #[test]
    fn test_expand_huge_papx_splices_in_place_keeping_order() {
        let prc = [0x17, 0x24, 0x01]; // sprmPFTtp = 1
        let data = data_with(6, &prc);
        let mut g = vec![0x16, 0x24, 0x01]; // sprmPFInTable = 1 (before)
        g.extend_from_slice(&huge(6));
        g.extend_from_slice(&[0x49, 0x66, 0x01, 0, 0, 0]); // sprmPItap (after)
        let (out, bad) =
            expand_huge_papx(&g, &data, &mut HugeBudget::new(1_000_000)).expect("huge present");
        assert_eq!(bad, 0);
        let ops: Vec<u16> = parse_grpprl(&out).iter().map(|s| s.opcode).collect();
        assert_eq!(ops, vec![0x2416, 0x2417, 0x6649]);
    }

    #[test]
    fn test_expand_huge_papx_counts_and_drops_unresolvable_references() {
        let data = data_with(0, &[0x17, 0x24, 0x01]);
        for fc in [0x1000_0000u32, u32::MAX] {
            let (out, bad) =
                expand_huge_papx(&huge(fc), &data, &mut HugeBudget::new(1_000_000)).unwrap();
            assert!(out.is_empty());
            assert_eq!(bad, 1);
        }
        // cbGrpprl longer than the stream, negative, and above the spec cap.
        let mut short = vec![0u8; 2];
        short.extend_from_slice(&[1, 2, 3]);
        short[..2].copy_from_slice(&100i16.to_le_bytes());
        assert_eq!(
            expand_huge_papx(&huge(0), &short, &mut HugeBudget::new(1_000_000))
                .unwrap()
                .1,
            1
        );
        short[..2].copy_from_slice(&(-5i16).to_le_bytes());
        assert_eq!(
            expand_huge_papx(&huge(0), &short, &mut HugeBudget::new(1_000_000))
                .unwrap()
                .1,
            1
        );
        let mut big = vec![0u8; 2 + 0x3FA3];
        big[..2].copy_from_slice(&0x3FA3i16.to_le_bytes());
        assert_eq!(
            expand_huge_papx(&huge(0), &big, &mut HugeBudget::new(1_000_000))
                .unwrap()
                .1,
            1
        );
    }

    #[test]
    fn test_expand_huge_papx_does_not_follow_a_reference_inside_prc_data() {
        let data = data_with(4, &huge(4)); // PrcData that names itself
        let (out, bad) =
            expand_huge_papx(&huge(4), &data, &mut HugeBudget::new(1_000_000)).unwrap();
        assert_eq!(bad, 0);
        assert_eq!(out, huge(4), "the inner reference is copied, not followed");
    }

    #[test]
    fn test_expand_huge_papx_truncated_operand_is_unresolved_not_a_panic() {
        // Opcode with only two operand bytes left.
        let mut g = SPRM_P_HUGE_PAPX.to_le_bytes().to_vec();
        g.extend_from_slice(&[0x01, 0x00]);
        let (out, bad) = expand_huge_papx(&g, &[0u8; 16], &mut HugeBudget::new(1_000_000)).unwrap();
        assert_eq!((out.len(), bad), (0, 1));
    }

    #[test]
    fn test_table_header_sprm_decodes_the_flag_bit() {
        assert!(extract_pap_props(&[0x04, 0x34, 0x01]).f_table_header);
        assert!(!extract_pap_props(&[0x04, 0x34, 0x00]).f_table_header);
        assert!(!extract_pap_props(&[0x03, 0x34, 0x01]).f_table_header, "0x3403 is fCantSplit");
    }

    #[test]
    fn test_expand_huge_papx_charges_every_copy_and_counts_the_rest() {
        let prc = [0x17, 0x24, 0x01, 0x16, 0x24, 0x01]; // 6 bytes
        let data = data_with(0, &prc);
        let mut budget = HugeBudget::new(10); // room for one copy, not two
        let mut g = huge(0);
        g.extend_from_slice(&huge(0)); // the same PrcData named twice
        let (out, bad) = expand_huge_papx(&g, &data, &mut budget).unwrap();
        assert_eq!(bad, 1, "the second copy is charged too");
        assert_eq!(out.len(), prc.len());
        assert_eq!(budget.remaining, 4);
        // Another list naming the same PrcData is charged again: 4 left, 6 wanted.
        let (out, bad) = expand_huge_papx(&huge(0), &data, &mut budget).unwrap();
        assert_eq!((out.len(), bad), (0, 1));
    }

    #[test]
    fn test_expand_huge_papx_accepts_the_word97_opcode() {
        let data = data_with(2, &[0x17, 0x24, 0x01]);
        let mut g = vec![0x45, 0x66];
        g.extend_from_slice(&2u32.to_le_bytes());
        let (out, bad) = expand_huge_papx(&g, &data, &mut HugeBudget::new(1_000_000)).unwrap();
        assert_eq!((out, bad), (vec![0x17, 0x24, 0x01], 0));
    }
}
