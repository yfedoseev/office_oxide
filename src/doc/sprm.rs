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

use crate::ir::TabStop;

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
/// Only the SPRMs needed for table reconstruction are tracked; every
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
    /// Parsed row definition (`sprmTDefTable` operand) for row-terminator
    /// paragraphs. `None` when the paragraph is not a row mark or the TAP
    /// is malformed.
    pub tap: Option<TapInfo>,
    /// Tab stops from `sprmPChgTabs` (0xC615) / `sprmPChgTabsPapx` (0xC60D).
    /// Empty when the paragraph carries no tab-stop SPRM. Populated by the
    /// list/tab-stop PR; in the tables-only PR this field is always empty, but
    /// it is cloned through the IR so the `Paragraph.tabs` shape stays uniform.
    pub tabs: Vec<TabStop>,
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
        assert!(props.tap.is_none());
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
}
