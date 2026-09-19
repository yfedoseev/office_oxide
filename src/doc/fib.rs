//! File Information Block (FIB) parsing.
//!
//! The FIB is at the start of the WordDocument stream. It contains metadata
//! and pointers to other structures in the Table stream.

use super::error::{DocError, Result};

/// Parsed FIB fields needed for text extraction.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct Fib {
    /// Word version identifier.
    pub version: u16,
    /// FibBase `lid`: the document's default language ID (Windows LCID),
    /// e.g. `0x0419` = Russian. Determines which codepage compressed
    /// (8-bit) text runs are stored in — before this field existed,
    /// every compressed run was decoded as CP1252 regardless of the
    /// document's actual authoring locale (issue #310).
    pub lid: u16,
    /// Which table stream to use: true = "1Table", false = "0Table".
    pub use_table1: bool,
    /// Offset of the CLX (piece table) in the Table stream.
    pub clx_offset: u32,
    /// Size of the CLX in the Table stream.
    pub clx_size: u32,
    /// Total length of text in the main document (in characters).
    pub text_len: u32,
    /// Length of footnote text.
    pub footnote_len: u32,
    /// Length of header/footer text.
    pub header_len: u32,
    /// Length of comment text.
    pub comment_len: u32,
    /// Length of endnote text.
    pub endnote_len: u32,
    /// Length of textbox text.
    pub textbox_len: u32,
    /// Length of header textbox text.
    pub header_textbox_len: u32,
    /// Offset of the PlcfBtePapx (PAPX FKP index) in the Table stream.
    /// FIB absolute offset 0x0102. Zero when the file has no PAPX FKP.
    pub fc_plcf_bte_papx: u32,
    /// Byte length of the PlcfBtePapx in the Table stream (0x0106).
    pub lcb_plcf_bte_papx: u32,
    /// Offset of the PlfLst (list definitions: LSTF + LVL arrays) in the
    /// Table stream (0x02E2). Zero when the file defines no lists. Named
    /// `fc_plcf_lst` here (not `fc_plf_lst`) for historical reasons — the
    /// struct/field spelling predates issue #250, which is the first
    /// consumer of the pointer it holds.
    pub fc_plcf_lst: u32,
    /// Byte length of the PlfLst in the Table stream (0x02E6).
    pub lcb_plcf_lst: u32,
    /// Offset of the PlfLfo (list format overrides: LFO array, one per
    /// `ilfo`) in the Table stream (0x02EA). `sprmPIlfo`'s value is a
    /// 1-based index into this array, not into `PlfLst` directly — an
    /// `LFO.lsid` is what actually selects the matching `LSTF` (issue
    /// #250). Zero when the file uses no lists.
    pub fc_plf_lfo: u32,
    /// Byte length of the PlfLfo in the Table stream (0x02EE).
    pub lcb_plf_lfo: u32,
    /// Offset of the GrpXstAtnOwners (comment author name array) in the
    /// Table stream (0x01BA). Zero when the document has no comments.
    /// Issue #298.
    pub fc_grp_xst_atn_owners: u32,
    /// Byte length of the GrpXstAtnOwners in the Table stream (0x01BE).
    pub lcb_grp_xst_atn_owners: u32,
}

impl Fib {
    /// Parse the FIB from the WordDocument stream.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 68 {
            return Err(DocError::InvalidFib(format!(
                "WordDocument stream too short: {} bytes",
                data.len()
            )));
        }

        let wident = u16::from_le_bytes([data[0], data[1]]);
        // 0xA5EC = Word 97 and later. 0xA5DC = Word 6.0/95, whose FIB has a
        // completely different layout: every FibRgFcLcb97 offset below is
        // wrong for it. Accepting the file and then reading Word 97 offsets
        // out of it produced a confident empty result — a 426 KB document
        // extracted as the empty string with `Ok`. Say what it is instead.
        if wident == 0xA5DC {
            return Err(DocError::UnsupportedVersion(
                "Word 6.0/95 (wIdent 0xA5DC); only Word 97 and later are supported".into(),
            ));
        }
        if wident != 0xA5EC {
            return Err(DocError::InvalidFib(format!("unknown wIdent: 0x{wident:04X}")));
        }

        let version = u16::from_le_bytes([data[2], data[3]]);
        // FibBase.lid, absolute offset 0x06 (u16 LE) — verified against
        // FibBase's field layout (wIdent, nFib, unused, lid, ...).
        let lid = u16::from_le_bytes([data[0x06], data[0x07]]);

        // Flags at offset 0x0A (u16): [MS-DOC] §2.5.1 FibBase.
        //   bit 8  = fEncrypted
        //   bit 9  = fWhichTblStm
        let flags = u16::from_le_bytes([data[0x0A], data[0x0B]]);
        // An encrypted document's text is ciphertext. Walking the piece
        // table over it yields either nothing or mojibake, both reported as
        // a successful extraction of a document that "has no text".
        if (flags & (1 << 8)) != 0 {
            return Err(DocError::Encrypted);
        }
        let use_table1 = (flags & (1 << 9)) != 0;

        // FibRgLw97 starts at offset 0x22, its size field at 0x22 (u16, should be 0x16).
        // Text lengths in FibRgLw97 ([MS-DOC] FibRgLw97): cbMac, reserved1,
        // reserved2, ccpText, ccpFtn, ccpHdd, reserved3 (MUST be zero, MUST
        // be ignored — NOT ccpAtn), ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx.
        // Issue #247: every field from `comment_len` on used to be read one
        // slot early (`comment_len` landed on the always-zero `reserved3`,
        // so it silently read as 0 for every `.doc` ever opened; comments
        // ended up mislabeled as endnotes, endnotes as textboxes, and the
        // real `ccpHdrTxbx` — header-anchored textbox text — was never read
        // at all).
        let text_len = read_u32(data, 0x4C);
        let footnote_len = read_u32(data, 0x50);
        let header_len = read_u32(data, 0x54);
        // 0x58 = reserved3, MUST be zero, MUST be ignored — deliberately unread.
        let comment_len = read_u32(data, 0x5C);
        let endnote_len = read_u32(data, 0x60);
        let textbox_len = read_u32(data, 0x64);
        let header_textbox_len = read_u32(data, 0x68);

        // FibRgFcLcb97 is laid out at a fixed set of absolute offsets within
        // the WordDocument stream for Word 97+ (nFib = 0x00C1). The offsets
        // below are absolute (measured from the start of the stream), which
        // matches what the on-disk FIB stores. See [MS-DOC] §2.5.5.
        //
        // fcClx / lcbClx — piece table pointer (absolute 0x01A2 / 0x01A6).
        let (clx_offset, clx_size) = if data.len() > 0x01AA {
            (read_u32(data, 0x01A2), read_u32(data, 0x01A6))
        } else {
            (0, 0)
        };

        // fcPlcfBtePapx / lcbPlcfBtePapx — PAPX FKP index (0x0102 / 0x0106).
        // Used to locate paragraph property (PAPX) pages, which carry the
        // fInTable / TAP (table) SPRMs needed for table reconstruction.
        let (fc_plcf_bte_papx, lcb_plcf_bte_papx) = if data.len() > 0x010A {
            (read_u32(data, 0x0102), read_u32(data, 0x0106))
        } else {
            (0, 0)
        };

        // fcPlfLst / lcbPlfLst — list definitions (0x02E2 / 0x02E6).
        // Zero when the document defines no lists.
        let (fc_plcf_lst, lcb_plcf_lst) = if data.len() > 0x02EA {
            (read_u32(data, 0x02E2), read_u32(data, 0x02E6))
        } else {
            (0, 0)
        };

        // fcPlfLfo / lcbPlfLfo — list format overrides (0x02EA / 0x02EE),
        // immediately following the fcPlfLst/lcbPlfLst pair above. Verified
        // against the published [MS-DOC] worked example ("Example of a
        // List"), which shows both pairs at these exact offsets.
        let (fc_plf_lfo, lcb_plf_lfo) = if data.len() > 0x02F2 {
            (read_u32(data, 0x02EA), read_u32(data, 0x02EE))
        } else {
            (0, 0)
        };

        // fcGrpXstAtnOwners / lcbGrpXstAtnOwners — comment author names
        // (0x01BA / 0x01BE). Derived from FibRgFcLcb97's fixed field order
        // ([MS-DOC] §2.5.5): each entry is 4 bytes, starting at absolute
        // offset 0x9A (fcStshfOrig); fcGrpXstAtnOwners is the 73rd entry
        // (0-based index 72), landing at 0x9A + 72*4 = 0x1BA. Cross-checked
        // against the two already-verified anchors above: fcPlcfBtePapx
        // (index 26) lands at 0x9A + 26*4 = 0x102, and fcClx (index 66)
        // lands at 0x9A + 66*4 = 0x1A2 — both match their hard-coded
        // offsets already in this file. Issue #298.
        let (fc_grp_xst_atn_owners, lcb_grp_xst_atn_owners) = if data.len() > 0x01C2 {
            (read_u32(data, 0x01BA), read_u32(data, 0x01BE))
        } else {
            (0, 0)
        };

        Ok(Self {
            version,
            lid,
            use_table1,
            clx_offset,
            clx_size,
            text_len,
            footnote_len,
            header_len,
            comment_len,
            endnote_len,
            textbox_len,
            header_textbox_len,
            fc_plcf_bte_papx,
            lcb_plcf_bte_papx,
            fc_plcf_lst,
            lcb_plcf_lst,
            fc_plf_lfo,
            lcb_plf_lfo,
            fc_grp_xst_atn_owners,
            lcb_grp_xst_atn_owners,
        })
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    if offset + 4 <= data.len() {
        u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ])
    } else {
        0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_minimal_fib() -> Vec<u8> {
        let mut data = vec![0u8; 1024];
        // wIdent = Word 97
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        // nFib (version)
        data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        // flags: use 1Table (bit 9)
        data[0x0A..0x0C].copy_from_slice(&(1u16 << 9).to_le_bytes());
        // ccpText = 100
        data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes());
        // fcClx
        data[0x01A2..0x01A6].copy_from_slice(&512u32.to_le_bytes());
        // lcbClx
        data[0x01A6..0x01AA].copy_from_slice(&64u32.to_le_bytes());
        // fcPlcfBtePapx = 300, lcbPlcfBtePapx = 28
        data[0x0102..0x0106].copy_from_slice(&300u32.to_le_bytes());
        data[0x0106..0x010A].copy_from_slice(&28u32.to_le_bytes());
        // fcPlfLst = 400, lcbPlfLst = 12
        data[0x02E2..0x02E6].copy_from_slice(&400u32.to_le_bytes());
        data[0x02E6..0x02EA].copy_from_slice(&12u32.to_le_bytes());
        // fcPlfLfo / lcbPlfLfo — the exact values from [MS-DOC]'s own
        // "Example of a List" worked example (issue #250).
        data[0x02EA..0x02EE].copy_from_slice(&0x000007E1u32.to_le_bytes());
        data[0x02EE..0x02F2].copy_from_slice(&0x00000018u32.to_le_bytes());
        data
    }

    #[test]
    fn parse_valid_fib() {
        let data = build_minimal_fib();
        let fib = Fib::parse(&data).unwrap();
        assert_eq!(fib.version, 0x00C1);
        assert!(fib.use_table1);
        assert_eq!(fib.text_len, 100);
        assert_eq!(fib.clx_offset, 512);
        assert_eq!(fib.clx_size, 64);
        assert_eq!(fib.fc_plcf_bte_papx, 300);
        assert_eq!(fib.lcb_plcf_bte_papx, 28);
        assert_eq!(fib.fc_plcf_lst, 400);
        assert_eq!(fib.lcb_plcf_lst, 12);
        assert_eq!(fib.fc_plf_lfo, 0x000007E1);
        assert_eq!(fib.lcb_plf_lfo, 0x00000018);
    }

    #[test]
    fn bad_wident_rejected() {
        let mut data = build_minimal_fib();
        data[0..2].copy_from_slice(&0x1234u16.to_le_bytes());
        assert!(Fib::parse(&data).is_err());
    }

    #[test]
    fn too_short_rejected() {
        let data = vec![0u8; 100];
        assert!(Fib::parse(&data).is_err());
    }

    #[test]
    fn use_table0() {
        let mut data = build_minimal_fib();
        data[0x0A..0x0C].copy_from_slice(&0u16.to_le_bytes()); // clear bit 9
        let fib = Fib::parse(&data).unwrap();
        assert!(!fib.use_table1);
    }

    /// issue #247 — every `FibRgLw97` field from `comment_len` on used to be
    /// read one slot early (landing on `reserved3`, spec-mandated always
    /// zero, at 0x58) instead of its real offset. Each field below gets a
    /// distinct value so a shift in either direction is caught, and
    /// `reserved3` itself is set to a nonzero value to prove it's never
    /// read at all.
    #[test]
    fn fibrglw97_fields_read_from_their_real_spec_offsets() {
        let mut data = build_minimal_fib();
        data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes()); // ccpText
        data[0x50..0x54].copy_from_slice(&11u32.to_le_bytes()); // ccpFtn
        data[0x54..0x58].copy_from_slice(&22u32.to_le_bytes()); // ccpHdd
        data[0x58..0x5C].copy_from_slice(&0xDEADBEEFu32.to_le_bytes()); // reserved3 — MUST be ignored
        data[0x5C..0x60].copy_from_slice(&33u32.to_le_bytes()); // ccpAtn (comments)
        data[0x60..0x64].copy_from_slice(&44u32.to_le_bytes()); // ccpEdn (endnotes)
        data[0x64..0x68].copy_from_slice(&55u32.to_le_bytes()); // ccpTxbx (textboxes)
        data[0x68..0x6C].copy_from_slice(&66u32.to_le_bytes()); // ccpHdrTxbx (header textboxes)

        let fib = Fib::parse(&data).unwrap();
        assert_eq!(fib.text_len, 100);
        assert_eq!(fib.footnote_len, 11);
        assert_eq!(fib.header_len, 22);
        assert_eq!(fib.comment_len, 33, "must read ccpAtn at 0x5C, not reserved3 at 0x58");
        assert_eq!(fib.endnote_len, 44, "must read ccpEdn at 0x60");
        assert_eq!(fib.textbox_len, 55, "must read ccpTxbx at 0x64");
        assert_eq!(
            fib.header_textbox_len, 66,
            "must read the real ccpHdrTxbx at 0x68, previously never read at all"
        );
    }

    /// issue #298 — `fcGrpXstAtnOwners`/`lcbGrpXstAtnOwners` (comment
    /// author names) were never parsed at all. Offset derived from
    /// FibRgFcLcb97's fixed field order; cross-checked against the two
    /// offsets already verified elsewhere in this file (`fcPlcfBtePapx`
    /// at 0x0102, `fcClx` at 0x01A2).
    #[test]
    fn grp_xst_atn_owners_read_from_its_real_offset() {
        let mut data = build_minimal_fib();
        data[0x01BA..0x01BE].copy_from_slice(&700u32.to_le_bytes());
        data[0x01BE..0x01C2].copy_from_slice(&40u32.to_le_bytes());

        let fib = Fib::parse(&data).unwrap();
        assert_eq!(fib.fc_grp_xst_atn_owners, 700);
        assert_eq!(fib.lcb_grp_xst_atn_owners, 40);
    }
}
