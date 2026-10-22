//! Paragraph properties of Word 6.0/95 files: the PAPX that carries a table.
//!
//! Word 6/95 keeps paragraph properties in the same `WordDocument` stream as the
//! text (there is no table stream), in a layout that differs from Word 97's in
//! every detail this module reads:
//!
//! - the FIB's `fcPlcfbtePapx` / `lcbPlcfbtePapx` are at `0xC0` / `0xC4`;
//! - the bin table is `(n + 1)` 4-byte file positions followed by `n` **2-byte**
//!   FKP page numbers (Word 97 uses 4);
//! - an FKP page is `crun` at byte 511, `crun + 1` 4-byte file positions, then
//!   `crun` **7-byte** entries (a word offset and a 6-byte paragraph height,
//!   where Word 97 has 13);
//! - a PAPX is `cb` (in words), a 2-byte `istd`, and the property list;
//! - sprm numbers are **one byte**, and an operand's size is a property of the
//!   sprm number, taken from the table below, not from bits of the number.
//!
//! Only the sprms that make a table are acted on, but every sprm's size is
//! needed to step over the ones before them. The table is the Word for
//! Windows 6.0 Binary File Format's sprm table; its numbers and operand sizes
//! were cross-checked against the importer in LibreOffice (`ww8scan.cxx`) and
//! against 14,393 property lists in 32 real Word 6/95 files, 99.97% of which
//! it walks to exactly their end.
//!
//! The recognised sprms are re-encoded as the Word 97 sprms the paragraph
//! builder already understands, so a Word 6 table goes through the same table
//! walk as a Word 97 one.

use super::papx::FkpParagraph;
use super::word6::Word6Fib;

/// How an operand's size is found.
#[derive(Clone, Copy)]
enum Operand {
    /// A fixed number of bytes.
    Fixed(u8),
    /// A 1-byte length, then that many bytes.
    Var,
    /// A 2-byte length counting itself, then `length - 1` bytes.
    Var2,
}

// Word 6.0/95 sprm numbers ([Word for Windows 6.0 Binary File Format], the
// sprm table; the same numbers LibreOffice's `ww8scan.cxx` lists as `v6`):

/// `sprmPFInTable` (24): the paragraph is in a table.
const SPRM_P_F_IN_TABLE: u8 = 24;
/// `sprmPFTtp` (25): the paragraph ends a table row.
const SPRM_P_F_TTP: u8 = 25;
/// `sprmTTableHeader` (186): the row is a header row.
const SPRM_T_TABLE_HEADER: u8 = 186;
/// `sprmTDefTable` (190): the row definition (`itcMac`, the column boundaries
/// and a `TC` per cell); a 2-byte length counting itself, then `length - 1`
/// bytes.
const SPRM_T_DEF_TABLE: u8 = 190;
/// `sprmPChgTabs` (23) and `sprmPChgTabsPapx` (15): tab stops, whose length
/// byte of 255 means "derive the length from the contents".
const SPRM_P_CHG_TABS: u8 = 23;
const SPRM_P_CHG_TABS_PAPX: u8 = 15;

// The Word 97 sprms these are re-encoded as ([MS-DOC]):

/// `sprmPFInTable`, 0x2416 ([MS-DOC] §2.6.2).
const SPRM97_P_F_IN_TABLE: [u8; 2] = [0x16, 0x24];
/// `sprmTTableHeader`, 0x3404 ([MS-DOC] §2.6.3).
const SPRM97_T_TABLE_HEADER: [u8; 2] = [0x04, 0x34];
/// `sprmTDefTable`, 0xD608 ([MS-DOC] §2.6.3); a 2-byte length counting itself.
const SPRM97_T_DEF_TABLE: [u8; 2] = [0x08, 0xD6];

/// Operand sizes of the Word 6/95 sprms, by sprm number (sorted).
const SPRMS: &[(u8, Operand)] = &[
    (0, Operand::Fixed(0)),    // (unnamed)
    (2, Operand::Fixed(2)),    // sprmPIstd
    (3, Operand::Var),         // sprmPIstdPermute
    (4, Operand::Fixed(1)),    // sprmPIncLv1
    (5, Operand::Fixed(1)),    // sprmPJc
    (6, Operand::Fixed(1)),    // sprmPFSideBySide
    (7, Operand::Fixed(1)),    // sprmPFKeep
    (8, Operand::Fixed(1)),    // sprmPFKeepFollow
    (9, Operand::Fixed(1)),    // sprmPPageBreakBefore
    (10, Operand::Fixed(1)),   // sprmPBrcl
    (11, Operand::Fixed(1)),   // sprmPBrcp
    (12, Operand::Var),        // sprmPAnld
    (13, Operand::Fixed(1)),   // sprmPNLvlAnm
    (14, Operand::Fixed(1)),   // sprmPFNoLineNumb
    (15, Operand::Var),        // sprmPChgTabsPapx
    (16, Operand::Fixed(2)),   // sprmPDxaRight
    (17, Operand::Fixed(2)),   // sprmPDxaLeft
    (18, Operand::Fixed(2)),   // sprmPNest
    (19, Operand::Fixed(2)),   // sprmPDxaLeft1
    (20, Operand::Fixed(4)),   // sprmPDyaLine
    (21, Operand::Fixed(2)),   // sprmPDyaBefore
    (22, Operand::Fixed(2)),   // sprmPDyaAfter
    (23, Operand::Var),        // sprmPChgTabs
    (24, Operand::Fixed(1)),   // sprmPFInTable
    (25, Operand::Fixed(1)),   // sprmPTtp
    (26, Operand::Fixed(2)),   // sprmPDxaAbs
    (27, Operand::Fixed(2)),   // sprmPDyaAbs
    (28, Operand::Fixed(2)),   // sprmPDxaWidth
    (29, Operand::Fixed(1)),   // sprmPPc
    (30, Operand::Fixed(2)),   // sprmPBrcTop10
    (31, Operand::Fixed(2)),   // sprmPBrcLeft10
    (32, Operand::Fixed(2)),   // sprmPBrcBottom10
    (33, Operand::Fixed(2)),   // sprmPBrcRight10
    (34, Operand::Fixed(2)),   // sprmPBrcBetween10
    (35, Operand::Fixed(2)),   // sprmPBrcBar10
    (36, Operand::Fixed(2)),   // sprmPFromText10
    (37, Operand::Fixed(1)),   // sprmPWr
    (38, Operand::Fixed(2)),   // sprmPBrcTop
    (39, Operand::Fixed(2)),   // sprmPBrcLeft
    (40, Operand::Fixed(2)),   // sprmPBrcBottom
    (41, Operand::Fixed(2)),   // sprmPBrcRight
    (42, Operand::Fixed(2)),   // sprmPBrcBetween
    (43, Operand::Fixed(2)),   // sprmPBrcBar
    (44, Operand::Fixed(1)),   // sprmPFNoAutoHyph
    (45, Operand::Fixed(2)),   // sprmPWHeightAbs
    (46, Operand::Fixed(2)),   // sprmPDcs
    (47, Operand::Fixed(2)),   // sprmPShd
    (48, Operand::Fixed(2)),   // sprmPDyaFromText
    (49, Operand::Fixed(2)),   // sprmPDxaFromText
    (50, Operand::Fixed(1)),   // sprmPFLocked
    (51, Operand::Fixed(1)),   // sprmPFWidowControl
    (52, Operand::Fixed(0)),   // sprmPRuler
    (64, Operand::Var),        // (unnamed)
    (65, Operand::Fixed(1)),   // sprmCFStrikeRM
    (66, Operand::Fixed(1)),   // sprmCFRMark
    (67, Operand::Fixed(1)),   // sprmCFFldVanish
    (68, Operand::Var),        // sprmCPicLocation
    (69, Operand::Fixed(2)),   // sprmCIbstRMark
    (70, Operand::Fixed(4)),   // sprmCDttmRMark
    (71, Operand::Fixed(1)),   // sprmCFData
    (72, Operand::Fixed(2)),   // sprmCRMReason
    (73, Operand::Fixed(3)),   // sprmCChse
    (74, Operand::Var),        // sprmCSymbol
    (75, Operand::Fixed(1)),   // sprmCFOle2
    (77, Operand::Var),        // (unnamed)
    (79, Operand::Var),        // (unnamed)
    (80, Operand::Fixed(2)),   // sprmCIstd
    (81, Operand::Var),        // sprmCIstdPermute
    (82, Operand::Var),        // sprmCDefault
    (83, Operand::Fixed(0)),   // sprmCPlain
    (85, Operand::Fixed(1)),   // sprmCFBold
    (86, Operand::Fixed(1)),   // sprmCFItalic
    (87, Operand::Fixed(1)),   // sprmCFStrike
    (88, Operand::Fixed(1)),   // sprmCFOutline
    (89, Operand::Fixed(1)),   // sprmCFShadow
    (90, Operand::Fixed(1)),   // sprmCFSmallCaps
    (91, Operand::Fixed(1)),   // sprmCFCaps
    (92, Operand::Fixed(1)),   // sprmCFVanish
    (93, Operand::Fixed(2)),   // sprmCFtc
    (94, Operand::Fixed(1)),   // sprmCKul
    (95, Operand::Fixed(3)),   // sprmCSizePos
    (96, Operand::Fixed(2)),   // sprmCDxaSpace
    (97, Operand::Fixed(2)),   // sprmCLid
    (98, Operand::Fixed(1)),   // sprmCIco
    (99, Operand::Fixed(2)),   // sprmCHps
    (100, Operand::Fixed(1)),  // sprmCHpsInc
    (101, Operand::Fixed(2)),  // sprmCHpsPos
    (102, Operand::Fixed(1)),  // sprmCHpsPosAdj
    (103, Operand::Var),       // sprmCMajority
    (104, Operand::Fixed(1)),  // sprmCIss
    (105, Operand::Var),       // sprmCHpsNew50
    (106, Operand::Var),       // sprmCHpsInc1
    (107, Operand::Fixed(2)),  // sprmCHpsKern
    (108, Operand::Var),       // sprmCMajority50
    (109, Operand::Fixed(2)),  // sprmCHpsMul
    (110, Operand::Fixed(2)),  // sprmCCondHyhen
    (111, Operand::Var),       // (unnamed)
    (112, Operand::Var),       // (unnamed)
    (113, Operand::Var),       // (unnamed)
    (114, Operand::Var),       // (unnamed)
    (115, Operand::Var),       // (unnamed)
    (116, Operand::Var),       // (unnamed)
    (117, Operand::Fixed(1)),  // sprmCFSpec
    (118, Operand::Fixed(1)),  // sprmCFObj
    (119, Operand::Fixed(1)),  // sprmPicBrcl
    (120, Operand::Var),       // sprmPicScale
    (121, Operand::Fixed(2)),  // sprmPicBrcTop
    (122, Operand::Fixed(2)),  // sprmPicBrcLeft
    (123, Operand::Fixed(2)),  // sprmPicBrcBottom
    (124, Operand::Fixed(2)),  // sprmPicBrcRight
    (131, Operand::Fixed(1)),  // sprmSScnsPgn
    (132, Operand::Fixed(1)),  // sprmSiHeadingPgn
    (133, Operand::Var),       // sprmSOlstAnm
    (136, Operand::Fixed(3)),  // sprmSDxaColWidth
    (137, Operand::Fixed(3)),  // sprmSDxaColSpacing
    (138, Operand::Fixed(1)),  // sprmSFEvenlySpaced
    (139, Operand::Fixed(1)),  // sprmSFProtected
    (140, Operand::Fixed(2)),  // sprmSDmBinFirst
    (141, Operand::Fixed(2)),  // sprmSDmBinOther
    (142, Operand::Fixed(1)),  // sprmSBkc
    (143, Operand::Fixed(1)),  // sprmSFTitlePage
    (144, Operand::Fixed(2)),  // sprmSCcolumns
    (145, Operand::Fixed(2)),  // sprmSDxaColumns
    (146, Operand::Fixed(1)),  // sprmSFAutoPgn
    (147, Operand::Fixed(1)),  // sprmSNfcPgn
    (148, Operand::Fixed(2)),  // sprmSDyaPgn
    (149, Operand::Fixed(2)),  // sprmSDxaPgn
    (150, Operand::Fixed(1)),  // sprmSFPgnRestart
    (151, Operand::Fixed(1)),  // sprmSFEndnote
    (152, Operand::Fixed(1)),  // sprmSLnc
    (153, Operand::Fixed(1)),  // sprmSGprfIhdt
    (154, Operand::Fixed(2)),  // sprmSNLnnMod
    (155, Operand::Fixed(2)),  // sprmSDxaLnn
    (156, Operand::Fixed(2)),  // sprmSDyaHdrTop
    (157, Operand::Fixed(2)),  // sprmSDyaHdrBottom
    (158, Operand::Fixed(1)),  // sprmSLBetween
    (159, Operand::Fixed(1)),  // sprmSVjc
    (160, Operand::Fixed(2)),  // sprmSLnnMin
    (161, Operand::Fixed(2)),  // sprmSPgnStart
    (162, Operand::Fixed(1)),  // sprmSBOrientation
    (163, Operand::Fixed(0)),  // sprmSBCustomize
    (164, Operand::Fixed(2)),  // sprmSXaPage
    (165, Operand::Fixed(2)),  // sprmSYaPage
    (166, Operand::Fixed(2)),  // sprmSDxaLeft
    (167, Operand::Fixed(2)),  // sprmSDxaRight
    (168, Operand::Fixed(2)),  // sprmSDyaTop
    (169, Operand::Fixed(2)),  // sprmSDyaBottom
    (170, Operand::Fixed(2)),  // sprmSDzaGutter
    (171, Operand::Fixed(2)),  // sprmSDMPaperReq
    (179, Operand::Var),       // (unnamed)
    (181, Operand::Var),       // (unnamed)
    (182, Operand::Fixed(2)),  // sprmTJc
    (183, Operand::Fixed(2)),  // sprmTDxaLeft
    (184, Operand::Fixed(2)),  // sprmTDxaGapHalf
    (185, Operand::Fixed(1)),  // sprmTFCantSplit
    (186, Operand::Fixed(1)),  // sprmTTableHeader
    (187, Operand::Fixed(12)), // sprmTTableBorders
    (188, Operand::Var2),      // sprmTDefTable10
    (189, Operand::Fixed(2)),  // sprmTDyaRowHeight
    (190, Operand::Var2),      // sprmTDefTable
    (191, Operand::Var),       // sprmTDefTableShd
    (192, Operand::Fixed(4)),  // sprmTTlp
    (193, Operand::Fixed(5)),  // sprmTSetBrc
    (194, Operand::Fixed(4)),  // sprmTInsert
    (195, Operand::Fixed(2)),  // sprmTDelete
    (196, Operand::Fixed(4)),  // sprmTDxaCol
    (197, Operand::Fixed(2)),  // sprmTMerge
    (198, Operand::Fixed(2)),  // sprmTSplit
    (199, Operand::Fixed(5)),  // sprmTSetBrc10
    (200, Operand::Fixed(4)),  // sprmTSetShd
    (207, Operand::Var),       // (unnamed)
];

fn operand_of(sprm: u8) -> Option<Operand> {
    SPRMS
        .binary_search_by_key(&sprm, |&(s, _)| s)
        .ok()
        .map(|i| SPRMS[i].1)
}

/// Bytes of translated property lists allowed on top of four per byte of the
/// file. Re-encoding a row definition grows it (zeroed cell descriptors), and
/// any number of runs may name the same PAPX, so the total is bounded.
const TRANSLATION_ALLOWANCE: usize = 64 << 10;

/// The most columns a row definition may claim before it is not trusted.
const MAX_COLUMNS: usize = 63;

/// The Word 97 sprm bytes for the table properties found in a Word 6 property
/// list, and whether the whole list could be walked.
struct Translated {
    /// `sprmPFInTable`, `sprmTTableHeader` and, for a row end, `sprmTDefTable`
    /// in the Word 97 encoding.
    grpprl: Vec<u8>,
    /// `false` when an unknown sprm or a damaged length stopped the walk.
    complete: bool,
    /// A row definition was present but unusable (too many columns, or shorter
    /// than its own column count says).
    lost_row_definition: bool,
}

fn translate(list: &[u8]) -> Translated {
    let mut out = Vec::new();
    let mut p = 0usize;
    let mut complete = true;
    let mut ttp = false;
    let mut lost_row_definition = false;
    let mut def_table: Option<&[u8]> = None;
    while p < list.len() {
        let sprm = list[p];
        p += 1;
        let Some(kind) = operand_of(sprm) else {
            complete = false;
            break;
        };
        let (start, len) = match kind {
            Operand::Fixed(n) => (p, usize::from(n)),
            Operand::Var => {
                let Some(&n) = list.get(p) else {
                    complete = false;
                    break;
                };
                if n == 255 && (sprm == SPRM_P_CHG_TABS || sprm == SPRM_P_CHG_TABS_PAPX) {
                    // Length derived from the tab list's own counts; not stepped.
                    complete = false;
                    break;
                }
                (p + 1, usize::from(n))
            },
            Operand::Var2 => {
                let Some(b) = list.get(p..p + 2) else {
                    complete = false;
                    break;
                };
                (p + 2, usize::from(u16::from_le_bytes([b[0], b[1]])).saturating_sub(1))
            },
        };
        let Some(operand) = list.get(start..start.saturating_add(len)) else {
            complete = false;
            break;
        };
        p = start + len;
        match sprm {
            SPRM_P_F_IN_TABLE => {
                out.extend_from_slice(&SPRM97_P_F_IN_TABLE);
                out.push(operand.first().copied().unwrap_or(0));
            },
            SPRM_T_TABLE_HEADER => {
                out.extend_from_slice(&SPRM97_T_TABLE_HEADER);
                out.push(operand.first().copied().unwrap_or(0));
            },
            SPRM_P_F_TTP => ttp = ttp || operand.first().is_some_and(|b| b & 1 != 0),
            SPRM_T_DEF_TABLE => def_table = Some(operand),
            _ => {},
        }
    }
    if ttp || def_table.is_some() {
        // The Word 97 row mark is a `sprmTDefTable`. Only the column count and
        // boundaries are used; Word 6's `TC`s (10 bytes: merge flags and borders)
        // are not the Word 97 ones (20 bytes), so zeroed Word 97 ones are written.
        let operand = match def_table.map(word97_row_definition) {
            Some(Some(operand)) => operand,
            Some(None) => {
                // A row definition that cannot be used: the row still ends,
                // but without its column structure.
                lost_row_definition = true;
                Vec::new()
            },
            None => Vec::new(),
        };
        out.extend_from_slice(&SPRM97_T_DEF_TABLE);
        out.extend_from_slice(&(operand.len() as u16).saturating_add(1).to_le_bytes());
        out.extend_from_slice(&operand);
    }
    Translated {
        grpprl: out,
        complete,
        lost_row_definition,
    }
}

/// A Word 6 `sprmTDefTable` operand (`itcMac`, then `itcMac + 1` column
/// boundaries) as the Word 97 operand: the same, plus 20 zero bytes per cell.
/// `None` when the operand is shorter than its own `itcMac` says or claims more
/// columns than a table has.
fn word97_row_definition(operand: &[u8]) -> Option<Vec<u8>> {
    let columns = usize::from(*operand.first()?);
    let boundaries_end = 1 + 2 * (columns + 1);
    if columns == 0 || columns > MAX_COLUMNS || operand.len() < boundaries_end {
        return None;
    }
    let mut out = operand[..boundaries_end].to_vec();
    out.resize(boundaries_end + 20 * columns, 0);
    Some(out)
}

/// What reading a Word 6 file's PAPX found.
pub(super) struct PapxScan {
    /// One run per paragraph, with the table properties in Word 97 encoding.
    pub runs: Vec<FkpParagraph>,
    /// Whether any paragraph is in a table. Only then is the file read
    /// structurally (see `from_word6`); a file without one keeps the text-only
    /// view it always had.
    pub has_table: bool,
    /// Why the structure is incomplete, for `Metadata::warnings`.
    pub warnings: Vec<String>,
}

/// Clip the runs to a disjoint set, earliest first, dropping any that an
/// earlier one covers entirely.
///
/// Runs of a well-formed file never overlap. Overlapping ones are meaningless,
/// and the paragraph builder intersects every run with every piece, so a few
/// runs that each span the whole text would multiply against a large piece
/// table.
fn make_disjoint(runs: &mut Vec<FkpParagraph>) {
    runs.sort_by_key(|r| (r.fc_start, r.fc_end));
    let mut at = 0u32;
    runs.retain_mut(|r| {
        r.fc_start = r.fc_start.max(at);
        if r.fc_start >= r.fc_end {
            return false;
        }
        at = r.fc_end;
        true
    });
}

/// Give every range of the stream that no run covers a run with no properties.
///
/// Text outside every PAPX run has the default paragraph properties; a page
/// can stop short of the text its bin-table entry is indexed for. The paragraph
/// builder only emits text that some run covers, so without these the tail of
/// such a file would vanish from the structured view.
fn fill_uncovered(runs: &mut Vec<FkpParagraph>, stream_len: u32) {
    if runs.is_empty() {
        return;
    }
    let mut spans: Vec<(u32, u32)> = runs
        .iter()
        .map(|r| (r.fc_start, r.fc_end))
        .filter(|(a, b)| b > a)
        .collect();
    spans.sort_unstable();
    let mut gaps = Vec::new();
    let mut at = 0u32;
    for (start, end) in spans {
        if start > at {
            gaps.push((at, start.min(stream_len)));
        }
        at = at.max(end);
    }
    if at < stream_len {
        gaps.push((at, stream_len));
    }
    runs.extend(
        gaps.into_iter()
            .filter(|(a, b)| b > a)
            .map(|(fc_start, fc_end)| FkpParagraph {
                fc_start,
                fc_end,
                grpprl: Vec::new(),
                istd: 0,
            }),
    );
}

/// Read the paragraph properties of a Word 6.0/95 file.
pub(super) fn read_papx(word_doc: &[u8], fib: &Word6Fib) -> PapxScan {
    let mut scan = PapxScan {
        runs: Vec::new(),
        has_table: false,
        warnings: Vec::new(),
    };
    if fib.lcb_plcf_bte_papx == 0 {
        return scan;
    }
    let start = fib.fc_plcf_bte_papx as usize;
    let Some(plc) = start
        .checked_add(fib.lcb_plcf_bte_papx as usize)
        .and_then(|end| word_doc.get(start..end))
        .filter(|plc| plc.len() >= 4)
    else {
        scan.warnings.push(
            "the paragraph-properties (PAPX) table of this Word 6/95 file lies outside the file; \
             tables and paragraph structure are not available"
                .into(),
        );
        return scan;
    };
    // `n + 1` file positions of 4 bytes, then `n` page numbers of 2 bytes.
    let n = (plc.len() - 4) / 6;
    let pages = &plc[4 * (n + 1)..4 * (n + 1) + 2 * n];
    // At most one distinct page per 512 bytes of the file; a page named twice
    // is read once, so a hostile table cannot make the work grow.
    let mut seen = std::collections::HashSet::new();
    let max_pages = word_doc.len() / 512;
    // The bin table can list fewer pages than `cpnBtePap` says there are; the
    // rest follow the first page in order.
    let listed = pages.len() / 2;
    let missing = usize::from(fib.cpn_bte_pap)
        .saturating_sub(listed)
        .min(max_pages);
    let completed: Vec<[u8; 2]> = (listed..listed + missing)
        .filter_map(|i| u16::try_from(usize::from(fib.pn_pap_first) + i).ok())
        .map(u16::to_le_bytes)
        .collect();
    let mut unreadable_pages = 0usize;
    let mut partial_lists = 0usize;
    let mut lost_row_definitions = 0usize;
    let mut over_budget = 0usize;
    let mut budget = word_doc
        .len()
        .saturating_mul(4)
        .saturating_add(TRANSLATION_ALLOWANCE);
    for pn in pages
        .as_chunks::<2>()
        .0
        .iter()
        .take(max_pages)
        .chain(completed.iter())
    {
        let pn = usize::from(u16::from_le_bytes(*pn));
        if !seen.insert(pn) {
            continue;
        }
        let Some(page) = word_doc.get(pn * 512..pn * 512 + 512) else {
            unreadable_pages += 1;
            continue;
        };
        // `crun` runs need 4 * (crun + 1) + 7 * crun bytes below byte 511.
        let crun = usize::from(page[511]).min((511 - 4) / 11);
        let fc_at = |k: usize| {
            u32::from_le_bytes([
                page[4 * k],
                page[4 * k + 1],
                page[4 * k + 2],
                page[4 * k + 3],
            ])
        };
        for k in 0..crun {
            let offset = usize::from(page[4 * (crun + 1) + 7 * k]) * 2;
            let (istd, grpprl) = match page.get(offset..) {
                Some(rest) if offset != 0 && offset < 511 => {
                    let cb = usize::from(rest[0]) * 2;
                    let body = &rest[1..(1 + cb).min(rest.len() - 1)];
                    if body.len() >= 2 {
                        let t = translate(&body[2..]);
                        partial_lists += usize::from(!t.complete);
                        lost_row_definitions += usize::from(t.lost_row_definition);
                        let istd = u16::from_le_bytes([body[0], body[1]]);
                        // Many runs may name one PAPX, and a row definition
                        // grows when re-encoded, so the lists are budgeted
                        // against the size of the file.
                        if t.grpprl.len() > budget {
                            over_budget += 1;
                            (istd, Vec::new())
                        } else {
                            budget -= t.grpprl.len();
                            (istd, t.grpprl)
                        }
                    } else {
                        (0, Vec::new())
                    }
                },
                _ => (0, Vec::new()),
            };
            scan.runs.push(FkpParagraph {
                fc_start: fc_at(k),
                fc_end: fc_at(k + 1),
                grpprl,
                istd,
            });
        }
    }
    make_disjoint(&mut scan.runs);
    scan.has_table = scan
        .runs
        .iter()
        .any(|r| r.grpprl.starts_with(&[0x16, 0x24, 1]));
    fill_uncovered(&mut scan.runs, word_doc.len() as u32);
    if unreadable_pages > 0 {
        scan.warnings.push(format!(
            "{unreadable_pages} paragraph-properties (PAPX) page(s) of this Word 6/95 file lie \
             outside the file; the paragraphs they describe have no structure"
        ));
    }
    if partial_lists > 0 {
        scan.warnings.push(format!(
            "{partial_lists} paragraph property list(s) of this Word 6/95 file could not be read \
             to their end; table structure may be incomplete"
        ));
    }
    if over_budget > 0 {
        scan.warnings.push(format!(
            "{over_budget} paragraph property list(s) of this Word 6/95 file exceeded the \
             translation budget and were not read; table structure may be incomplete"
        ));
    }
    if lost_row_definitions > 0 {
        scan.warnings.push(format!(
            "{lost_row_definitions} table row definition(s) of this Word 6/95 file could not be \
             used; those rows have no column structure"
        ));
    }
    scan
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_sprm_table_is_sorted_and_unique() {
        assert!(SPRMS.windows(2).all(|w| w[0].0 < w[1].0), "binary search needs a sorted table");
    }

    #[test]
    fn test_the_four_table_sprms_have_the_sizes_the_walk_depends_on() {
        assert!(matches!(operand_of(SPRM_P_F_IN_TABLE), Some(Operand::Fixed(1))));
        assert!(matches!(operand_of(SPRM_P_F_TTP), Some(Operand::Fixed(1))));
        assert!(matches!(operand_of(SPRM_T_TABLE_HEADER), Some(Operand::Fixed(1))));
        assert!(matches!(operand_of(SPRM_T_DEF_TABLE), Some(Operand::Var2)));
    }

    #[test]
    fn test_translate_row_end_gives_the_word97_row_definition_with_zeroed_cells() {
        // 24=1, 25=1, 190 (cb 8: itcMac 2, three boundaries).
        let list = [24, 1, 25, 1, 190, 8, 0, 2, 0, 0, 0xE8, 3, 0xD0, 7];
        let t = translate(&list);
        assert!(t.complete);
        let p97 = super::super::sprm::extract_pap_props(&t.grpprl);
        assert!(p97.f_in_table && p97.is_table_trailing_mark);
        let tap = p97.tap.expect("row definition");
        assert_eq!((tap.itc_mac, tap.centers.as_slice()), (2, &[0, 1000, 2000][..]));
        assert_eq!(tap.cells.len(), 2);
    }

    #[test]
    fn test_translate_ttp_without_a_row_definition_still_marks_the_row_end() {
        let t = translate(&[24, 1, 25, 1]);
        let p97 = super::super::sprm::extract_pap_props(&t.grpprl);
        assert!(p97.is_table_trailing_mark);
        assert!(p97.tap.is_none());
    }

    #[test]
    fn test_translate_stops_at_an_unknown_sprm_and_says_so() {
        let t = translate(&[24, 1, 250, 9]);
        assert!(!t.complete);
        assert!(super::super::sprm::extract_pap_props(&t.grpprl).f_in_table);
    }

    #[test]
    fn test_translate_never_reads_past_a_truncated_operand() {
        for list in [
            &[190u8][..],
            &[190, 9][..],
            &[190, 9, 0, 2][..],
            &[24][..],
            &[23, 200, 1][..],
        ] {
            let _ = translate(list);
        }
    }

    #[test]
    fn test_row_definition_with_too_many_or_too_few_columns_is_refused() {
        assert!(word97_row_definition(&[0]).is_none());
        assert!(word97_row_definition(&[2, 0, 0]).is_none(), "shorter than itcMac says");
        let mut big = vec![(MAX_COLUMNS + 1) as u8];
        big.resize(1 + 2 * (MAX_COLUMNS + 2), 0);
        assert!(word97_row_definition(&big).is_none());
    }

    // ── hostile streams, built directly ──

    /// A Word 6 FIB with only the PAPX bin table filled in.
    fn fib_for(plc_at: usize, plc_len: usize) -> Word6Fib {
        Word6Fib {
            lid: 0x0409,
            complex: false,
            ext_char: false,
            fc_min: 0x300,
            fc_mac: 0x300,
            ccp: [0; 8],
            fc_clx: 0,
            lcb_clx: 0,
            fc_plcf_bte_papx: plc_at as u32,
            lcb_plcf_bte_papx: plc_len as u32,
            pn_pap_first: 0,
            cpn_bte_pap: 0,
            fc_plcf_sed: 0,
            lcb_plcf_sed: 0,
        }
    }

    /// An FKP page of runs `(fc start, fc end, property list)`; the entries
    /// that share a list share one PAPX, as the 7-byte entries allow.
    fn page(runs: &[(u32, u32, Vec<u8>)]) -> Vec<u8> {
        let crun = runs.len();
        let mut page = vec![0u8; 512];
        let mut papx_at = 4 * (crun + 1) + 7 * crun;
        papx_at += papx_at % 2;
        let mut placed: Vec<(Vec<u8>, usize)> = Vec::new();
        for (i, (start, _, grpprl)) in runs.iter().enumerate() {
            page[4 * i..4 * i + 4].copy_from_slice(&start.to_le_bytes());
            let at = match placed.iter().find(|(g, _)| g == grpprl) {
                Some((_, at)) => *at,
                None => {
                    let mut body = vec![0u8, 0u8];
                    body.extend_from_slice(grpprl);
                    if body.len() % 2 == 1 {
                        body.push(0);
                    }
                    let at = papx_at;
                    page[at] = (body.len() / 2) as u8;
                    page[at + 1..at + 1 + body.len()].copy_from_slice(&body);
                    papx_at += 1 + body.len();
                    papx_at += papx_at % 2;
                    placed.push((grpprl.clone(), at));
                    at
                },
            };
            page[4 * (crun + 1) + 7 * i] = (at / 2) as u8;
        }
        page[4 * crun..4 * crun + 4].copy_from_slice(&runs[crun - 1].1.to_le_bytes());
        page[511] = crun as u8;
        page
    }

    /// A WordDocument stream holding `pages`, all named by the bin table.
    fn stream_with(pages: &[Vec<u8>]) -> (Vec<u8>, Word6Fib) {
        let mut wd = vec![0u8; 0x400];
        wd.resize(wd.len().div_ceil(512) * 512, 0);
        let first = wd.len() / 512;
        for p in pages {
            wd.extend_from_slice(p);
        }
        let plc_at = wd.len();
        for i in 0..=pages.len() {
            wd.extend_from_slice(&(i as u32).to_le_bytes());
        }
        for i in 0..pages.len() {
            wd.extend_from_slice(&((first + i) as u16).to_le_bytes());
        }
        let plc_len = 4 * (pages.len() + 1) + 2 * pages.len();
        (wd, fib_for(plc_at, plc_len))
    }

    #[test]
    fn test_translated_property_lists_are_bounded_by_the_size_of_the_file() {
        // 33 runs on a page share one 136-byte PAPX that translates to about
        // 1.4 KB (a 63-column row definition with zeroed Word 97 cell
        // descriptors). Pages are cheap to name, so without a budget the output
        // is runs x 1.4 KB for a few KB of input.
        let mut list = vec![24, 1, 25, 1, 190];
        let mut operand = vec![63u8];
        operand.extend(std::iter::repeat_n(0u8, 2 * 64));
        list.extend_from_slice(&((operand.len() + 1) as u16).to_le_bytes());
        list.extend_from_slice(&operand);
        let runs: Vec<(u32, u32, Vec<u8>)> = (0..33)
            .map(|i| (i * 10, i * 10 + 10, list.clone()))
            .collect();
        let pages: Vec<Vec<u8>> = (0..60)
            .map(|k| {
                let shifted: Vec<_> = runs
                    .iter()
                    .map(|(a, b, g)| (a + k * 400, b + k * 400, g.clone()))
                    .collect();
                page(&shifted)
            })
            .collect();
        let (wd, fib) = stream_with(&pages);
        let scan = read_papx(&wd, &fib);
        let translated: usize = scan.runs.iter().map(|r| r.grpprl.len()).sum();
        assert!(
            translated <= 4 * wd.len() + (64 << 10),
            "{translated} bytes of translated lists from a {} byte stream",
            wd.len()
        );
        assert!(scan.warnings.iter().any(|w| w.contains("budget")), "{:?}", scan.warnings);
    }

    #[test]
    fn test_overlapping_runs_are_made_disjoint_before_they_reach_the_paragraph_builder() {
        // Runs that each span the whole text multiply against every piece in
        // the builder (runs x pieces). A run is clipped to what the ones before
        // it left, and one wholly covered is dropped.
        let wide = |g: Vec<u8>| (0u32, 0x1000_0000u32, g);
        let pages = vec![
            page(&[wide(vec![24, 1]), wide(vec![24, 1])]),
            page(&[wide(vec![])]),
        ];
        let (wd, fib) = stream_with(&pages);
        let scan = read_papx(&wd, &fib);
        let mut spans: Vec<(u32, u32)> = scan
            .runs
            .iter()
            .map(|r| (r.fc_start, r.fc_end))
            .filter(|(a, b)| b > a)
            .collect();
        spans.sort_unstable();
        assert!(spans.windows(2).all(|w| w[0].1 <= w[1].0), "runs still overlap: {spans:?}");
    }
}
