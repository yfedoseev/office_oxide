//! Minimal synthetic `.doc` (CFB/OLE2) writer for tests.
//!
//! Builds a valid legacy-Word binary document **entirely in code** so the
//! integration tests need no third-party fixture blobs (AGENTS.md rule #4 —
//! prefer code-internal minimal synthetic documents over committed
//! third-party/real files). The writer covers only what `DocDocument`
//! actually parses:
//!
//! - a FIB (File Information Block) in the `WordDocument` stream,
//! - a CLX piece table in the `0Table` stream,
//! - a `PlcfBtePapx` indexing one PAPX FKP page per paragraph,
//! - the main text (UTF-16LE) in the `WordDocument` stream.
//!
//! It is deliberately minimal and self-contained: it produces a single
//! Unicode piece, one FKP page per paragraph, and a single CFB storage
//! holding `WordDocument` + `0Table` (no mini-stream, no DIFAT chain).

use office_oxide::cfb::{CFB_SIGNATURE, END_OF_CHAIN, FAT_SECT, FREE_SECT};
use office_oxide::{Document, DocumentFormat};
use std::io::Cursor;

pub mod ppt;

/// Sentinel directory id: no sibling / child.
const NO_ENTRY: u32 = 0xFFFF_FFFF;

/// 512-byte pages reserved for the FIB at the start of `WordDocument`: its
/// `FibRgFcLcb97` runs past the first page (e.g. `fcPlcftxbxTxt` at
/// 0x25A), so the FKP pages and text start after two.
const FIB_PAGES: usize = 2;

/// A single main-text paragraph to encode.
pub struct Para {
    /// Paragraph text (assumed BMP / ASCII — encoded as UTF-16LE).
    pub text: &'static str,
    /// Terminator: `'\r'` (paragraph) or `'\u{7}'` (cell / row mark).
    pub terminator: char,
    /// PAPX `grpprl` bytes — the part *after* the `cw`/`istd` header that
    /// [`crate`] decoding walks via `extract_pap_props`.
    pub grpprl: Vec<u8>,
}

/// Empty prose paragraph grpprl.
#[allow(dead_code)]
pub fn prose_grpprl() -> Vec<u8> {
    Vec::new()
}

/// List-item grpprl: `sprmPIlvl` (0x260A, 1-byte operand) for the level plus
/// `sprmPIlfo` (0x460B, 2-byte operand) for the list format override id. Per
/// [MS-DOC] §2.4.6.3 a paragraph is a list item only when its `ilfo` is a valid
/// index, so both SPRMs are emitted; emitting only the level SPRM would make
/// the walker (correctly) treat the paragraph as prose.
#[allow(dead_code)]
pub fn list_grpprl(ilvl: u8) -> Vec<u8> {
    let mut g = vec![0x0A, 0x26, ilvl]; // sprmPIlvl
    g.extend_from_slice(&[0x0B, 0x46, 0x01, 0x00]); // sprmPIlfo = 1
    g
}

/// Cell paragraph grpprl: `sprmPFInTable` (0x2416)=1, `sprmPItap` (0x6649)=1.
#[allow(dead_code)]
pub fn cell_grpprl() -> Vec<u8> {
    cell_grpprl_at(1)
}

/// Cell paragraph grpprl at table nesting depth `itap` (`sprmPItap`, 1 = outermost).
#[allow(dead_code)]
pub fn cell_grpprl_at(itap: u8) -> Vec<u8> {
    vec![0x16, 0x24, 0x01, 0x49, 0x66, itap, 0x00, 0x00, 0x00]
}

/// Row-terminator grpprl: `fInTable` + `fInTableTtp` + `itap`, then
/// `sprmTDefTable` (0xD608) carrying the column boundaries and per-cell
/// `rgf` flags (the TAP). `centers` has `rgfs.len() + 1` entries (column
/// edges in twips); `rgfs[i]` is the `rgf` flags word of cell `i` (MS-DOC
/// `TCGRF` 2-bit `fVertMerge` field: `fvmMerge` = 0x0020, `fvmRestart` = 0x0060).
#[allow(dead_code)]
pub fn row_grpprl(centers: &[i16], rgfs: &[u16]) -> Vec<u8> {
    row_grpprl_at(1, centers, rgfs)
}

/// [`row_grpprl`] for a row at table nesting depth `itap`.
#[allow(dead_code)]
pub fn row_grpprl_at(itap: u8, centers: &[i16], rgfs: &[u16]) -> Vec<u8> {
    assert_eq!(centers.len(), rgfs.len() + 1, "centers = cells + 1");
    let itc_mac = (centers.len() - 1) as u8;
    let mut operand = Vec::new();
    operand.push(itc_mac);
    for &c in centers {
        operand.extend_from_slice(&c.to_le_bytes());
    }
    for &rf in rgfs {
        // TKBKTAP descriptor: rgf (2 bytes) + w_width (2 bytes) + 16 padding.
        operand.extend_from_slice(&rf.to_le_bytes());
        operand.extend_from_slice(&0u16.to_le_bytes());
        operand.extend_from_slice(&[0u8; 16]);
    }
    let cb = (operand.len() + 1) as u16; // operand length is `cb - 1`
    let mut g = vec![0x08, 0xD6, cb as u8, (cb >> 8) as u8];
    g.extend_from_slice(&operand);

    // Head: sprmPFInTable = 1, sprmPFInTableTtp = 1, sprmPItap = 1.
    let mut head = vec![
        0x16, 0x24, 0x01, 0x17, 0x24, 0x01, 0x49, 0x66, itap, 0x00, 0x00, 0x00,
    ];
    head.extend_from_slice(&g);
    head
}

/// Last paragraph of a cell of a **nested** table (`itap` >= 2) as Word writes
/// it: an ordinary `\r`-terminated paragraph carrying `sprmPFInTable` (0x2416),
/// `sprmPItap` (0x6649) and `sprmPFInnerTableCell` (0x244B) -- not the `\x07`
/// end mark of an outermost cell.
#[allow(dead_code)]
pub fn inner_cell_grpprl(itap: u8) -> Vec<u8> {
    vec![
        0x16, 0x24, 0x01, 0x49, 0x66, itap, 0x00, 0x00, 0x00, 0x4B, 0x24, 0x01,
    ]
}

/// Row end of a **nested** table as Word writes it: a `\r` paragraph carrying
/// `sprmPFInTable`, `sprmPFInnerTtp` (0x244C, not the outermost `sprmPFTtp`),
/// `sprmPItap` and the row definition (`sprmTDefTable`).
#[allow(dead_code)]
pub fn inner_row_grpprl(itap: u8, centers: &[i16], rgfs: &[u16]) -> Vec<u8> {
    // Reuse the TAP bytes of the outermost row definition, then swap the head.
    let outer = row_grpprl_at(itap, centers, rgfs);
    let mut g = vec![
        0x16, 0x24, 0x01, 0x4C, 0x24, 0x01, 0x49, 0x66, itap, 0x00, 0x00, 0x00,
    ];
    g.extend_from_slice(&outer[12..]);
    g
}

/// Build a complete synthetic `.doc` from the given paragraphs.
#[allow(dead_code)]
pub fn build_doc(paras: &[Para]) -> Vec<u8> {
    build_doc_full(paras, &Subdocs::default(), FibTweaks::default())
}

/// Character counts for the subdocuments that follow the main text in the
/// piece table's character space. Each string is appended, in the order
/// [MS-DOC] fixes, and its length written to the matching `ccp*` field.
#[derive(Default)]
#[allow(dead_code)]
pub struct Subdocs {
    /// `ccpFtn` — footnote bodies.
    pub footnotes: &'static str,
    /// `ccpHdd` — header and footer bodies.
    pub headers: &'static str,
    /// `ccpAtn` — comment bodies.
    pub comments: &'static str,
    /// `ccpEdn` — endnote bodies.
    pub endnotes: &'static str,
    /// `ccpTxbx` — text-box bodies.
    pub textboxes: &'static str,
}

/// FIB fields a test wants to set to something other than the valid default.
#[derive(Default)]
#[allow(dead_code)]
pub struct FibTweaks {
    /// Override `wIdent` (0xA5DC selects the unsupported Word 6.0/95 layout).
    pub wident: Option<u16>,
    /// Set `fEncrypted` (flags bit 8).
    pub encrypted: bool,
    /// Override `fcClx`, to point the piece table outside the table stream.
    pub clx_offset: Option<u32>,
    /// Start the first paragraph's FKP run this many bytes *before* the
    /// text (Word leaves the first `rgfc` at the start of the text area,
    /// not at `fcMin`).
    pub first_fkp_fc_before_text: u32,
    /// `PlcfSed` section boundary CPs (`aCP`, one more than the section
    /// count; each `Sed` is written as zeros). Empty: no `PlcfSed`.
    pub plcf_sed: Vec<u32>,
    /// `PlcfHdd` header-story CPs (`aCP`, stories + 2 entries). Empty: no
    /// `PlcfHdd`.
    pub plcf_hdd: Vec<u32>,
    /// `PlcftxbxTxt` text-box story CPs (`aCP`, one more than the story
    /// count; each `FTXBXS` is written as zeros). Empty: no `PlcftxbxTxt`.
    pub plcf_txbx_txt: Vec<u32>,
    /// `FTXBXS.lid` of each `PlcftxbxTxt` story, in order (the shape id the
    /// box's text begins in). Stories past this list keep `lid` 0; the
    /// last `FTXBXS` is always written with `fReusable` = 1.
    pub txbx_lids: Vec<u32>,
    /// `PlcSpaMom` shape anchors as `(anchor CP, Spa.lid)`, followed by a
    /// final CP of `text_len`. Empty: no `PlcSpaMom`.
    pub plcf_spa: Vec<(u32, u32)>,
}

/// Build a synthetic `.doc`, optionally with subdocuments and FIB tweaks.
#[allow(dead_code)]
pub fn build_doc_full(paras: &[Para], subdocs: &Subdocs, tweaks: FibTweaks) -> Vec<u8> {
    let (word_doc, table) = build_doc_streams(paras, subdocs, tweaks);
    build_cfb(&word_doc, &table)
}

/// The `WordDocument` and `0Table` stream bytes of
/// [`build_doc_full`], for a test that assembles its own container.
#[allow(dead_code)]
pub fn build_doc_streams(
    paras: &[Para],
    subdocs: &Subdocs,
    tweaks: FibTweaks,
) -> (Vec<u8>, Vec<u8>) {
    build_streams(paras, subdocs, tweaks, &[], &[])
}

/// Build a synthetic `.doc` whose paragraphs carry the given PAPX `istd` values
/// and whose Table stream ends with `stsh` (an STSH style sheet).
///
/// Separate from [`build_doc_full`] so the existing tests keep building
/// style-free documents, and so an integration test can exercise the
/// style-sheet path through the public API alone — which is what the
/// revert-check gate needs, since it reverts everything under `src/`.
#[allow(dead_code)]
pub fn build_doc_with_styles(paras: &[Para], istds: &[u16], stsh: &[u8]) -> Vec<u8> {
    let (word_doc, table) =
        build_streams(paras, &Subdocs::default(), FibTweaks::default(), istds, stsh);
    build_cfb(&word_doc, &table)
}

fn build_streams(
    paras: &[Para],
    subdocs: &Subdocs,
    tweaks: FibTweaks,
    istds: &[u16],
    stsh: &[u8],
) -> (Vec<u8>, Vec<u8>) {
    let n = paras.len();

    // Build the main text (UTF-16LE) and the CP range of each paragraph.
    let mut units: Vec<u16> = Vec::new();
    let mut cp_starts: Vec<u32> = Vec::with_capacity(n);
    for p in paras {
        cp_starts.push(units.len() as u32);
        for ch in p.text.chars() {
            units.push(ch as u16);
        }
        units.push(p.terminator as u16);
    }
    let text_len = units.len() as u32;

    // Subdocument text follows the main text in the same character space.
    let subdoc_specs = [
        subdocs.footnotes,
        subdocs.headers,
        subdocs.comments,
        subdocs.endnotes,
        subdocs.textboxes,
    ];
    let mut ccps = [0u32; 5];
    for (i, text) in subdoc_specs.iter().enumerate() {
        if text.is_empty() {
            continue;
        }
        let start = units.len();
        for ch in text.chars() {
            units.push(ch as u16);
        }
        units.push('\r' as u16);
        ccps[i] = (units.len() - start) as u32;
    }
    let text_bytes: Vec<u8> = units.iter().flat_map(|u| u.to_le_bytes()).collect();
    let total_chars = units.len() as u32;
    // Text lives after the FIB pages (0..FIB_PAGES) and the N FKP pages.
    // Text lives after the FIB page (page 0) and the N FKP pages (pages 1..N).
    let text_offset = ((n + FIB_PAGES) as u32) * 512;

    // ── 0Table stream: CLX (piece table) followed by PlcfBtePapx. ──
    let mut table = build_clx(text_offset, total_chars);
    let fc_plcf = table.len() as u32;
    table.extend_from_slice(&build_plcf_bte_papx(n, &cp_starts, text_len));
    let lcb_plcf = (table.len() as u32) - fc_plcf;
    // Optional PLCs, each placed after what is already in the table stream:
    // (FIB fc offset, aCP, data-element size).
    let mut plcs: Vec<(usize, u32, u32)> = Vec::new();
    let spa_cps: Vec<u32> = if tweaks.plcf_spa.is_empty() {
        Vec::new()
    } else {
        let mut v: Vec<u32> = tweaks.plcf_spa.iter().map(|&(cp, _)| cp).collect();
        v.push(text_len);
        v
    };
    for (fib_off, cps, cb) in [
        (0x00CA, &tweaks.plcf_sed, 12usize), // fcPlcfSed: Sed = 12 bytes
        (0x00F2, &tweaks.plcf_hdd, 0),       // fcPlcfHdd: no data elements
        (0x025A, &tweaks.plcf_txbx_txt, 22), // fcPlcftxbxTxt: FTXBXS = 22 bytes
        (0x01DA, &spa_cps, 26),              // fcPlcSpaMom: Spa = 26 bytes
    ] {
        if cps.is_empty() {
            continue;
        }
        let fc = table.len() as u32;
        for cp in cps.iter() {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        let elems = if cb == 0 { 0 } else { cps.len() - 1 };
        for i in 0..elems {
            let mut elem = vec![0u8; cb];
            match fib_off {
                0x025A => {
                    // FTXBXS: fReusable at 8, lid at 14.
                    if i + 1 == elems {
                        elem[8..10].copy_from_slice(&1u16.to_le_bytes());
                    } else if let Some(lid) = tweaks.txbx_lids.get(i) {
                        elem[14..18].copy_from_slice(&lid.to_le_bytes());
                    }
                },
                // Spa: lid first.
                0x01DA => elem[0..4].copy_from_slice(&tweaks.plcf_spa[i].1.to_le_bytes()),
                _ => {},
            }
            table.extend_from_slice(&elem);
        }
        plcs.push((fib_off, fc, table.len() as u32 - fc));
    }

    // An optional style sheet appended after the PAPX index; `fcStshf` is an
    // offset into the Table stream, and 0 is only "absent" when `lcbStshf` is.
    let fc_stshf = table.len() as u32;
    let lcb_stshf = stsh.len() as u32;
    table.extend_from_slice(stsh);

    // ── WordDocument stream: FIB + N FKP pages + text. ──
    let wd_len = text_offset as usize + text_bytes.len();
    let wd_sectors = wd_len.div_ceil(512);
    let mut word_doc = vec![0u8; wd_sectors * 512];
    write_fib(&mut word_doc, text_len, fc_plcf, lcb_plcf, fc_stshf, lcb_stshf);
    // [MS-DOC] FibRgLw97 real offsets for these fields: 0x58
    // is `reserved3`, MUST be zero/ignored, and sits between ccpHdd and
    // ccpAtn — not a slot in this array, so the mapping isn't a flat
    // `0x50 + i*4` stride.
    const CCP_OFFSETS: [usize; 5] = [
        0x50, // ccpFtn (footnotes)
        0x54, // ccpHdd (headers)
        0x5C, // ccpAtn (comments) — NOT 0x58, which is reserved3
        0x60, // ccpEdn (endnotes)
        0x64, // ccpTxbx (textboxes)
    ];
    for (i, &ccp) in ccps.iter().enumerate() {
        let off = CCP_OFFSETS[i];
        word_doc[off..off + 4].copy_from_slice(&ccp.to_le_bytes());
    }
    if let Some(w) = tweaks.wident {
        word_doc[0..2].copy_from_slice(&w.to_le_bytes());
    }
    if tweaks.encrypted {
        word_doc[0x0A..0x0C].copy_from_slice(&(1u16 << 8).to_le_bytes());
    }
    if let Some(off) = tweaks.clx_offset {
        word_doc[0x01A2..0x01A6].copy_from_slice(&off.to_le_bytes());
    }
    for (fib_off, fc, lcb) in plcs {
        word_doc[fib_off..fib_off + 4].copy_from_slice(&fc.to_le_bytes());
        word_doc[fib_off + 4..fib_off + 8].copy_from_slice(&lcb.to_le_bytes());
    }
    for (i, p) in paras.iter().enumerate() {
        let cp0 = cp_starts[i];
        let cp1 = if i + 1 < n {
            cp_starts[i + 1]
        } else {
            text_len
        };
        let mut fc0 = text_offset + cp0 * 2;
        if i == 0 {
            fc0 -= tweaks.first_fkp_fc_before_text;
        }
        let fc1 = text_offset + cp1 * 2;
        let istd = istds.get(i).copied().unwrap_or(0);
        let page = build_fkp_page(fc0, fc1, &p.grpprl, istd);
        let off = (i + FIB_PAGES) * 512;
        word_doc[off..off + 512].copy_from_slice(&page);
    }
    word_doc[text_offset as usize..text_offset as usize + text_bytes.len()]
        .copy_from_slice(&text_bytes);

    (word_doc, table)
}

/// One paragraph style of a synthetic style sheet, in `istd` order.
///
/// Built from the spec ([MS-DOC] §2.9.258 `STD`), not from the crate's parser,
/// so a fixture cannot merely re-encode the parser's own bugs.
#[allow(dead_code)]
pub struct Style<'a> {
    /// Built-in style id (`StdfBase.sti`); `0x0FFE` = user-defined.
    pub sti: u16,
    /// `xstzName`.
    pub name: &'a str,
    /// `istdBase`: the style this one is based on (`None` = `0x0FFF`).
    pub base: Option<u16>,
    /// The `sprmPOutLvl` operand in the style's own `UpxPapx` (`0..=8` heading,
    /// `9` body text); `None` = the style states no level.
    pub outline: Option<u8>,
    /// `StdfBase.stk` ([MS-DOC] §2.9.253): 1 paragraph, 2 character, 3 table,
    /// 4 numbering.
    pub kind: u16,
}

#[allow(dead_code)]
impl<'a> Style<'a> {
    /// A style with no base and no level of its own.
    pub fn new(sti: u16, name: &'a str) -> Self {
        Style {
            sti,
            name,
            base: None,
            outline: None,
            kind: 1,
        }
    }
    /// A user-defined style based on `base`.
    pub fn based_on(name: &'a str, base: u16) -> Self {
        Style {
            sti: 0x0FFE,
            name,
            base: Some(base),
            outline: None,
            kind: 1,
        }
    }
    /// The same style with a different style kind (`stk`).
    pub fn kind(mut self, stk: u16) -> Self {
        self.kind = stk;
        self
    }
    /// States its own outline level (zero-based operand; 9 = body text).
    pub fn outline(mut self, operand: u8) -> Self {
        self.outline = Some(operand);
        self
    }
}

/// One `LPStd`: `cbStd(u16)` + `STD` (+ the pad that keeps the next entry on an
/// even offset; `cbStd` excludes it, §2.9.135).
///
/// `STD` = `StdfBase` (10 bytes: `sti`; `sgc` 1 = paragraph with `istdBase`
/// above it; `cupx`) + `xstzName` + an even-aligned `LPUpx` holding a
/// `UpxPapx` (`istd` + a `sprmPOutLvl` grpprl) when the style states a level.
fn lpstd(style: &Style) -> Vec<u8> {
    let mut std = vec![0u8; 10];
    std[0..2].copy_from_slice(&style.sti.to_le_bytes());
    let istd_base = style.base.unwrap_or(0x0FFF);
    std[2..4].copy_from_slice(&(style.kind | (istd_base << 4)).to_le_bytes());
    std[4..6].copy_from_slice(&u16::from(style.outline.is_some()).to_le_bytes());
    let units: Vec<u16> = style.name.encode_utf16().collect();
    std.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for u in &units {
        std.extend_from_slice(&u.to_le_bytes());
    }
    std.extend_from_slice(&0u16.to_le_bytes());
    if let Some(operand) = style.outline {
        if std.len() % 2 == 1 {
            std.push(0);
        }
        std.extend_from_slice(&5u16.to_le_bytes()); // cbUpx: istd(2) + grpprl(3)
        std.extend_from_slice(&0u16.to_le_bytes()); // UpxPapx.istd
        std.extend_from_slice(&[0x40, 0x26, operand]); // sprmPOutLvl
    }
    let mut out = (std.len() as u16).to_le_bytes().to_vec();
    out.extend_from_slice(&std);
    if std.len() % 2 == 1 {
        out.push(0);
    }
    out
}

/// A spec-conformant STSH ([MS-DOC] §2.7.1 / §2.9.274) holding `styles` in
/// `istd` order: `LPStshi` (`cbStshi` = 18) followed by one `LPStd` each.
#[allow(dead_code)]
pub fn style_sheet(styles: &[Style]) -> Vec<u8> {
    let mut d = 18u16.to_le_bytes().to_vec(); // cbStshi
    d.extend_from_slice(&(styles.len() as u16).to_le_bytes()); // cstd
    d.extend_from_slice(&0x000Au16.to_le_bytes()); // cbSTDBaseInFile
    d.extend_from_slice(&[0u8; 14]); // rest of the 18-byte Stshif
    for style in styles {
        d.extend_from_slice(&lpstd(style));
    }
    d
}

/// A spec-conformant STSH ([MS-DOC] §2.7.1 / §2.9.274) whose `istd` 1-9 are the
/// built-in `Heading 1`-`Heading 9` styles and `istd` 0 is `Normal`, so a
/// paragraph whose PAPX `istd` is 3 resolves to Heading 3.
#[allow(dead_code)]
pub fn heading_style_sheet() -> Vec<u8> {
    let mut styles = vec![Style::new(0, "Normal")];
    let names: Vec<String> = (1..=9).map(|lvl| format!("Heading {lvl}")).collect();
    for (lvl, name) in (1..=9u16).zip(&names) {
        styles.push(Style {
            sti: lvl,
            name,
            base: Some(0),
            outline: None,
            kind: 1,
        });
    }
    let mut d = style_sheet(&styles);
    // Five empty `LPStd` slots (`cbStd` = 0), counted in `cstd`.
    d[2..4].copy_from_slice(&15u16.to_le_bytes());
    for _ in 0..5 {
        d.extend_from_slice(&[0u8; 2]);
    }
    d
}

/// Open a synthetic `.doc` byte buffer through the public API.
#[allow(dead_code)]
pub fn open_doc(bytes: &[u8]) -> Document {
    Document::from_reader(Cursor::new(bytes.to_vec()), DocumentFormat::Doc)
        .expect("synthetic .doc must parse")
}

/// Build a synthetic Word 6.0/95 `.doc`: the Word 6 FIB layout (text from
/// `fcMin`, story lengths at 0x34, no table stream), `wident` for the
/// family variant, `text` as one-byte characters in a `\r`-delimited
/// story sequence.
#[allow(dead_code)]
pub fn build_word6_doc(wident: u16, text: &[u8]) -> Vec<u8> {
    build_word6_doc_with_tail(wident, text, &[])
}

/// As [`build_word6_doc`], with `tail` stored in the `WordDocument` stream
/// after the text (outside every story) — where Word 6.0/95 keeps
/// pictures.
#[allow(dead_code)]
pub fn build_word6_doc_with_tail(wident: u16, text: &[u8], tail: &[u8]) -> Vec<u8> {
    let fc_min = 0x300u32;
    let mut wd = vec![0u8; fc_min as usize];
    wd[0..2].copy_from_slice(&wident.to_le_bytes());
    wd[2..4].copy_from_slice(&101u16.to_le_bytes()); // nFib: Word 6.0
    wd[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
    wd[0x18..0x1C].copy_from_slice(&fc_min.to_le_bytes());
    wd[0x1C..0x20].copy_from_slice(&(fc_min + text.len() as u32).to_le_bytes());
    wd[0x34..0x38].copy_from_slice(&(text.len() as u32).to_le_bytes());
    wd.extend_from_slice(text);
    wd.extend_from_slice(tail);
    let pad = (512 - wd.len() % 512) % 512;
    wd.extend(std::iter::repeat_n(0u8, pad));
    build_cfb(&wd, &[0u8; 512])
}

/// Word 6.0/95 `sprmPFInTable` (sprm 24): the paragraph is in a table.
#[allow(dead_code)]
pub fn w6_in_table() -> Vec<u8> {
    vec![24, 1]
}

/// Word 6.0/95 row end as Word writes it: `sprmPFInTable` (24), `sprmPFTtp`
/// (25), the row's other table sprms (`sprmTDxaGapHalf` 184, `sprmTFCantSplit`
/// 185, `sprmTDyaRowHeight` 189 and the 12-byte `sprmTTableBorders` 187), then
/// `sprmTDefTable` (190, a 2-byte length counting itself) for `cols` equal
/// columns: `itcMac`, the column boundaries and one 10-byte `TC` per cell.
#[allow(dead_code)]
pub fn w6_row_end(cols: usize, header: bool) -> Vec<u8> {
    let mut g = vec![24, 1, 25, 1];
    if header {
        g.extend_from_slice(&[186, 1]); // sprmTTableHeader
    }
    g.extend_from_slice(&[184, 0x6C, 0x00]); // sprmTDxaGapHalf
    g.extend_from_slice(&[185, 1]); // sprmTFCantSplit
    g.push(187); // sprmTTableBorders
    g.extend_from_slice(&[0u8; 12]);
    g.extend_from_slice(&[189, 0x20, 0x00]); // sprmTDyaRowHeight
    let mut operand = vec![cols as u8];
    for c in 0..=cols {
        operand.extend_from_slice(&((c * 1000) as i16).to_le_bytes());
    }
    if cols <= 16 {
        // rgtc: one 10-byte TC per cell (a very wide row would not fit a page)
        operand.extend(std::iter::repeat_n(0u8, 10 * cols));
    }
    g.push(190); // sprmTDefTable
    g.extend_from_slice(&((operand.len() + 1) as u16).to_le_bytes());
    g.extend_from_slice(&operand);
    g
}

/// How a synthetic Word 6.0/95 `.doc` is laid out.
#[allow(dead_code)]
pub struct Word6Spec<'a> {
    /// `wIdent`: `0xA5DC`, or one of the three Macintosh variants.
    pub wident: u16,
    /// FKP pages; each page's paragraphs are `(text, Word 6 grpprl)` and the
    /// pages follow one another in the text.
    pub pages: Vec<Vec<(&'a [u8], Vec<u8>)>>,
    /// Text after the last page's paragraphs that no FKP run describes.
    pub tail: &'a [u8],
    /// How many pages the bin table lists. The FIB's `cpnBtePap` always counts
    /// them all, so listing fewer is an incomplete bin table.
    pub listed_pages: usize,
    /// Name each listed page twice.
    pub duplicate_listed: bool,
    /// `PlcfSed` section end CPs; empty writes no `PlcfSed`.
    pub section_ends: Vec<u32>,
    /// `fExtChar`: the text is two-byte code units.
    pub ext_char: bool,
    /// A fast-saved file: a CLX with one piece describes the text.
    pub complex: bool,
}

#[allow(dead_code)]
impl<'a> Word6Spec<'a> {
    pub fn new(pages: Vec<Vec<(&'a [u8], Vec<u8>)>>) -> Self {
        let listed_pages = pages.len();
        Self {
            wident: 0xA5DC,
            pages,
            tail: &[],
            listed_pages,
            duplicate_listed: false,
            section_ends: Vec::new(),
            ext_char: false,
            complex: false,
        }
    }

    /// The `.doc`: FIB, text, one 512-byte FKP page per spec page, then the
    /// `PlcfBtePapx` (FCs, then 2-byte page numbers), the `PlcfSed` and the CLX.
    pub fn build(&self) -> Vec<u8> {
        let unit = if self.ext_char { 2 } else { 1 };
        let widen = |t: &[u8]| -> Vec<u8> {
            if self.ext_char {
                t.iter().flat_map(|&b| [b, 0]).collect()
            } else {
                t.to_vec()
            }
        };
        let fc_min = 0x300u32;
        let mut text = Vec::new();
        for page in &self.pages {
            for (t, _) in page {
                text.extend(widen(t));
            }
        }
        text.extend(widen(self.tail));
        let ccp = (text.len() / unit) as u32;

        let mut wd = vec![0u8; fc_min as usize];
        wd[0..2].copy_from_slice(&self.wident.to_le_bytes());
        wd[2..4].copy_from_slice(&101u16.to_le_bytes());
        wd[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        let mut flags = 0u16;
        if self.complex {
            flags |= 0x0004;
        }
        if self.ext_char {
            flags |= 0x1000;
        }
        wd[0x0A..0x0C].copy_from_slice(&flags.to_le_bytes());
        wd[0x18..0x1C].copy_from_slice(&fc_min.to_le_bytes());
        wd[0x1C..0x20].copy_from_slice(&(fc_min + text.len() as u32).to_le_bytes());
        wd[0x34..0x38].copy_from_slice(&ccp.to_le_bytes());
        wd.extend_from_slice(&text);
        wd.resize(wd.len().div_ceil(512) * 512, 0);

        // The FKP pages.
        let first_pn = (wd.len() / 512) as u16;
        let mut fc = fc_min;
        for page_runs in &self.pages {
            let crun = page_runs.len();
            let mut page = vec![0u8; 512];
            let mut papx_at = 4 * (crun + 1) + 7 * crun;
            papx_at += papx_at % 2;
            for (i, (t, grpprl)) in page_runs.iter().enumerate() {
                page[4 * i..4 * i + 4].copy_from_slice(&fc.to_le_bytes());
                fc += (t.len() * unit) as u32;
                let mut body = vec![0u8, 0u8]; // istd
                body.extend_from_slice(grpprl);
                if body.len() % 2 == 1 {
                    body.push(0);
                }
                page[4 * (crun + 1) + 7 * i] = (papx_at / 2) as u8;
                page[papx_at] = (body.len() / 2) as u8;
                page[papx_at + 1..papx_at + 1 + body.len()].copy_from_slice(&body);
                papx_at += 1 + body.len();
                papx_at += papx_at % 2;
            }
            page[4 * crun..4 * crun + 4].copy_from_slice(&fc.to_le_bytes());
            page[511] = crun as u8;
            wd.extend_from_slice(&page);
        }

        // The PlcfBtePapx lists the first `listed_pages` pages.
        let listed: Vec<u16> = (0..self.listed_pages.min(self.pages.len()))
            .flat_map(|i| {
                let pn = first_pn + i as u16;
                if self.duplicate_listed {
                    vec![pn, pn]
                } else {
                    vec![pn]
                }
            })
            .collect();
        let plc_at = wd.len() as u32;
        for i in 0..=listed.len() {
            let fc_i = fc_min + (i as u32) * 8;
            wd.extend_from_slice(&fc_i.to_le_bytes());
        }
        for pn in &listed {
            wd.extend_from_slice(&pn.to_le_bytes());
        }
        let plc_len = (4 * (listed.len() + 1) + 2 * listed.len()) as u32;
        wd[0xC0..0xC4].copy_from_slice(&plc_at.to_le_bytes());
        wd[0xC4..0xC8].copy_from_slice(&plc_len.to_le_bytes());
        // pnPapFirst / cpnBtePap: where the pages are, and how many there are.
        wd[0x18C..0x18E].copy_from_slice(&(first_pn as i16).to_le_bytes());
        wd[0x190..0x192].copy_from_slice(&(self.pages.len() as i16).to_le_bytes());

        if !self.section_ends.is_empty() {
            let sed_at = wd.len() as u32;
            wd.extend_from_slice(&0u32.to_le_bytes());
            for cp in &self.section_ends {
                wd.extend_from_slice(&cp.to_le_bytes());
            }
            wd.extend(std::iter::repeat_n(0u8, 12 * self.section_ends.len())); // SEDs
            let len = (4 * (self.section_ends.len() + 1) + 12 * self.section_ends.len()) as u32;
            wd[0x88..0x8C].copy_from_slice(&sed_at.to_le_bytes());
            wd[0x8C..0x90].copy_from_slice(&len.to_le_bytes());
        }

        if self.complex {
            // A CLX with one 8-bit (or, with `fExtChar`, 16-bit) piece over all
            // the text: `fc` carries `byte offset * 2` with bit 30 set.
            let clx_at = wd.len() as u32;
            let mut clx = vec![0x02u8];
            clx.extend_from_slice(&16u32.to_le_bytes());
            clx.extend_from_slice(&0u32.to_le_bytes());
            clx.extend_from_slice(&ccp.to_le_bytes());
            clx.extend_from_slice(&0u16.to_le_bytes());
            let piece_fc = if self.ext_char {
                fc_min
            } else {
                (fc_min * 2) | 0x4000_0000
            };
            clx.extend_from_slice(&piece_fc.to_le_bytes());
            clx.extend_from_slice(&0u16.to_le_bytes());
            wd.extend_from_slice(&clx);
            wd[0x160..0x164].copy_from_slice(&clx_at.to_le_bytes());
            wd[0x164..0x168].copy_from_slice(&(clx.len() as u32).to_le_bytes());
        }
        wd.resize(wd.len().div_ceil(512) * 512, 0);
        build_cfb(&wd, &[0u8; 512])
    }
}

/// A Word 6.0/95 `.doc` whose paragraphs carry PAPX property lists: one FKP
/// page, every paragraph described.
#[allow(dead_code)]
pub fn build_word6_doc_with_papx(wident: u16, paras: &[(&[u8], Vec<u8>)]) -> Vec<u8> {
    let mut spec = Word6Spec::new(vec![paras.to_vec()]);
    spec.wident = wident;
    spec.build()
}

/// As [`build_word6_doc_with_papx`], but the FKP describes only the first
/// `covered` paragraphs: the rest of the text has no PAPX run, as in a real file
/// whose last page stops short of the text it is indexed for.
#[allow(dead_code)]
pub fn build_word6_doc_with_papx_covering(
    wident: u16,
    paras: &[(&[u8], Vec<u8>)],
    covered: usize,
) -> Vec<u8> {
    let covered = covered.min(paras.len());
    let tail: Vec<u8> = paras[covered..]
        .iter()
        .flat_map(|(t, _)| t.iter().copied())
        .collect();
    let mut spec = Word6Spec::new(vec![paras[..covered].to_vec()]);
    spec.wident = wident;
    spec.tail = &tail;
    spec.build()
}

/// A Word 6.0/95 document whose FIB points `PlcfBtePapx` far past the end of the
/// stream.
#[allow(dead_code)]
pub fn open_word6_with_bad_papx_pointer(wident: u16) -> Document {
    let fc_min = 0x300u32;
    let text = b"A\x07B\x07\x07";
    let mut wd = vec![0u8; fc_min as usize];
    wd[0..2].copy_from_slice(&wident.to_le_bytes());
    wd[2..4].copy_from_slice(&101u16.to_le_bytes());
    wd[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
    wd[0x18..0x1C].copy_from_slice(&fc_min.to_le_bytes());
    wd[0x1C..0x20].copy_from_slice(&(fc_min + text.len() as u32).to_le_bytes());
    wd[0x34..0x38].copy_from_slice(&(text.len() as u32).to_le_bytes());
    wd[0xC0..0xC4].copy_from_slice(&0x00F0_0000u32.to_le_bytes());
    wd[0xC4..0xC8].copy_from_slice(&16u32.to_le_bytes());
    wd.extend_from_slice(text);
    wd.resize(wd.len().div_ceil(512) * 512, 0);
    open_doc(&build_cfb(&wd, &[0u8; 512]))
}

// ── CFB / DOC byte construction ──

/// Write the FIB into the first 512 bytes of `word_doc`.
fn write_fib(
    word_doc: &mut [u8],
    text_len: u32,
    fc_plcf: u32,
    lcb_plcf: u32,
    fc_stshf: u32,
    lcb_stshf: u32,
) {
    word_doc[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent = Word 97+
    word_doc[2..4].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib
    // flags at 0x0A: bit 9 (fWhichTblStm) clear → use 0Table.
    word_doc[0x0A..0x0C].copy_from_slice(&0u16.to_le_bytes());
    word_doc[0x4C..0x50].copy_from_slice(&text_len.to_le_bytes()); // ccpText
    word_doc[0x00A2..0x00A6].copy_from_slice(&fc_stshf.to_le_bytes()); // fcStshf
    word_doc[0x00A6..0x00AA].copy_from_slice(&lcb_stshf.to_le_bytes()); // lcbStshf
    word_doc[0x01A2..0x01A6].copy_from_slice(&0u32.to_le_bytes()); // fcClx = 0
    word_doc[0x01A6..0x01AA].copy_from_slice(&21u32.to_le_bytes()); // lcbClx = CLX size
    word_doc[0x0102..0x0106].copy_from_slice(&fc_plcf.to_le_bytes());
    word_doc[0x0106..0x010A].copy_from_slice(&lcb_plcf.to_le_bytes());
}

/// Build the CLX: a single Pcdt piece table with one Unicode piece.
fn build_clx(text_offset: u32, text_len: u32) -> Vec<u8> {
    let mut clx = Vec::new();
    clx.push(0x02); // Pcdt marker
    clx.extend_from_slice(&16u32.to_le_bytes()); // PlcPcd size
    // PlcPcd: CP[0]=0, CP[1]=text_len, then one PCD (8 bytes).
    clx.extend_from_slice(&0u32.to_le_bytes());
    clx.extend_from_slice(&text_len.to_le_bytes());
    clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: unused u16
    clx.extend_from_slice(&text_offset.to_le_bytes()); // fc (byte offset in WordDocument)
    clx.extend_from_slice(&0u16.to_le_bytes()); // PCD: prm u16
    clx
}

/// Build the PlcfBtePapx: `(n+1)` u32 CP boundaries + `n` u32 BTEs, one FKP
/// page (pn = i+1) per paragraph.
fn build_plcf_bte_papx(n: usize, cp_starts: &[u32], text_len: u32) -> Vec<u8> {
    let mut v = Vec::new();
    for i in 0..=n {
        let cp = if i < n { cp_starts[i] } else { text_len };
        v.extend_from_slice(&cp.to_le_bytes());
    }
    for i in 0..n {
        v.extend_from_slice(&((i as u32) + FIB_PAGES as u32).to_le_bytes()); // page number
    }
    v
}

/// Build one 512-byte PAPX FKP page holding a single paragraph.
fn build_fkp_page(fc0: u32, fc1: u32, grpprl: &[u8], istd: u16) -> [u8; 512] {
    let mut page = [0u8; 512];
    page[0..4].copy_from_slice(&fc0.to_le_bytes());
    page[4..8].copy_from_slice(&fc1.to_le_bytes());

    // Word writes a paragraph whose formatting comes entirely from its style as
    // a *istd-only* PAPX: `cw = 0` triggers the Word8 re-read, whose follow
    // byte `cb' = 1` means a 2-byte `GrpPrlAndIstd` holding just the `istd`
    // and no grpprl. This is the shape the production parser must not drop, so
    // we emit it here whenever the grpprl is empty — the default fixture used
    // to test a `cb != 0` form that real files never use for style-only text.
    let papx = if grpprl.is_empty() {
        let mut p = vec![0u8, 1, 0, 0]; // cw = 0, cb' = 1, istd (2 bytes)
        p[2..4].copy_from_slice(&istd.to_le_bytes());
        p
    } else {
        // PAPX header = [cw][istd:2][grpprl]. Pad grpprl to an odd length so
        // that `3 + grpprl.len()` is even and `cw = (3 + grpprl.len()) / 2`.
        let mut g = grpprl.to_vec();
        if g.len().is_multiple_of(2) {
            g.push(0);
        }
        let cw = ((3 + g.len()) / 2) as u8;
        let mut p = vec![cw, 0, 0];
        p[1..3].copy_from_slice(&istd.to_le_bytes());
        p.extend_from_slice(&g);
        p
    };

    // PAPX stored at word offset 11 (byte 22); BX byte 0 carries the word off.
    let word_off = 11u8;
    page[8] = word_off;
    let start = word_off as usize * 2;
    page[start..start + papx.len()].copy_from_slice(&papx);

    page[511] = 1; // crun
    page
}

/// Assemble the CFB container around the `WordDocument` and `0Table` streams.
fn build_cfb(word_doc: &[u8], table: &[u8]) -> Vec<u8> {
    let wd_sectors = word_doc.len() / 512;
    let zero_table_sectors = table.len().div_ceil(512);
    let total_sectors = 2 + wd_sectors + zero_table_sectors; // dir + FAT + streams
    let mut file = vec![0u8; 512 + total_sectors * 512];

    write_header(&mut file);

    // Directory (sector 0) at offset 512.
    let dir_off = 512;
    // Root Entry's `child` points to entry 1 (WordDocument), which links
    // to entry 2 (0Table) as its right sibling — a minimal but real tree,
    // not just 3 unlinked entries.
    write_dir_entry(&mut file[dir_off..dir_off + 128], "Root Entry", 5, 1, END_OF_CHAIN, 0);
    let wd_start = 2u32;
    let zt_start = (2 + wd_sectors) as u32;
    write_dir_entry_with_sibling(
        &mut file[dir_off + 128..dir_off + 256],
        "WordDocument",
        2,
        NO_ENTRY,
        2,
        wd_start,
        word_doc.len() as u32,
    );
    write_dir_entry(
        &mut file[dir_off + 256..dir_off + 384],
        "0Table",
        2,
        NO_ENTRY,
        zt_start,
        table.len() as u32,
    );
    // Entry 3 is left as the zero-filled Empty slot.

    // FAT (sector 1) at offset 1024.
    write_fat(&mut file, wd_sectors, zero_table_sectors);

    // Stream data after header + dir + FAT.
    let wd_off = 512 + 2 * 512;
    file[wd_off..wd_off + word_doc.len()].copy_from_slice(word_doc);
    let zt_off = wd_off + wd_sectors * 512;
    file[zt_off..zt_off + table.len()].copy_from_slice(table);

    file
}

fn write_header(file: &mut [u8]) {
    file[0..8].copy_from_slice(&CFB_SIGNATURE);
    file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes()); // minor version
    file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes()); // major version = 3
    file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes()); // byte order
    file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes()); // sector size = 512
    file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes()); // mini sector = 64
    file[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes()); // 1 FAT sector
    file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes()); // first dir sector = 0
    file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes()); // mini-stream cutoff
    file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes()); // no mini-FAT
    file[0x40..0x44].copy_from_slice(&0u32.to_le_bytes());
    file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes()); // no DIFAT chain
    file[0x48..0x4C].copy_from_slice(&0u32.to_le_bytes());
    // DIFAT[0] = sector 1 (the FAT); the rest are FREE_SECT (already zero).
    file[0x4C..0x50].copy_from_slice(&1u32.to_le_bytes());
    for i in 1..109 {
        let off = 0x4C + i * 4;
        file[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
    }
}

fn write_fat(file: &mut [u8], wd_sectors: usize, zero_table_sectors: usize) {
    let fat_off = 512 + 512;
    let mut set = |idx: usize, val: u32| {
        let off = fat_off + idx * 4;
        file[off..off + 4].copy_from_slice(&val.to_le_bytes());
    };
    set(0, END_OF_CHAIN); // directory (single sector)
    set(1, FAT_SECT); // FAT

    let wd_end = 2 + wd_sectors;
    for k in 2..wd_end {
        let val = if k + 1 < wd_end {
            (k + 1) as u32
        } else {
            END_OF_CHAIN
        };
        set(k, val);
    }
    let zt_start = 2 + wd_sectors;
    let zt_end = zt_start + zero_table_sectors;
    for k in zt_start..zt_end {
        let val = if k + 1 < zt_end {
            (k + 1) as u32
        } else {
            END_OF_CHAIN
        };
        set(k, val);
    }
    for k in zt_end..128 {
        set(k, FREE_SECT);
    }
}

fn write_dir_entry(
    buf: &mut [u8],
    name: &str,
    entry_type: u8,
    child: u32,
    start_sector: u32,
    stream_size: u32,
) {
    write_dir_entry_with_sibling(buf, name, entry_type, child, NO_ENTRY, start_sector, stream_size)
}

/// As `write_dir_entry`, with an explicit right-sibling link.
///
/// `CfbReader::find_entry` walks the directory as a real red-black tree
/// from the root entry's own `child` pointer, not a flat scan — so a
/// stream with no sibling link from the root is present in the file but
/// unreachable, and `open_stream` reports it missing. Every
/// entry after the first one under a given parent needs a `right` link
/// to the next, or it simply never gets visited.
fn write_dir_entry_with_sibling(
    buf: &mut [u8],
    name: &str,
    entry_type: u8,
    child: u32,
    right_sibling: u32,
    start_sector: u32,
    stream_size: u32,
) {
    let utf16: Vec<u16> = name.encode_utf16().collect();
    for (i, &ch) in utf16.iter().enumerate() {
        let bytes = ch.to_le_bytes();
        buf[i * 2] = bytes[0];
        buf[i * 2 + 1] = bytes[1];
    }
    let name_size = ((utf16.len() + 1) * 2) as u16;
    buf[0x40..0x42].copy_from_slice(&name_size.to_le_bytes());
    buf[0x42] = entry_type;
    buf[0x43] = 1; // black
    buf[0x44..0x48].copy_from_slice(&NO_ENTRY.to_le_bytes()); // left
    buf[0x48..0x4C].copy_from_slice(&right_sibling.to_le_bytes());
    buf[0x4C..0x50].copy_from_slice(&child.to_le_bytes());
    buf[0x74..0x78].copy_from_slice(&start_sector.to_le_bytes());
    buf[0x78..0x7C].copy_from_slice(&stream_size.to_le_bytes());
}

// ── BIFF / CFB builders for synthetic .xls ──

/// Wrap `data` in a BIFF record header.
#[allow(dead_code)]
pub fn biff(rt: u16, data: &[u8]) -> Vec<u8> {
    let mut v = rt.to_le_bytes().to_vec();
    v.extend_from_slice(&(data.len() as u16).to_le_bytes());
    v.extend_from_slice(data);
    v
}

/// A minimal CFB v3 container holding one stream at the root, in
/// consecutive sectors.
#[allow(dead_code)]
pub fn cfb_with_stream(name: &str, data: &[u8]) -> Vec<u8> {
    const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
    const FAT_SECT: u32 = 0xFFFF_FFFD;
    const FREE_SECT: u32 = 0xFFFF_FFFF;
    const NO_ENTRY: u32 = 0xFFFF_FFFF;
    let data_sectors = data.len().div_ceil(512).max(1);
    let fat_sectors = (2 + data_sectors).div_ceil(128);
    let total = 1 + fat_sectors + data_sectors; // directory + FAT + data
    let mut file = vec![0u8; 512 * (1 + total)];
    file[0..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
    file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
    file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
    file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
    file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
    file[0x2C..0x30].copy_from_slice(&(fat_sectors as u32).to_le_bytes());
    file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes());
    file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
    file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    for i in 0..109 {
        let v = if i < fat_sectors {
            (1 + i) as u32
        } else {
            FREE_SECT
        };
        file[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&v.to_le_bytes());
    }
    let first_data = (1 + fat_sectors) as u32;
    // Directory: root (child = entry 1), then the stream.
    let dir = 512;
    let write_entry =
        |file: &mut [u8], off: usize, name: &str, kind: u8, child: u32, start: u32, size: u32| {
            let utf16: Vec<u16> = name.encode_utf16().collect();
            for (i, ch) in utf16.iter().enumerate() {
                file[off + i * 2..off + i * 2 + 2].copy_from_slice(&ch.to_le_bytes());
            }
            file[off + 0x40..off + 0x42]
                .copy_from_slice(&(((utf16.len() + 1) * 2) as u16).to_le_bytes());
            file[off + 0x42] = kind;
            file[off + 0x43] = 1;
            file[off + 0x44..off + 0x48].copy_from_slice(&NO_ENTRY.to_le_bytes());
            file[off + 0x48..off + 0x4C].copy_from_slice(&NO_ENTRY.to_le_bytes());
            file[off + 0x4C..off + 0x50].copy_from_slice(&child.to_le_bytes());
            file[off + 0x74..off + 0x78].copy_from_slice(&start.to_le_bytes());
            file[off + 0x78..off + 0x7C].copy_from_slice(&size.to_le_bytes());
        };
    write_entry(&mut file, dir, "Root Entry", 5, 1, END_OF_CHAIN, 0);
    write_entry(&mut file, dir + 128, name, 2, NO_ENTRY, first_data, data.len() as u32);
    // FAT.
    let fat = 512 + 512;
    let mut set = |sector: usize, v: u32| {
        let off = fat + sector * 4;
        file[off..off + 4].copy_from_slice(&v.to_le_bytes());
    };
    set(0, END_OF_CHAIN);
    for i in 0..fat_sectors {
        set(1 + i, FAT_SECT);
    }
    for i in 0..data_sectors {
        let s = first_data as usize + i;
        set(
            s,
            if i + 1 == data_sectors {
                END_OF_CHAIN
            } else {
                (s + 1) as u32
            },
        );
    }
    let data_off = 512 + first_data as usize * 512;
    file[data_off..data_off + data.len()].copy_from_slice(data);
    file
}

/// A minimal CFB v3 container holding the given streams, each in its own
/// run of consecutive sectors (no mini stream, so streams of any size are
/// read through the FAT). A name with `/` places the stream in storages
/// created along its path (`"Macros/VBA/dir"`). Each storage's children
/// form a chain of right siblings.
#[allow(dead_code)]
pub fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    // Directory entries: (name, type, parent index, data index).
    let mut entries: Vec<(String, u8, usize, Option<usize>)> =
        vec![("Root Entry".to_string(), 5, usize::MAX, None)];
    for (si, (path, _)) in streams.iter().enumerate() {
        let parts: Vec<&str> = path.split('/').collect();
        let mut parent = 0usize;
        for (k, part) in parts.iter().enumerate() {
            let is_stream = k + 1 == parts.len();
            let found = entries
                .iter()
                .position(|e| e.2 == parent && e.0 == *part && e.1 == 1);
            parent = match (found, is_stream) {
                (Some(i), false) => i,
                _ => {
                    let kind = if is_stream { 2 } else { 1 };
                    entries.push((part.to_string(), kind, parent, is_stream.then_some(si)));
                    entries.len() - 1
                },
            };
        }
    }
    let sectors: Vec<usize> = streams
        .iter()
        .map(|(_, d)| d.len().div_ceil(512).max(1))
        .collect();
    let dir_sectors = entries.len().div_ceil(4);
    let data_sectors: usize = sectors.iter().sum();
    let mut fat_sectors = 1;
    while fat_sectors * 128 < dir_sectors + fat_sectors + data_sectors {
        fat_sectors += 1;
    }
    assert!(fat_sectors <= 109, "header DIFAT only");
    let total = dir_sectors + fat_sectors + data_sectors;
    let mut file = vec![0u8; 512 * (1 + total)];
    file[0..8].copy_from_slice(&CFB_SIGNATURE);
    file[0x18..0x1A].copy_from_slice(&0x003Eu16.to_le_bytes());
    file[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
    file[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
    file[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
    file[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
    file[0x2C..0x30].copy_from_slice(&(fat_sectors as u32).to_le_bytes());
    file[0x30..0x34].copy_from_slice(&0u32.to_le_bytes()); // directory at sector 0
    file[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
    file[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    file[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    for i in 0..109 {
        let v = if i < fat_sectors {
            (dir_sectors + i) as u32
        } else {
            FREE_SECT
        };
        file[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&v.to_le_bytes());
    }
    let mut fat = vec![FREE_SECT; fat_sectors * 128];
    for i in 0..dir_sectors {
        fat[i] = if i + 1 == dir_sectors {
            END_OF_CHAIN
        } else {
            (i + 1) as u32
        };
    }
    for i in 0..fat_sectors {
        fat[dir_sectors + i] = FAT_SECT;
    }
    // Stream data, in stream order.
    let mut starts = Vec::with_capacity(streams.len());
    let mut next = dir_sectors + fat_sectors;
    for ((_, data), &n) in streams.iter().zip(&sectors) {
        starts.push(next);
        for k in 0..n {
            fat[next + k] = if k + 1 == n {
                END_OF_CHAIN
            } else {
                (next + k + 1) as u32
            };
        }
        let data_off = 512 + next * 512;
        file[data_off..data_off + data.len()].copy_from_slice(data);
        next += n;
    }
    // Directory entries: each storage's first child, siblings chained.
    for (i, (name, kind, parent, data)) in entries.iter().enumerate() {
        let child = entries
            .iter()
            .position(|e| e.2 == i)
            .map_or(NO_ENTRY, |c| c as u32);
        let right = entries
            .iter()
            .enumerate()
            .skip(i + 1)
            .find(|(_, e)| e.2 == *parent)
            .map_or(NO_ENTRY, |(j, _)| j as u32);
        let right = if i == 0 { NO_ENTRY } else { right };
        let (start, size) = match data {
            Some(si) => (starts[*si] as u32, streams[*si].1.len() as u32),
            None => (END_OF_CHAIN, 0),
        };
        let off = 512 + i * 128;
        write_dir_entry_with_sibling(
            &mut file[off..off + 128],
            name,
            *kind,
            child,
            right,
            start,
            size,
        );
    }
    for (i, v) in fat.iter().enumerate() {
        let off = 512 + dir_sectors * 512 + i * 4;
        file[off..off + 4].copy_from_slice(&v.to_le_bytes());
    }
    file
}
