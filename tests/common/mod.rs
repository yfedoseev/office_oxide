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

/// Sentinel directory id: no sibling / child.
const NO_ENTRY: u32 = 0xFFFF_FFFF;

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

/// Cell paragraph grpprl: `sprmPFInTable` (0x2416)=1, `sprmPItap` (0x6649)=1.
#[allow(dead_code)]
pub fn cell_grpprl() -> Vec<u8> {
    vec![0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00]
}

/// Row-terminator grpprl: `fInTable` + `fInTableTtp` + `itap`, then
/// `sprmTDefTable` (0xD608) carrying the column boundaries and per-cell
/// `rgf` flags (the TAP). `centers` has `rgfs.len() + 1` entries (column
/// edges in twips); `rgfs[i]` is the `rgf` flags word of cell `i` (MS-DOC
/// `TCGRF` 2-bit `fVertMerge` field: `fvmMerge` = 0x0020, `fvmRestart` = 0x0060).
#[allow(dead_code)]
pub fn row_grpprl(centers: &[i16], rgfs: &[u16]) -> Vec<u8> {
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
        0x16, 0x24, 0x01, 0x17, 0x24, 0x01, 0x49, 0x66, 0x01, 0x00, 0x00, 0x00,
    ];
    head.extend_from_slice(&g);
    head
}

/// Build a complete synthetic `.doc` from the given paragraphs.
pub fn build_doc(paras: &[Para]) -> Vec<u8> {
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
    let text_bytes: Vec<u8> = units.iter().flat_map(|u| u.to_le_bytes()).collect();

    // Text lives after the FIB page (page 0) and the N FKP pages (pages 1..N).
    let text_offset = ((n as u32) + 1) * 512;

    // ── 0Table stream: CLX (piece table) followed by PlcfBtePapx. ──
    let mut table = build_clx(text_offset, text_len);
    let fc_plcf = table.len() as u32;
    table.extend_from_slice(&build_plcf_bte_papx(n, &cp_starts, text_len));
    let lcb_plcf = (table.len() as u32) - fc_plcf;

    // ── WordDocument stream: FIB + N FKP pages + text. ──
    let wd_len = text_offset as usize + text_bytes.len();
    let wd_sectors = wd_len.div_ceil(512);
    let mut word_doc = vec![0u8; wd_sectors * 512];
    write_fib(&mut word_doc, text_len, fc_plcf, lcb_plcf);
    for (i, p) in paras.iter().enumerate() {
        let cp0 = cp_starts[i];
        let cp1 = if i + 1 < n {
            cp_starts[i + 1]
        } else {
            text_len
        };
        let fc0 = text_offset + cp0 * 2;
        let fc1 = text_offset + cp1 * 2;
        let page = build_fkp_page(fc0, fc1, &p.grpprl);
        let off = (i + 1) * 512;
        word_doc[off..off + 512].copy_from_slice(&page);
    }
    word_doc[text_offset as usize..text_offset as usize + text_bytes.len()]
        .copy_from_slice(&text_bytes);

    build_cfb(&word_doc, &table)
}

/// Open a synthetic `.doc` byte buffer through the public API.
pub fn open_doc(bytes: &[u8]) -> Document {
    Document::from_reader(Cursor::new(bytes.to_vec()), DocumentFormat::Doc)
        .expect("synthetic .doc must parse")
}

// ── CFB / DOC byte construction ──

/// Write the FIB into the first 512 bytes of `word_doc`.
fn write_fib(word_doc: &mut [u8], text_len: u32, fc_plcf: u32, lcb_plcf: u32) {
    word_doc[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes()); // wIdent = Word 97+
    word_doc[2..4].copy_from_slice(&0x00C1u16.to_le_bytes()); // nFib
    // flags at 0x0A: bit 9 (fWhichTblStm) clear → use 0Table.
    word_doc[0x0A..0x0C].copy_from_slice(&0u16.to_le_bytes());
    word_doc[0x4C..0x50].copy_from_slice(&text_len.to_le_bytes()); // ccpText
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
        v.extend_from_slice(&((i as u32) + 1).to_le_bytes()); // page number
    }
    v
}

/// Build one 512-byte PAPX FKP page holding a single paragraph.
fn build_fkp_page(fc0: u32, fc1: u32, grpprl: &[u8]) -> [u8; 512] {
    let mut page = [0u8; 512];
    page[0..4].copy_from_slice(&fc0.to_le_bytes());
    page[4..8].copy_from_slice(&fc1.to_le_bytes());

    // PAPX header = [cw][istd:2][grpprl]. Pad grpprl to an odd length so that
    // `3 + grpprl.len()` is even and `cw = (3 + grpprl.len()) / 2`.
    let mut g = grpprl.to_vec();
    if g.len().is_multiple_of(2) {
        g.push(0);
    }
    let cw = ((3 + g.len()) / 2) as u8;
    let mut papx = vec![cw, 0, 0];
    papx.extend_from_slice(&g);

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
    write_dir_entry(&mut file[dir_off..dir_off + 128], "Root Entry", 5, NO_ENTRY, END_OF_CHAIN, 0);
    let wd_start = 2u32;
    let zt_start = (2 + wd_sectors) as u32;
    write_dir_entry(
        &mut file[dir_off + 128..dir_off + 256],
        "WordDocument",
        2,
        NO_ENTRY,
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
    buf[0x48..0x4C].copy_from_slice(&NO_ENTRY.to_le_bytes()); // right
    buf[0x4C..0x50].copy_from_slice(&child.to_le_bytes());
    buf[0x74..0x78].copy_from_slice(&start_sector.to_le_bytes());
    buf[0x78..0x7C].copy_from_slice(&stream_size.to_le_bytes());
}
