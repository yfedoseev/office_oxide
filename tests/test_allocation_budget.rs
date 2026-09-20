//! Heap-allocation budgets for the hot parsing loops.
//!
//! A constant-factor speed-up cannot be pinned by a timing assertion — it
//! depends on the machine and on what else is running — but the number of
//! heap allocations a loop makes per element is deterministic. Each test
//! here builds a document whose size is dominated by one repeated element,
//! counts allocations across the parse, and asserts a per-element budget
//! that the old per-element `String`s and attribute rescans blew through.
//!
//! The counter is a global allocator, so this file is its own test binary.

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write};
use std::sync::atomic::{AtomicUsize, Ordering};

struct Counting;

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

// SAFETY: every call is forwarded to `System` unchanged; the counter is the
// only addition and touches no allocator state.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        // SAFETY: forwarded unchanged.
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn allocations_during<T>(f: impl FnOnce() -> T) -> (T, usize) {
    let before = ALLOCATIONS.load(Ordering::SeqCst);
    let out = f();
    (out, ALLOCATIONS.load(Ordering::SeqCst) - before)
}

fn zip_of(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
    for (name, data) in parts {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn column(mut i: u32) -> String {
    let mut s = String::new();
    i += 1;
    while i > 0 {
        let r = (i - 1) % 26;
        s.insert(0, (b'A' + r as u8) as char);
        i = (i - 1) / 26;
    }
    s
}

/// One sheet of `rows` × `cols` cells produced by `cell(row, col)`.
fn xlsx_with_cells(rows: u32, cols: u32, cell: impl Fn(u32, u32) -> String) -> Vec<u8> {
    let mut sheet = String::from(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
    );
    for r in 1..=rows {
        sheet.push_str(&format!(r#"<row r="{r}">"#));
        for c in 0..cols {
            sheet.push_str(&cell(r, c));
        }
        sheet.push_str("</row>");
    }
    sheet.push_str("</sheetData></worksheet>");
    let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let workbook = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let styles = br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0" applyFont="1"/></cellXfs></styleSheet>"#;
    zip_of(&[
        ("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#),
        ("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ("xl/_rels/workbook.xml.rels", rels),
        ("xl/workbook.xml", workbook),
        ("xl/styles.xml", styles),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
    ])
}

fn open_xlsx(bytes: Vec<u8>) -> office_oxide::Document {
    office_oxide::Document::from_reader(Cursor::new(bytes), office_oxide::DocumentFormat::Xlsx)
        .unwrap()
}

/// A numeric cell (`<c r="B7"><v>12.5</v></c>`) used to cost three heap
/// strings — the reference, the unescaped text, and the text buffer it was
/// appended to — plus one attribute rescan per key. It now costs none:
/// the reference and the number parse from the borrowed XML.
#[test]
fn test_numeric_cells_parse_without_a_heap_allocation_per_cell() {
    const ROWS: u32 = 200;
    const COLS: u32 = 50;
    let bytes = xlsx_with_cells(ROWS, COLS, |r, c| {
        format!(r#"<c r="{}{r}"><v>{}.5</v></c>"#, column(c), r * 100 + c)
    });
    let baseline = allocations_during(|| {
        open_xlsx(xlsx_with_cells(1, 1, |_, _| r#"<c r="A1"><v>1</v></c>"#.into()))
    })
    .1;
    let (doc, allocs) = allocations_during(|| open_xlsx(bytes));
    let per_cell = (allocs.saturating_sub(baseline)) as f64 / (ROWS * COLS) as f64;
    assert_eq!(doc.to_ir().sections.len(), 1);
    assert!(
        per_cell < 0.5,
        "{allocs} allocations for {} numeric cells ({per_cell:.2} per cell, {baseline} fixed)",
        ROWS * COLS
    );
}

/// Excel writes `<c r="B7" s="3"/>` for every formatted-but-empty cell in
/// the used range; each one used to allocate a `String` for the reference
/// it then parsed and dropped.
#[test]
fn test_empty_styled_cells_parse_without_a_heap_allocation_per_cell() {
    const ROWS: u32 = 200;
    const COLS: u32 = 50;
    let bytes = xlsx_with_cells(ROWS, COLS, |r, c| format!(r#"<c r="{}{r}" s="1"/>"#, column(c)));
    let baseline = allocations_during(|| {
        open_xlsx(xlsx_with_cells(1, 1, |_, _| r#"<c r="A1" s="1"/>"#.into()))
    })
    .1;
    let (doc, allocs) = allocations_during(|| open_xlsx(bytes));
    let per_cell = (allocs.saturating_sub(baseline)) as f64 / (ROWS * COLS) as f64;
    assert_eq!(doc.to_ir().sections.len(), 1);
    assert!(
        per_cell < 0.5,
        "{allocs} allocations for {} empty cells ({per_cell:.2} per cell, {baseline} fixed)",
        ROWS * COLS
    );
}

// ---------------------------------------------------------------- XLS

/// Wrap `data` in a BIFF record header.
fn biff(rt: u16, data: &[u8]) -> Vec<u8> {
    let mut v = rt.to_le_bytes().to_vec();
    v.extend_from_slice(&(data.len() as u16).to_le_bytes());
    v.extend_from_slice(data);
    v
}

/// A minimal CFB v3 container holding one stream at the root, in
/// consecutive sectors.
fn cfb_with_stream(name: &str, data: &[u8]) -> Vec<u8> {
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

/// A BIFF8 workbook with `strings` shared strings and one sheet of
/// `rows` × `cols` cells alternating `LABELSST` and `NUMBER`.
fn xls_with_cells(rows: u16, cols: u16, strings: &[&str]) -> Vec<u8> {
    let mut s = Vec::new();
    let mut bof = 0x0600u16.to_le_bytes().to_vec();
    bof.extend_from_slice(&0x0005u16.to_le_bytes());
    bof.extend_from_slice(&[0u8; 12]);
    s.extend(biff(0x0809, &bof));
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    let mut sst = ((rows as u32) * (cols as u32)).to_le_bytes().to_vec();
    sst.extend_from_slice(&(strings.len() as u32).to_le_bytes());
    for t in strings {
        sst.extend_from_slice(&(t.len() as u16).to_le_bytes());
        sst.push(0);
        sst.extend_from_slice(t.as_bytes());
    }
    s.extend(biff(0x00FC, &sst));
    s.extend(biff(0x000A, &[]));
    let mut bof = 0x0600u16.to_le_bytes().to_vec();
    bof.extend_from_slice(&0x0010u16.to_le_bytes());
    bof.extend_from_slice(&[0u8; 12]);
    s.extend(biff(0x0809, &bof));
    for r in 0..rows {
        for c in 0..cols {
            let mut d = r.to_le_bytes().to_vec();
            d.extend_from_slice(&c.to_le_bytes());
            d.extend_from_slice(&0u16.to_le_bytes());
            if c % 2 == 0 {
                d.extend_from_slice(
                    &(((r as usize + c as usize) % strings.len()) as u32).to_le_bytes(),
                );
                s.extend(biff(0x00FD, &d));
            } else {
                d.extend_from_slice(&((r as f64) * 10.0 + c as f64 + 0.5).to_le_bytes());
                s.extend(biff(0x0203, &d));
            }
        }
    }
    s.extend(biff(0x000A, &[]));
    cfb_with_stream("Workbook", &s)
}

/// Every BIFF record was copied into its own `Vec`, every cell's text
/// was rendered into a `display` grid at `open()`, and `plain_text()`
/// copied each cell a third time. A string cell now costs its one copy
/// out of the shared-string table and a number cell nothing at `open()`;
/// rendering borrows.
#[test]
fn test_xls_cells_open_within_one_allocation_each_and_render_borrowing() {
    const ROWS: u16 = 200;
    const COLS: u16 = 50;
    let strings = ["alpha", "beta", "gamma", "delta", "epsilon"];
    let bytes = xls_with_cells(ROWS, COLS, &strings);
    let open = |b: Vec<u8>| {
        office_oxide::Document::from_reader(Cursor::new(b), office_oxide::DocumentFormat::Xls)
            .unwrap()
    };
    let baseline = allocations_during(|| open(xls_with_cells(1, 2, &strings))).1;
    let (doc, allocs) = allocations_during(|| open(bytes));
    let cells = (ROWS as usize) * (COLS as usize);
    let per_cell = allocs.saturating_sub(baseline) as f64 / cells as f64;
    assert!(
        per_cell < 0.9,
        "{allocs} allocations opening {cells} cells ({per_cell:.2} per cell, {baseline} fixed)"
    );
    let (text, render_allocs) = allocations_during(|| doc.plain_text());
    assert!(
        text.contains("alpha") && text.contains("\t1.5\t"),
        "{}",
        &text[..200.min(text.len())]
    );
    let per_cell = render_allocs as f64 / cells as f64;
    // A number still formats into a temporary; a string must not copy.
    assert!(
        per_cell < 0.8,
        "{render_allocs} allocations rendering {cells} cells ({per_cell:.2} per cell)"
    );
}
