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

