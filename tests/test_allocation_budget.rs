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

mod common;

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::{Cursor, Write};
use std::sync::atomic::{AtomicUsize, Ordering};

struct Counting;

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

thread_local! {
    /// Bytes requested by this thread, so a byte budget is not disturbed
    /// by tests running in parallel.
    static THREAD_BYTES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

fn count_bytes(n: usize) {
    // `try_with`: the allocator also runs during thread teardown.
    let _ = THREAD_BYTES.try_with(|b| b.set(b.get().saturating_add(n)));
}

// SAFETY: every call is forwarded to `System` unchanged; the counter is the
// only addition and touches no allocator state.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        count_bytes(layout.size());
        // SAFETY: same layout contract as the caller's.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        // SAFETY: `ptr` came from `System.alloc` with this layout.
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        count_bytes(new_size);
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

/// The allocation counter is process-wide, so one test's work would land
/// in another's window if they ran in parallel.
fn serial() -> std::sync::MutexGuard<'static, ()> {
    static SERIAL: std::sync::Mutex<()> = std::sync::Mutex::new(());
    SERIAL.lock().unwrap_or_else(|e| e.into_inner())
}

/// Bytes this thread requested from the allocator while running `f`.
fn bytes_during<T>(f: impl FnOnce() -> T) -> (T, usize) {
    let before = THREAD_BYTES.with(std::cell::Cell::get);
    let out = f();
    (out, THREAD_BYTES.with(std::cell::Cell::get) - before)
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
    let _serial = serial();
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
    let _serial = serial();
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

use common::{biff, cfb_with_stream};

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
    let _serial = serial();
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

// ---------------------------------------------------------------- DOCX

/// A table whose widest row is wide (here, ten cells spanning 1,000 grid
/// columns each) and which has many ordinary one-cell rows. The per-span
/// and per-table clamps bound the *width*, but the converter and the
/// renderers sized dense `rows x width` grids, so every narrow row cost
/// the width of the widest one: quadratic in a few kilobytes of
/// compressed XML.
#[test]
fn test_a_wide_row_does_not_make_every_row_cost_the_full_width() {
    let _serial = serial();
    const NARROW_ROWS: usize = 2_000;
    let docx = |span: u32| {
        let mut body = String::from("<w:tbl><w:tr>");
        for _ in 0..10 {
            body.push_str(&format!(
                r#"<w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr><w:p/></w:tc>"#
            ));
        }
        body.push_str("</w:tr>");
        for i in 0..NARROW_ROWS {
            body.push_str(&format!(
                "<w:tr><w:tc><w:p><w:r><w:t>r{i}</w:t></w:r></w:p></w:tc></w:tr>"
            ));
        }
        body.push_str("</w:tbl>");
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
        );
        let bytes = zip_of(&[
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
            ("word/document.xml", xml.as_bytes()),
        ]);
        office_oxide::Document::from_reader(Cursor::new(bytes), office_oxide::DocumentFormat::Docx)
            .unwrap()
    };
    // Bytes requested by to_ir(), to_markdown() and plain_text().
    let measure = |doc: &office_oxide::Document| {
        let (ir, ir_bytes) = bytes_during(|| doc.to_ir());
        let (md, md_bytes) = bytes_during(|| ir.to_markdown());
        let (text, text_bytes) = bytes_during(|| ir.plain_text());
        ([ir_bytes, md_bytes, text_bytes], md, text)
    };
    // The same table with one-column cells is the linear baseline.
    let (baseline, _, _) = measure(&docx(1));
    let (wide, md, text) = measure(&docx(1_000));
    for (what, (w, b)) in ["to_ir()", "to_markdown()", "plain_text()"]
        .iter()
        .zip(wide.iter().zip(baseline.iter()))
    {
        assert!(
            *w <= b * 3 + (1 << 20),
            "{what} requested {w} bytes for the wide table, {b} for the narrow one"
        );
    }
    // Every row's content is still there.
    for i in [0, NARROW_ROWS / 2, NARROW_ROWS - 1] {
        let needle = format!("r{i}");
        assert!(md.contains(&needle) && text.contains(&needle), "row {i} lost");
    }
}

/// Flattening an HTML `w:altChunk` lowercased the whole remaining input
/// once per character to look for `<script`/`</script>`: quadratic bytes
/// copied, so a 1 MB script block in a mail-merge chunk was ~10^12.
#[test]
fn test_html_alt_chunk_flattening_is_linear() {
    let _serial = serial();
    let html = |script_len: usize| {
        format!(
            "<html><body><p>Before</p><SCRIPT>{}</Script><p>After &amp; done</p></body></html>",
            "x<y;".repeat(script_len / 4)
        )
    };
    let docx = |html: &str| {
        zip_of(&[
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/chunk.html" ContentType="text/html"/></Types>"#,
            ),
            (
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
            (
                "word/_rels/document.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.html"/></Relationships>"#,
            ),
            (
                "word/document.xml",
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:altChunk r:id="rId9"/></w:body></w:document>"#,
            ),
            ("word/chunk.html", html.as_bytes()),
        ])
    };
    const SCRIPT: usize = 64 * 1024;
    let big = docx(&html(SCRIPT));
    let small = docx(&html(4));
    let open = |b: Vec<u8>| office_oxide::docx::DocxDocument::from_reader(Cursor::new(b)).unwrap();
    let (_, baseline) = bytes_during(|| open(small));
    let (doc, bytes) = bytes_during(|| open(big));
    let text = doc.plain_text();
    assert!(text.contains("Before") && text.contains("After & done"), "{text:?}");
    assert!(!text.contains("x<y"), "script body leaked: {}", &text[..text.len().min(80)]);
    // A handful of copies of the input, not one per character.
    assert!(
        bytes.saturating_sub(baseline) < 32 * SCRIPT,
        "{bytes} bytes requested for a {SCRIPT}-byte script ({baseline} for a tiny one)"
    );
}
