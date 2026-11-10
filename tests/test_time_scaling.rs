//! Reading time grows with the input, not faster.
//!
//! The limits bound sizes and counts, not time, so a reader or converter
//! whose work is quadratic in something the file controls turns a modest
//! file into minutes of CPU with no limit tripping. Each shape here is one
//! that once was quadratic, run end to end through the public API at two
//! sizes: four times the input must cost well under sixteen times the time.
//! Each run is timed as the fastest of three, at a size that takes at least
//! 20 ms, and the bound (10x for 4x the input) leaves room for a busy machine
//! while a quadratic path (16x) still fails.

mod common;

use std::io::Cursor;
use std::time::{Duration, Instant};

use office_oxide::{Document, DocumentFormat};

fn fastest_of_three(run: &dyn Fn()) -> Duration {
    (0..3)
        .map(|_| {
            let t = Instant::now();
            run();
            t.elapsed()
        })
        .min()
        .unwrap()
}

/// The tests time one at a time: run alongside each other, one test's small
/// input would be timed while the others were busy and its large one after,
/// which hides a quadratic path.
static TIMING: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Assert that `run(make(4n))` takes under 10x `run(make(n))`, for an `n` at
/// least `start` at which one run takes 20 ms or more.
fn assert_linear<T>(what: &str, start: usize, make: impl Fn(usize) -> T, run: impl Fn(&T)) {
    let _one_at_a_time = TIMING.lock().unwrap_or_else(|e| e.into_inner());
    let mut n = start;
    let mut base;
    loop {
        let input = make(n);
        base = fastest_of_three(&|| run(&input));
        if base >= Duration::from_millis(20) {
            break;
        }
        // Sub-millisecond runs measure nothing: a shape that never takes
        // 20 ms needs a larger start, not a pass.
        assert!(n < start << 10, "{what}: {n} took only {base:?}; start larger");
        n *= 2;
    }
    let input = make(4 * n);
    let big = fastest_of_three(&|| run(&input));
    let ratio = big.as_secs_f64() / base.as_secs_f64().max(1e-9);
    assert!(
        ratio < 10.0,
        "{what}: {n} took {base:?}, {} took {big:?} ({ratio:.1}x for 4x the input)",
        4 * n
    );
}

fn zip_of(parts: &[(&str, &[u8])]) -> Vec<u8> {
    use std::io::Write;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    for (name, data) in parts {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn docx(body: &str) -> Vec<u8> {
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    zip_of(&[
        (
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", xml.as_bytes()),
    ])
}

fn open(bytes: &[u8], format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes.to_vec()), format).unwrap()
}

/// A table whose first row spans a wide grid, and `rows` one-cell rows: the
/// layout once padded every row to the widest.
#[test]
fn test_a_wide_table_converts_in_time_linear_in_its_rows() {
    let _limits = common::default_limits();
    assert_linear(
        "wide docx table",
        250,
        |rows| {
            let mut body = String::from("<w:tbl><w:tr>");
            for _ in 0..10 {
                body.push_str(r#"<w:tc><w:tcPr><w:gridSpan w:val="500"/></w:tcPr><w:p/></w:tc>"#);
            }
            body.push_str("</w:tr>");
            for i in 0..rows {
                body.push_str(&format!(
                    "<w:tr><w:tc><w:p><w:r><w:t>r{i}</w:t></w:r></w:p></w:tc></w:tr>"
                ));
            }
            body.push_str("</w:tbl>");
            docx(&body)
        },
        |bytes| {
            let doc = open(bytes, DocumentFormat::Docx);
            let _ = doc.to_markdown();
            let _ = doc.to_html();
        },
    );
}

/// Unclosed emphasis openers: inline Markdown once rescanned the rest of
/// the line for each.
#[test]
fn test_markdown_with_unclosed_openers_parses_in_linear_time() {
    assert_linear(
        "markdown openers",
        2_000,
        |n| "*a _b `c [d ".repeat(n),
        |md| {
            let _ = office_oxide::DocumentIR::from_markdown(md, DocumentFormat::Docx);
        },
    );
}

/// A word split over one run per letter: replacing across runs once
/// re-joined the run text for every match.
#[test]
fn test_replacing_across_many_runs_takes_linear_time() {
    let _limits = common::default_limits();
    assert_linear(
        "docx replace across runs",
        500,
        |n| {
            let runs: String = (0..n)
                .map(|i| format!("<w:r><w:t>{}</w:t></w:r>", if i % 2 == 0 { "a" } else { "b" }))
                .collect();
            docx(&format!("<w:p>{runs}</w:p>"))
        },
        |bytes| {
            let mut doc = office_oxide::edit::EditableDocument::from_reader(
                Cursor::new(bytes.clone()),
                DocumentFormat::Docx,
            )
            .unwrap();
            let _ = doc.replace_text("ab", "x").unwrap();
        },
    );
}

/// Many sheets with the same requested name: naming once searched every
/// earlier sheet for each.
#[test]
fn test_writing_many_clashing_sheet_names_takes_linear_time() {
    assert_linear(
        "xlsx sheet names",
        200,
        |n| n,
        |&n| {
            let mut w = office_oxide::xlsx::write::XlsxWriter::new();
            for _ in 0..n {
                w.add_sheet("Sheet");
            }
            let mut out = Cursor::new(Vec::new());
            w.write_to(&mut out).unwrap();
        },
    );
}

/// A workbook of many small sheets, read back.
#[test]
fn test_a_many_sheet_workbook_reads_in_linear_time() {
    let _limits = common::default_limits();
    assert_linear(
        "xlsx many sheets",
        50,
        |n| {
            let mut w = office_oxide::xlsx::write::XlsxWriter::new();
            for i in 0..n {
                w.add_sheet(&format!("S{i}")).set_cell(
                    0,
                    0,
                    office_oxide::xlsx::write::CellData::String(format!("v{i}")),
                );
            }
            let mut out = Cursor::new(Vec::new());
            w.write_to(&mut out).unwrap();
            out.into_inner()
        },
        |bytes| {
            let doc = open(bytes, DocumentFormat::Xlsx);
            let _ = doc.plain_text();
            let _ = doc.to_ir();
        },
    );
}

/// A legacy workbook's cells, opened and converted.
#[test]
fn test_a_large_xls_reads_in_linear_time() {
    let _limits = common::default_limits();
    assert_linear(
        "xls cells",
        100,
        |rows| {
            let rows = u16::try_from(rows).expect("a BIFF8 sheet has at most 65,536 rows");
            common::xls_with_cells(rows, 20, &["alpha", "beta", "gamma"])
        },
        |bytes| {
            let doc = open(bytes, DocumentFormat::Xls);
            let _ = doc.plain_text();
            let _ = doc.to_ir();
        },
    );
}

/// A long legacy Word document.
#[test]
fn test_a_long_doc_reads_in_linear_time() {
    let _limits = common::default_limits();
    assert_linear(
        "doc paragraphs",
        200,
        |n| {
            let paras: Vec<common::Para> = (0..n)
                .map(|_| common::Para {
                    text: "A paragraph of legacy text.",
                    terminator: '\r',
                    grpprl: common::prose_grpprl(),
                })
                .collect();
            // Past one FAT sector's 64 KB: the streams in a container of
            // their own size.
            let (word, table) = common::build_doc_streams(
                &paras,
                &common::Subdocs::default(),
                common::FibTweaks::default(),
            );
            common::cfb_with_streams(&[("WordDocument", &word), ("0Table", &table)])
        },
        |bytes| {
            let doc = open(bytes, DocumentFormat::Doc);
            let _ = doc.plain_text();
            let _ = doc.to_ir();
        },
    );
}
