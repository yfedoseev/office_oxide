//! The per-document text budget (`office_oxide::limits`).
//!
//! A shared string referenced from every cell is rendered once per cell;
//! nothing else bounds that product, and a 110 KB workbook rendered to
//! 393 MB. The budget is process-global, so every test in this binary
//! sets the same low value — the tests are safe to run in parallel with
//! each other and must not share a binary with tests that render more
//! than a few kilobytes.

mod common;

use std::io::{Cursor, Write};

use common::{biff, cfb_with_stream};
use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

/// Two 5,000-character cells fit; the third spends the budget.
const BUDGET: usize = 12_000;

/// One 5,000-character shared string referenced from `cells` cells in
/// one row — 5,000 × `cells` characters of output from a few KB of zip.
fn xlsx_fan_out(cells: usize) -> Vec<u8> {
    let big = "x".repeat(5_000);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default();
    let parts: Vec<(&str, String)> = vec![
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()),
        ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#.into()),
        ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#.into()),
        ("xl/sharedStrings.xml", format!(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>{big}</t></si></sst>"#)),
        ("xl/worksheets/sheet1.xml", {
            let mut row = String::from(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#);
            for c in 0..cells {
                row.push_str(&format!(r#"<c r="{}1" t="s"><v>0</v></c>"#, (b'A' + c as u8) as char));
            }
            row.push_str("</row></sheetData></worksheet>");
            row
        }),
    ];
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn open_xlsx(cells: usize) -> Document {
    Document::from_reader(Cursor::new(xlsx_fan_out(cells)), DocumentFormat::Xlsx).unwrap()
}

/// The output stops within the budget (plus the notice and separators)
/// and says so.
/// How many 5,000-character cells an output holds.
fn cells_in(out: &str, ch: char) -> usize {
    out.split(|c: char| c != ch)
        .filter(|run| run.len() >= 5_000)
        .count()
}

fn assert_truncated(surface: &str, out: &str) {
    assert!(
        out.contains("character budget"),
        "{surface}: no truncation notice:\n{}",
        &out[..out.len().min(300)]
    );
    assert_eq!(cells_in(out, 'x'), 2, "{surface}: the two cells that fit are kept");
    assert!(
        out.len() < BUDGET + 400,
        "{surface}: {} chars emitted against a budget of {BUDGET}",
        out.len()
    );
}

#[test]
fn test_xlsx_renderers_stop_at_the_budget_with_a_notice() {
    set_max_text_chars(BUDGET);
    let doc = open_xlsx(8);
    assert_truncated("plain_text", &doc.plain_text());
    assert_truncated("to_markdown", &doc.to_markdown());
    assert_truncated("to_html", &doc.to_html());
    let ir = doc.to_ir();
    let ir_text = ir.plain_text();
    assert_truncated("to_ir", &ir_text);
}

#[test]
fn test_a_workbook_within_the_budget_is_untouched() {
    set_max_text_chars(BUDGET);
    let doc = open_xlsx(2);
    for (surface, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("to_html", doc.to_html()),
    ] {
        assert!(!out.contains("character budget"), "{surface} carries a notice under the budget");
        assert_eq!(cells_in(&out, 'x'), 2, "{surface}: both cells present");
    }
}

/// The value is the process's: every test here lowered it from the
/// default (which would let the eight-cell fan-out through), and a
/// document that fits the lowered value still carries no notice.
#[test]
fn test_the_limit_is_configurable() {
    set_max_text_chars(BUDGET);
    assert_eq!(max_text_chars(), BUDGET);
    assert!(!open_xlsx(2).plain_text().contains("character budget"));
    assert!(open_xlsx(3).plain_text().contains("character budget"));
}

/// A BIFF8 workbook whose one 5,000-character shared string is named by
/// `cells` LABELSST cells: the legacy reader copies the string into each
/// cell at open, so the budget applies there and the workbook is flagged
/// truncated.
fn xls_fan_out(cells: u16) -> Vec<u8> {
    let big = "y".repeat(5_000);
    let mut s = Vec::new();
    let mut bof = 0x0600u16.to_le_bytes().to_vec();
    bof.extend_from_slice(&0x0005u16.to_le_bytes());
    bof.extend_from_slice(&[0u8; 12]);
    s.extend(biff(0x0809, &bof));
    let mut bs = 0u32.to_le_bytes().to_vec();
    bs.extend_from_slice(&[0, 0, 1, 0, b'S']);
    s.extend(biff(0x0085, &bs));
    // SST with one 5,000-char string: the record is split by CONTINUE
    // records at 8,224-byte boundaries in real files; 5,000 + header fits
    // one record.
    let mut sst = (cells as u32).to_le_bytes().to_vec();
    sst.extend_from_slice(&1u32.to_le_bytes());
    sst.extend_from_slice(&(big.len() as u16).to_le_bytes());
    sst.push(0);
    sst.extend_from_slice(big.as_bytes());
    s.extend(biff(0x00FC, &sst));
    s.extend(biff(0x000A, &[]));
    let mut bof = 0x0600u16.to_le_bytes().to_vec();
    bof.extend_from_slice(&0x0010u16.to_le_bytes());
    bof.extend_from_slice(&[0u8; 12]);
    s.extend(biff(0x0809, &bof));
    for c in 0..cells {
        let mut d = 0u16.to_le_bytes().to_vec();
        d.extend_from_slice(&c.to_le_bytes());
        d.extend_from_slice(&0u16.to_le_bytes());
        d.extend_from_slice(&0u32.to_le_bytes());
        s.extend(biff(0x00FD, &d));
    }
    s.extend(biff(0x000A, &[]));
    cfb_with_stream("Workbook", &s)
}

#[test]
fn test_xls_open_stops_copying_shared_strings_at_the_budget() {
    set_max_text_chars(BUDGET);
    let doc = Document::from_reader(Cursor::new(xls_fan_out(8)), DocumentFormat::Xls).unwrap();
    let text = doc.plain_text();
    assert_eq!(cells_in(&text, 'y'), 2, "the two cells that fit are kept");
    assert!(
        text.len() < BUDGET + 100,
        "{} chars materialised against a budget of {BUDGET}",
        text.len()
    );
    assert!(doc.to_ir().metadata.text_truncated, "the workbook is flagged truncated");
}
