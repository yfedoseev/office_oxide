//! The per-document text budget (`office_oxide::limits`).
//!
//! A shared string referenced from every cell is rendered once per cell;
//! nothing else bounds that product, and a 110 KB workbook rendered to
//! 393 MB. The budget is process-global, so every test holds the scope that
//! serialises them and puts the default back.

mod common;

use std::io::{Cursor, Write};

use common::{LimitScope, biff, cfb_with_stream};
use office_oxide::limits::{max_text_chars, set_max_text_chars};
use office_oxide::{Document, DocumentFormat};

/// Two 5,000-character cells fit; the third spends the budget.
const BUDGET: usize = 12_000;

/// One 5,000-character shared string referenced from `cells` cells in
/// one row — 5,000 × `cells` characters of output from a few KB of zip.
fn xlsx_fan_out(cells: usize) -> Vec<u8> {
    let big = "x".repeat(5_000);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
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
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
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
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
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
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
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
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
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

/// A document of `sections` sections that all reference one 5,000-character
/// header part: the part is stored once, but the IR gives each section a
/// copy, and those copies are charged to the budget.
fn docx_shared_header(sections: usize) -> Vec<u8> {
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let mut body = String::new();
    for i in 0..sections {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>s{i}</w:t></w:r><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:pPr></w:p>"#
        ));
    }
    let parts: Vec<(&str, String)> = vec![
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.into()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>"#)),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{wml}" xmlns:r="{rel}"><w:body>{body}</w:body></w:document>"#)),
        ("word/_rels/document.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/header" Target="header1.xml"/></Relationships>"#)),
        ("word/header1.xml", format!(r#"<w:hdr xmlns:w="{wml}"><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:hdr>"#, "z".repeat(5_000))),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn test_docx_ir_stops_copying_a_shared_header_at_the_budget() {
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
    let doc =
        Document::from_reader(Cursor::new(docx_shared_header(8)), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    let with_header = ir.sections.iter().filter(|s| s.header.is_some()).count();
    assert_eq!(with_header, 2, "the two copies that fit are kept");
    assert!(ir.metadata.text_truncated, "the document is flagged truncated");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("character budget")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// A configured limit is the ceiling for every input, whatever its size and
/// whatever the floor a small input would otherwise get.
#[test]
fn test_a_configured_ceiling_applies_to_every_input_size() {
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
    let small = xlsx_fan_out(3);
    let large = xlsx_fan_out(20);
    assert!(large.len() > small.len());
    let cut = |bytes: Vec<u8>| {
        let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap();
        cells_in(&doc.plain_text(), 'x')
    };
    // Both stop after the two cells that fit in the budget.
    assert_eq!((cut(small), cut(large)), (2, 2));
}

/// The truncation notice says where to raise the budget.
#[test]
fn test_the_truncation_notice_names_the_setter() {
    let mut limits = LimitScope::new();
    limits.set(max_text_chars, set_max_text_chars, BUDGET);
    let text = open_xlsx(8).plain_text();
    assert!(
        text.contains("set_max_text_chars"),
        "{}",
        &text[text.len().saturating_sub(300)..]
    );
}

/// A workbook whose one shared string of `chars` characters is named by
/// `cells` cells, plus `padding` incompressible bytes in a part nothing reads.
/// Deflated, the character data costs a few KB however long it is.
fn xlsx_big_string(chars: usize, cells: usize, padding: usize) -> Vec<u8> {
    let big = "x".repeat(chars);
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let deflated: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    let stored: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let mut row = String::from(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
    );
    for r in 1..=cells {
        row.push_str(&format!(r#"<row r="{r}"><c r="A{r}" t="s"><v>0</v></c></row>"#));
    }
    row.push_str("</sheetData></worksheet>");
    let parts: Vec<(&str, String)> = vec![
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()),
        ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#.into()),
        ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#.into()),
        ("xl/sharedStrings.xml", format!(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>{big}</t></si></sst>"#)),
        ("xl/worksheets/sheet1.xml", row),
    ];
    for (name, data) in parts {
        zip.start_file(name, deflated).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    if padding > 0 {
        let mut state = 0x9E37_79B9u32;
        let noise: Vec<u8> = (0..padding)
            .map(|_| {
                state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
                (state >> 24) as u8
            })
            .collect();
        zip.start_file("customXml/unused.bin", stored).unwrap();
        zip.write_all(&noise).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A tiny file is bounded by the floor, not by the whole limit: 3,000 cells of
/// a 32,000-character string (a cell holds at most 32,767) would be 96 million
/// characters, more than the floor of 64 Mi and far less than the 256 Mi limit.
#[test]
fn test_a_small_input_is_bounded_by_the_floor_not_the_whole_limit() {
    let _defaults = common::default_limits();
    let bytes = xlsx_big_string(32_000, 3_000, 0);
    assert!(bytes.len() < 64 * 1024, "the fixture must be a small file: {}", bytes.len());
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap();
    let text = doc.plain_text();
    let chars = text.matches('x').count();
    assert!(
        (60_000_000..=office_oxide::limits::TEXT_BUDGET_FLOOR_CHARS).contains(&chars),
        "{chars}"
    );
    assert!(text.contains("set_max_text_chars"), "the cut is reported");
}

/// The same text from a file of 64 KB or more is bounded by the limit alone
/// (96 million characters fit under 256 Mi), so nothing is cut.
#[test]
fn test_a_larger_input_is_bounded_by_the_limit() {
    let _defaults = common::default_limits();
    let bytes = xlsx_big_string(32_000, 3_000, 80 * 1024);
    assert!(bytes.len() >= 64 * 1024);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap();
    let text = doc.plain_text();
    assert_eq!(text.matches('x').count(), 96_000_000);
    assert!(!text.contains("set_max_text_chars"));
}
