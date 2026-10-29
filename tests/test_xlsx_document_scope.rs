//! When the allowance for a whole workbook's IR is spent, the sheet that
//! spent it is cut and says so; the sheets after it are still converted, each
//! within what is left, and none disappears without a notice that names it.
//!
//! Limits are process-global, so every test holds the `LimitScope` of [`serial`],
//! sets its own through it, and has the defaults back when it ends.

use std::io::{Cursor, Write};

use office_oxide::ir::{Element, Section};
use office_oxide::limits::{
    max_memory_bytes, max_text_chars, set_max_memory_bytes, set_max_text_chars,
};
use office_oxide::{Document, DocumentFormat};

mod common;
use common::LimitScope;

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn serial() -> LimitScope {
    LimitScope::new()
}

fn cell_bytes() -> usize {
    std::mem::size_of::<office_oxide::ir::TableCell>()
        + std::mem::size_of::<office_oxide::ir::Element>()
}

fn package(sheets: &[(&str, String)]) -> Vec<u8> {
    let mut names = String::new();
    let mut rels = String::new();
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for (i, (name, xml)) in (1..).zip(sheets) {
        names.push_str(&format!(r#"<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>"#));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        ));
        parts.push((format!("xl/worksheets/sheet{i}.xml"), xml.clone().into_bytes()));
    }
    parts.push((
        "[Content_Types].xml".into(),
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
    ));
    parts.push((
        "_rels/.rels".into(),
        format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).into_bytes(),
    ));
    parts.push((
        "xl/workbook.xml".into(),
        format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets>{names}</sheets></workbook>"#)
            .into_bytes(),
    ));
    parts.push((
        "xl/_rels/workbook.xml.rels".into(),
        format!(r#"<Relationships xmlns="{REL_NS}">{rels}</Relationships>"#).into_bytes(),
    ));
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    for (name, data) in parts {
        zip.start_file(name.as_str(), opts).unwrap();
        zip.write_all(&data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// `rows` rows of a label in column A and a number in column `last` (`AD`
/// is 30 wide: 28 empty cells of padding a row), or in column B for a dense
/// sheet that needs none.
fn sheet(tag: &str, rows: usize, last: &str) -> String {
    let mut xml = format!(r#"<worksheet xmlns="{NS}"><sheetData>"#);
    for r in 1..=rows {
        xml.push_str(&format!(
            r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{tag}r{r}</t></is></c><c r="{last}{r}"><v>{r}</v></c></row>"#
        ));
    }
    xml.push_str("</sheetData></worksheet>");
    xml
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

fn table_rows(section: &Section) -> usize {
    section
        .elements
        .iter()
        .map(|e| match e {
            Element::Table(t) => t.rows.len(),
            _ => 0,
        })
        .sum()
}

fn text_of(section: &Section) -> String {
    serde_json::to_string(section).unwrap()
}

/// 100 padded rows, a dense sheet, and 10 more padded rows, against an
/// allowance of 1,680 empty cells (60 rows of padding).
fn big_dense_padded() -> Vec<u8> {
    package(&[
        ("Big", sheet("a", 100, "AD")),
        ("Second", sheet("b", 5, "B")),
        ("Third", sheet("c", 10, "AD")),
    ])
}

#[test]
fn test_a_sheet_after_the_one_that_spent_the_padding_allowance_is_still_converted() {
    let mut limits = serial();
    // The document-wide padding allowance is half the memory budget.
    limits.set(max_memory_bytes, set_max_memory_bytes, 2 * (1_680 * cell_bytes()) as u64);
    let ir = open(big_dense_padded()).to_ir();
    let titles: Vec<_> = ir
        .sections
        .iter()
        .map(|s| s.title.clone().unwrap_or_default())
        .collect();
    assert_eq!(titles, ["Big", "Second", "Third"], "a sheet vanished");

    // The sheet that spent the allowance is cut and says so.
    let big = &ir.sections[0];
    assert!((59..100).contains(&table_rows(big)), "{} rows", table_rows(big));
    assert!(
        text_of(big).contains("workbook truncated"),
        "no notice on the sheet that was cut"
    );

    // A sheet that needs no more padding is whole, and not flagged.
    let second = &ir.sections[1];
    assert_eq!(table_rows(second), 5);
    assert!(text_of(second).contains("br5"));
    assert!(!text_of(second).contains("truncated"));

    // A later sheet that needs padding is cut at the rows it cannot pay
    // for, with a notice of its own.
    let third = &ir.sections[2];
    assert_eq!(table_rows(third), 0);
    assert!(text_of(third).contains("workbook truncated"), "no notice on the later sheet");

    assert!(ir.metadata.text_truncated);
    for name in ["Big", "Third"] {
        assert!(
            ir.metadata.warnings.iter().any(|w| w.contains(name)),
            "no warning names {name}: {:?}",
            ir.metadata.warnings
        );
    }
    assert!(
        !ir.metadata.warnings.iter().any(|w| w.contains("Second")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_every_renderer_agrees_the_dense_sheet_is_whole() {
    let mut limits = serial();
    // The document-wide padding allowance is half the memory budget.
    limits.set(max_memory_bytes, set_max_memory_bytes, 2 * (1_680 * cell_bytes()) as u64);
    let doc = open(big_dense_padded());
    for (name, out) in [
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("to_html", doc.to_html()),
    ] {
        assert!(out.contains("br5"), "{name} lost the dense sheet");
    }
}

#[test]
fn test_sheets_left_out_by_the_text_budget_are_named() {
    let mut limits = serial();
    // Four sheets of 40 rows against a budget the first spends and the second
    // runs out in.
    limits.set(max_text_chars, set_max_text_chars, 300);
    let doc = open(package(&[
        ("One", sheet("a", 40, "B")),
        ("Two", sheet("b", 40, "B")),
        ("Three", sheet("c", 40, "B")),
        ("Four", sheet("d", 40, "B")),
    ]));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    let all = serde_json::to_string(&ir.sections).unwrap();
    for name in ["Three", "Four"] {
        assert!(all.contains(name), "{name} is left out without being named");
    }
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("Three") && w.contains("Four")),
        "{:?}",
        ir.metadata.warnings
    );
    // The sheet-by-sheet IR names them too.
    let items: Vec<_> = doc.as_xlsx().unwrap().sheets_ir().collect();
    let all = items
        .iter()
        .map(|i| text_of(&i.section))
        .collect::<String>();
    assert!(all.contains("Three") && all.contains("Four"), "sheets_ir drops them silently");
}
