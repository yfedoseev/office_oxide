//! A worksheet whose XML is malformed part-way through.
//!
//! The rows before the damage parsed fine. They are kept, the sheet is marked
//! as cut short in the text, the warnings and `text_truncated`, and the other
//! sheets are read as usual. A sheet damaged before its first complete row is
//! still reported as unreadable.

use std::io::{Cursor, Write};

use office_oxide::{Document, DocumentFormat};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn row(r: u32, a: &str, b: &str) -> String {
    format!(
        r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{a}</t></is></c><c r="B{r}" t="inlineStr"><is><t>{b}</t></is></c></row>"#
    )
}

fn sheet(body: &str) -> String {
    format!(r#"<worksheet xmlns="{SML}"><sheetData>{body}</sheetData></worksheet>"#)
}

fn workbook(sheets: &[(&str, String)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default();
    let mut put = |name: &str, body: &str| {
        zip.start_file(name, o).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    };
    let overrides: String = (1..=sheets.len())
        .map(|i| format!(r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#))
        .collect();
    put(
        "[Content_Types].xml",
        &format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{overrides}</Types>"#
        ),
    );
    put(
        "_rels/.rels",
        &format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        ),
    );
    let names: String = sheets
        .iter()
        .enumerate()
        .map(|(i, (n, _))| format!(r#"<sheet name="{n}" sheetId="{0}" r:id="rId{0}"/>"#, i + 1))
        .collect();
    put(
        "xl/workbook.xml",
        &format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets>{names}</sheets></workbook>"#),
    );
    let rels: String = (1..=sheets.len()).map(|i| format!(r#"<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>"#)).collect();
    put(
        "xl/_rels/workbook.xml.rels",
        &format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        ),
    );
    for (i, (_, xml)) in sheets.iter().enumerate() {
        put(&format!("xl/worksheets/sheet{}.xml", i + 1), xml);
    }
    zip.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

/// Sheet "Cut": rows 1 and 2 are fine and row 3 has a damaged attribute.
fn damaged_workbook() -> Vec<u8> {
    let damaged = format!(
        "{}{}{}",
        row(1, "alpha", "one"),
        row(2, "bravo", "two"),
        r#"<row r="3"><c r="A3" t=inlineStr"><is><t>charlie</t></is></c></row>"#
    );
    workbook(&[
        ("Cut", sheet(&damaged)),
        ("Fine", sheet(&row(1, "delta", "four"))),
    ])
}

#[test]
fn test_the_rows_before_the_damage_are_kept() {
    let doc = open(damaged_workbook());
    let text = doc.plain_text();
    for kept in ["alpha", "one", "bravo", "two", "delta", "four"] {
        assert!(text.contains(kept), "{kept}: {text}");
    }
    assert!(!text.contains("charlie"), "{text}");
    // The loss is said in the text, not only in the metadata.
    assert!(text.contains("\"Cut\"") && text.contains("malformed"), "{text}");
    assert!(!text.contains("unreadable sheet"), "{text}");
}

#[test]
fn test_every_surface_says_the_sheet_was_cut_short() {
    let doc = open(damaged_workbook());
    let md = doc.to_markdown();
    assert!(md.contains("bravo") && md.contains("malformed"), "{md}");
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("Cut") && w.contains("malformed")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(doc.to_html().contains("bravo"));
}

#[test]
fn test_a_sheet_damaged_before_its_first_row_is_still_unreadable() {
    // One sheet beside a damaged one, so the workbook opens: the damaged sheet
    // has no complete row, so there is nothing to keep, and it is reported as
    // unreadable, as before.
    let bad = workbook(&[
        (
            "Gone",
            sheet(r#"<row r="1"><c r="A1" t=inlineStr"><is><t>x</t></is></c></row>"#),
        ),
        ("Fine", sheet(&row(1, "kept", "also"))),
    ]);
    let doc = open(bad);
    let text = doc.plain_text();
    assert!(text.contains("unreadable sheet \"Gone\""), "{text}");
    assert!(!text.contains("cut short"), "{text}");
    assert!(text.contains("kept") && text.contains("also"), "{text}");
}

/// The cut falls between rows: a row the damage reaches is dropped whole, never
/// kept as half a row.
#[test]
fn test_a_row_the_damage_reaches_is_dropped_whole() {
    let damaged = format!(
        "{}{}",
        row(1, "kept-a", "kept-b"),
        r#"<row r="2"><c r="A2" t="inlineStr"><is><t>half-a</t></is></c><c r="B2" t=inlineStr"><is><t>half-b</t></is></c></row>"#
    );
    let text = open(workbook(&[("S", sheet(&damaged))])).plain_text();
    assert!(text.contains("kept-a") && text.contains("kept-b"), "{text}");
    assert!(!text.contains("half-a") && !text.contains("half-b"), "{text}");
}

/// A well-formed sheet is untouched: no notice, no flag.
#[test]
fn test_a_sound_sheet_carries_no_damage_notice() {
    let doc = open(workbook(&[("S", sheet(&row(1, "x", "y")))]));
    assert!(!doc.plain_text().contains("malformed"));
    assert!(!doc.to_ir().metadata.text_truncated);
}

/// A prefixed worksheet (`x:worksheet`) is salvaged too.
#[test]
fn test_a_prefixed_worksheet_is_salvaged() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{SML}"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>prefixed</x:t></x:is></x:c></x:row><x:row r="2"><x:c r="A2" t=inlineStr"></x:c></x:row></x:sheetData></x:worksheet>"#
    );
    let text = open(workbook(&[("P", xml)])).plain_text();
    assert!(text.contains("prefixed"), "{text}");
}
