//! A relationships part that cannot be read costs its sheet the hyperlinks,
//! drawings, comments and tables it names. That is reported (a warning naming
//! the part and the reason, and `text_truncated`), never an empty set of
//! relationships passing for a sheet that had none. A hyperlink target that
//! the package layer accepts is kept, whatever its length: the same limit
//! applies to every format.

use std::io::{Cursor, Write};

use office_oxide::{Document, DocumentFormat};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// One sheet of `cells` linked cells whose relationship targets `target`,
/// plus `extra` parts.
fn workbook(
    cells: usize,
    target: &str,
    sheet_rels_extra: &str,
    extra: &[(&str, String)],
) -> Vec<u8> {
    let mut sheet = format!(r#"<worksheet xmlns="{NS}" xmlns:r="{R}"><sheetData>"#);
    for i in 1..=cells {
        sheet.push_str(&format!(r#"<row r="{i}"><c r="A{i}"><v>{i}</v></c></row>"#));
    }
    sheet.push_str("</sheetData><hyperlinks>");
    for i in 1..=cells {
        sheet.push_str(&format!(r#"<hyperlink ref="A{i}" r:id="rId1"/>"#));
    }
    sheet.push_str("</hyperlinks>");
    if !sheet_rels_extra.is_empty() {
        sheet.push_str(r#"<drawing r:id="rId2"/>"#);
    }
    sheet.push_str("</worksheet>");
    let mut parts: Vec<(String, String)> = vec![
        (
            "[Content_Types].xml".into(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into(),
        ),
        (
            "_rels/.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml".into(),
            format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
        ),
        ("xl/worksheets/sheet1.xml".into(), sheet),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/hyperlink" Target="http://{target}" TargetMode="External"/>{sheet_rels_extra}</Relationships>"#),
        ),
    ];
    for (n, d) in extra {
        parts.push(((*n).into(), d.clone()));
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> = zip::write::FileOptions::default();
    for (n, d) in parts {
        zip.start_file(n.as_str(), opts).unwrap();
        zip.write_all(d.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap()
}

/// Links in the IR, and the bytes of their targets.
fn links(ir: &office_oxide::ir::DocumentIR) -> usize {
    serde_json::to_string(ir)
        .unwrap()
        .matches("\"hyperlink\":\"http")
        .count()
}

#[test]
fn test_a_hyperlink_target_past_two_thousand_characters_is_kept() {
    // Signed storage and sharing links run to several thousand characters.
    let doc = open(workbook(5, &"u".repeat(5_000), "", &[]));
    let ir = doc.to_ir();
    assert_eq!(links(&ir), 5);
    assert!(!ir.metadata.text_truncated, "{:?}", ir.metadata.warnings);
    assert!(ir.metadata.warnings.is_empty(), "{:?}", ir.metadata.warnings);
}

#[test]
fn test_a_sheet_relationships_part_that_cannot_be_read_is_reported() {
    // A target past the package limit (64 KiB) refuses the whole part.
    let doc = open(workbook(5, &"u".repeat(100_000), "", &[]));
    let ir = doc.to_ir();
    assert_eq!(links(&ir), 0);
    assert!(ir.metadata.text_truncated, "the loss is not flagged");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("xl/worksheets/_rels/sheet1.xml.rels") && w.contains("65536")),
        "no warning naming the part and the reason: {:?}",
        ir.metadata.warnings
    );
    // The cells stay.
    assert!(doc.plain_text().contains('5'));
}

#[test]
fn test_a_drawing_relationships_part_that_cannot_be_read_is_reported() {
    let drawing = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;
    let rels = format!(
        r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/image" Target="{}"/></Relationships>"#,
        "i".repeat(100_000)
    );
    let doc = open(workbook(
        1,
        "x",
        &format!(
            r#"<Relationship Id="rId2" Type="{R}/drawing" Target="../drawings/drawing1.xml"/>"#
        ),
        &[
            ("xl/drawings/drawing1.xml", drawing.to_string()),
            ("xl/drawings/_rels/drawing1.xml.rels", rels),
        ],
    ));
    let ir = doc.to_ir();
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("xl/drawings/_rels/drawing1.xml.rels")),
        "{:?}",
        ir.metadata.warnings
    );
}
