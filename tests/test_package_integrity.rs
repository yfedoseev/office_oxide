//! Zip entries whose bytes fail their recorded CRC-32.
//!
//! Refusing such a part cost real documents their content: a workbook whose
//! sheets carried a bad checksum went from its full text to "unreadable
//! sheet" for every one of them. The bytes are read as stored — as 7-Zip
//! does when it extracts with a CRC warning — and the mismatch is reported
//! through `Metadata::warnings` so possibly damaged content is not
//! presented as sound.

use std::io::{Cursor, Write};

use office_oxide::{Document, DocumentFormat};

/// A zip of stored (uncompressed) entries, so a byte of an entry's content
/// can be changed in place to break its CRC-32.
fn stored_zip(parts: &[(&str, String)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in parts {
        zip.start_file(*name, stored).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Change the first byte of `marker` in `data`.
fn corrupt(data: &mut [u8], marker: &[u8]) {
    let at = data
        .windows(marker.len())
        .position(|w| w == marker)
        .expect("marker present in a stored entry");
    data[at] = b'Z';
}

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn sheet(text: &str) -> String {
    format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row></sheetData></worksheet>"#
    )
}

fn workbook_with_two_sheets() -> Vec<u8> {
    stored_zip(&[
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_string(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#),
        ),
        ("xl/worksheets/sheet1.xml", sheet("alpha")),
        ("xl/worksheets/sheet2.xml", sheet("bravo-marker")),
    ])
}

#[test]
fn test_a_sheet_failing_its_crc_keeps_its_own_cells_and_is_reported() {
    let mut data = workbook_with_two_sheets();
    // "bravo-marker" -> "Zravo-marker": the stored bytes change, the
    // recorded CRC-32 does not.
    corrupt(&mut data, b"bravo-marker");

    let doc = Document::from_reader(Cursor::new(data), DocumentFormat::Xlsx).unwrap();
    let text = doc.plain_text();
    assert!(text.contains("alpha"), "{text}");
    // The damaged sheet's own cells — not the first sheet's, re-read under
    // the second sheet's name by an index-based fallback.
    assert!(text.contains("Zravo-marker"), "{text}");
    assert_eq!(text.matches("alpha").count(), 1, "{text}");

    let meta = doc.to_ir().metadata;
    assert!(
        meta.warnings
            .iter()
            .any(|w| w.contains("xl/worksheets/sheet2.xml") && w.contains("CRC-32")),
        "{:?}",
        meta.warnings
    );
}

#[test]
fn test_an_intact_package_carries_no_integrity_warning() {
    let doc = Document::from_reader(Cursor::new(workbook_with_two_sheets()), DocumentFormat::Xlsx)
        .unwrap();
    assert!(doc.to_ir().metadata.warnings.is_empty());
}
