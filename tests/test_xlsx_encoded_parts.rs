//! A part stored with a compression method this reader cannot decode (or
//! encrypted) is damaged, not absent: it is reported, and the workbook is
//! never presented as if the part had not been there.

use std::io::{Cursor, Write};

use office_oxide::{Document, DocumentFormat};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PERSON_REL: &str = "http://schemas.microsoft.com/office/2017/10/relationships/person";

/// A one-sheet workbook with a shared string, a stylesheet, a theme and a
/// persons part, every part stored uncompressed.
fn workbook() -> Vec<u8> {
    let parts: Vec<(&str, String)> = vec![
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.into(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{R}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="{R}/styles" Target="styles.xml"/><Relationship Id="rId4" Type="{R}/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="{PERSON_REL}" Target="persons/person.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#),
        ),
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{NS}" count="1" uniqueCount="1"><si><t>HELLO</t></si></sst>"#),
        ),
        (
            "xl/styles.xml",
            format!(r#"<styleSheet xmlns="{NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#),
        ),
        (
            "xl/theme/theme1.xml",
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"/>"#.into(),
        ),
        (
            "xl/persons/person.xml",
            r#"<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>"#.into(),
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts: zip::write::FileOptions<'_, '_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// `bytes` with the entry `name` marked as stored by compression method 99
/// (AES encryption), in its local and its central directory header.
fn with_method_99(mut bytes: Vec<u8>, name: &str) -> Vec<u8> {
    let find = |bytes: &[u8], sig: &[u8], header: usize| -> usize {
        (0..bytes.len() - header)
            .find(|&i| {
                bytes[i..].starts_with(sig) && bytes[i + header..].starts_with(name.as_bytes())
            })
            .unwrap()
    };
    let local = find(&bytes, b"PK\x03\x04", 30);
    bytes[local + 8..local + 10].copy_from_slice(&99u16.to_le_bytes());
    let central = find(&bytes, b"PK\x01\x02", 46);
    bytes[central + 10..central + 12].copy_from_slice(&99u16.to_le_bytes());
    bytes
}

#[test]
fn test_an_intact_workbook_reports_nothing() {
    let doc = Document::from_reader(Cursor::new(workbook()), DocumentFormat::Xlsx).unwrap();
    let xlsx = doc.as_xlsx().unwrap();
    assert!(xlsx.unreadable_parts.is_empty(), "{:?}", xlsx.unreadable_parts);
    assert!(doc.plain_text().contains("HELLO"));
}

#[test]
#[ignore = "needs the typed zip error from the core change"]
fn test_a_part_with_an_unsupported_compression_method_is_reported_not_taken_for_absent() {
    for part in [
        "xl/sharedStrings.xml",
        "xl/styles.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/theme/theme1.xml",
        "xl/persons/person.xml",
    ] {
        let bytes = with_method_99(workbook(), part);
        // Opening may refuse the workbook outright; if it opens, the part is
        // on record and the output says it is incomplete.
        let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx) else {
            continue;
        };
        let xlsx = doc.as_xlsx().unwrap();
        assert!(
            xlsx.unreadable_parts.iter().any(|(p, _)| p == part),
            "{part} is not reported: {:?}",
            xlsx.unreadable_parts
        );
        assert!(doc.to_ir().metadata.text_truncated, "{part}");
        assert!(doc.plain_text().contains(part), "{part} is not in the text");
    }
}
