//! A properties part that exists but cannot be read is reported, not taken
//! for an absent one: the document opens and reads as before, and the
//! warnings say which part was lost.

use std::io::{Cursor, Write};

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn docx(core: Option<&str>) -> Vec<u8> {
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let body = format!(
        r#"<w:document xmlns:w="{WML}"><w:body><w:p><w:r><w:t>body text</w:t></w:r></w:p></w:body></w:document>"#
    );
    writer
        .add_part(
            &PartName::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            body.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    if let Some(core) = core {
        writer
            .add_part(
                &PartName::new("/docProps/core.xml").unwrap(),
                "application/vnd.openxmlformats-package.core-properties+xml",
                core.as_bytes(),
            )
            .unwrap();
        writer.add_package_rel(rel_types::CORE_PROPERTIES, "docProps/core.xml");
    }
    writer.finish().unwrap().into_inner()
}

fn warnings_of(bytes: Vec<u8>, format: DocumentFormat) -> (String, Vec<String>) {
    let doc = Document::from_reader(Cursor::new(bytes), format).unwrap();
    (doc.plain_text(), doc.to_ir().metadata.warnings)
}

#[test]
fn test_a_damaged_core_properties_part_is_reported_in_a_docx() {
    let (text, warnings) = warnings_of(docx(Some("<cp:coreProperties><<")), DocumentFormat::Docx);
    assert!(text.contains("body text"));
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("docProps/core.xml") && w.contains("could not be read")),
        "{warnings:?}"
    );
}

#[test]
fn test_sound_or_absent_properties_report_nothing() {
    let core = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></cp:coreProperties>"#;
    for bytes in [docx(Some(core)), docx(None)] {
        let (_, warnings) = warnings_of(bytes, DocumentFormat::Docx);
        assert!(warnings.is_empty(), "{warnings:?}");
    }
}

fn xlsx(rels: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default();
    let mut put = |name: &str, body: &str| {
        zip.start_file(name, o).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    };
    put(
        "[Content_Types].xml",
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
    );
    put("_rels/.rels", rels);
    put(
        "xl/workbook.xml",
        &format!(
            r#"<workbook xmlns="{SML}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        ),
    );
    put(
        "xl/_rels/workbook.xml.rels",
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
    );
    put(
        "xl/worksheets/sheet1.xml",
        &format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell text</t></is></c></row></sheetData></worksheet>"#
        ),
    );
    zip.finish().unwrap().into_inner()
}

#[test]
fn test_damaged_package_relationships_of_a_workbook_are_reported() {
    let (text, warnings) = warnings_of(xlsx("<Relationships><<"), DocumentFormat::Xlsx);
    assert!(text.contains("cell text"), "{text}");
    assert!(warnings.iter().any(|w| w.contains("_rels/.rels")), "{warnings:?}");
    let sound = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let (_, warnings) = warnings_of(xlsx(sound), DocumentFormat::Xlsx);
    assert!(warnings.is_empty(), "{warnings:?}");
}
