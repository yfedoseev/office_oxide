//! Document properties end to end: OOXML core/app/custom parts, the
//! package thumbnail and signature, the legacy SummaryInformation /
//! DocumentSummaryInformation property sets — read into `ir::Metadata`
//! and written back by every OOXML writer.

#[path = "common/ppt_builder.rs"]
mod ppt_builder;

use std::io::{Cursor, Read};

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::create::create_from_ir_to_writer;
use office_oxide::ir::{CustomProperty, ImageFormat, Metadata};
use office_oxide::{Document, DocumentFormat};

const CORE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Annual Report</dc:title>
  <dc:creator>Author One</dc:creator>
  <dc:language>en-GB</dc:language>
  <cp:keywords>annual report; budget</cp:keywords>
  <cp:category>Finance</cp:category>
  <cp:contentStatus>Final</cp:contentStatus>
  <cp:lastModifiedBy>Editor Two</cp:lastModifiedBy>
  <cp:revision>7</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>
</cp:coreProperties>"#;

const APP_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><Company>AT&amp;T Labs</Company><Manager>Boss &amp; Co</Manager></Properties>"#;

const CUSTOM_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Acme &amp; Sons</vt:lpwstr></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Pages"><vt:i4>12</vt:i4></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Approved"><vt:bool>true</vt:bool></property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="List"><vt:vector size="1" baseType="lpwstr"><vt:lpwstr>x</vt:lpwstr></vt:vector></property>
</Properties>"#;

const THUMBNAIL: &[u8] = b"\xFF\xD8\xFF\xE0 not really a jpeg";

/// A DOCX with every package-level property part, the property parts at
/// non-conventional paths, a thumbnail and a signature origin.
fn docx_with_properties() -> Vec<u8> {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc = PartName::new("/word/document.xml").unwrap();
    w.add_part(
        &doc,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>"#,
    )
    .unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let parts: [(&str, &str, &str, &[u8]); 5] = [
        (
            rel_types::CORE_PROPERTIES,
            "meta/core-props.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            CORE_XML.as_bytes(),
        ),
        (
            rel_types::EXTENDED_PROPERTIES,
            "meta/app-props.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            APP_XML.as_bytes(),
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "meta/custom-props.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            CUSTOM_XML.as_bytes(),
        ),
        (rel_types::THUMBNAIL, "docProps/thumbnail.jpeg", "image/jpeg", THUMBNAIL),
        (
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            "_xmlsignatures/origin.sigs",
            "application/vnd.openxmlformats-package.digital-signature-origin",
            b"",
        ),
    ];
    for (rel, path, ct, data) in parts {
        w.add_part(&PartName::new(&format!("/{path}")).unwrap(), ct, data)
            .unwrap();
        w.add_package_rel(rel, path);
    }
    w.finish().unwrap().into_inner()
}

fn assert_full_metadata(m: &Metadata, what: &str) {
    assert_eq!(m.title.as_deref(), Some("Annual Report"), "{what}");
    assert_eq!(m.author.as_deref(), Some("Author One"), "{what}");
    assert_eq!(m.keywords, ["annual report", "budget"], "{what}: whitespace is not a separator");
    assert_eq!(m.category.as_deref(), Some("Finance"), "{what}");
    assert_eq!(m.content_status.as_deref(), Some("Final"), "{what}");
    assert_eq!(m.language.as_deref(), Some("en-GB"), "{what}");
    assert_eq!(m.last_modified_by.as_deref(), Some("Editor Two"), "{what}");
    assert_eq!(m.revision.as_deref(), Some("7"), "{what}");
    assert_eq!(m.company.as_deref(), Some("AT&T Labs"), "{what}");
    assert_eq!(m.manager.as_deref(), Some("Boss & Co"), "{what}");
    let custom: Vec<(&str, &str, &str)> = m
        .custom_properties
        .iter()
        .map(|p| (p.name.as_str(), p.value.as_str(), p.value_type.as_str()))
        .collect();
    assert_eq!(
        custom,
        [
            ("Client", "Acme & Sons", "lpwstr"),
            ("Pages", "12", "i4"),
            ("Approved", "true", "bool"),
        ],
        "{what}: scalar custom properties, the vector one skipped"
    );
}

#[test]
fn test_docx_properties_reach_the_ir() {
    let doc =
        Document::from_reader(Cursor::new(docx_with_properties()), DocumentFormat::Docx).unwrap();
    let m = doc.to_ir().metadata;
    assert_full_metadata(&m, "docx read");
    assert!(m.has_digital_signature);
    let thumb = m.thumbnail.expect("thumbnail surfaced");
    assert_eq!(thumb.format, Some(ImageFormat::Jpeg));
    assert_eq!(thumb.data.as_deref(), Some(THUMBNAIL));
}

/// Every OOXML writer writes core.xml with all its properties, app.xml
/// naming the producer plus company/manager, and custom.xml — so the IR
/// round trip keeps them.
#[test]
fn test_every_writer_writes_core_app_and_custom_properties() {
    let doc =
        Document::from_reader(Cursor::new(docx_with_properties()), DocumentFormat::Docx).unwrap();
    let ir = doc.to_ir();
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ] {
        let mut out = Cursor::new(Vec::new());
        create_from_ir_to_writer(&ir, format, &mut out).unwrap();
        let bytes = out.into_inner();

        let mut zip = zip::ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
        let mut app = String::new();
        zip.by_name("docProps/app.xml")
            .unwrap_or_else(|_| panic!("{format:?}: app.xml written"))
            .read_to_string(&mut app)
            .unwrap();
        assert!(app.contains("<Application>office_oxide</Application>"), "{format:?}: {app}");

        let back = Document::from_reader(Cursor::new(bytes), format).unwrap();
        let m = back.to_ir().metadata;
        assert_full_metadata(&m, &format!("{format:?} round trip"));
        // A re-written package is not signed, and the thumbnail (a
        // rendering of the source layout) is not carried over.
        assert!(!m.has_digital_signature);
    }
}

/// `normalize_w3cdtf` used to run only on a write path nothing called;
/// the real writers copied a malformed source date verbatim.
#[test]
fn test_writers_normalise_or_drop_malformed_dates() {
    let ir = office_oxide::ir::DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            created: Some("2021- 9- 3T20:25:22Z".into()),
            modified: Some("2015sss-06-20T07:40:00Z".into()),
            ..Default::default()
        },
        sections: Vec::new(),
        defined_names: Vec::new(),
    };
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
    ] {
        let mut out = Cursor::new(Vec::new());
        create_from_ir_to_writer(&ir, format, &mut out).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(out.into_inner())).unwrap();
        let mut core = String::new();
        zip.by_name("docProps/core.xml")
            .unwrap()
            .read_to_string(&mut core)
            .unwrap();
        assert!(core.contains(">2021-09-03T20:25:22Z<"), "{format:?}: {core}");
        assert!(!core.contains("dcterms:modified"), "{format:?}: {core}");
    }
}

#[test]
fn test_custom_property_values_are_written_back_with_their_type() {
    let ir = office_oxide::ir::DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            custom_properties: vec![
                CustomProperty {
                    name: "Ratio <x>".into(),
                    value: "0.5".into(),
                    value_type: "r8".into(),
                },
                CustomProperty {
                    name: "Odd".into(),
                    value: "v".into(),
                    value_type: "not-a-variant".into(),
                },
            ],
            ..Default::default()
        },
        sections: Vec::new(),
        defined_names: Vec::new(),
    };
    let mut out = Cursor::new(Vec::new());
    create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut out).unwrap();
    let back = Document::from_reader(Cursor::new(out.into_inner()), DocumentFormat::Docx).unwrap();
    let custom = back.to_ir().metadata.custom_properties;
    assert_eq!(custom.len(), 2);
    assert_eq!((custom[0].name.as_str(), custom[0].value_type.as_str()), ("Ratio <x>", "r8"));
    // An unknown variant type is written as a string rather than as an
    // element no reader accepts.
    assert_eq!(custom[1].value_type, "lpwstr");
}

/// The XLSX fast path read docProps only at the conventional paths; a
/// package whose relationships put them elsewhere lost all metadata.
#[test]
fn test_xlsx_finds_properties_through_relationships() {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let wb = PartName::new("/xl/workbook.xml").unwrap();
    w.add_part(
        &wb,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
    )
    .unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "xl/workbook.xml");
    w.add_part_rel(&wb, rel_types::WORKSHEET, "worksheets/sheet1.xml");
    w.add_part(
        &PartName::new("/xl/worksheets/sheet1.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell</t></is></c></row></sheetData></worksheet>"#,
    )
    .unwrap();
    for (rel, path, ct, data) in [
        (
            rel_types::CORE_PROPERTIES,
            "meta/core-props.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            CORE_XML,
        ),
        (
            rel_types::EXTENDED_PROPERTIES,
            "meta/app-props.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            APP_XML,
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "meta/custom-props.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            CUSTOM_XML,
        ),
    ] {
        w.add_part(&PartName::new(&format!("/{path}")).unwrap(), ct, data.as_bytes())
            .unwrap();
        w.add_package_rel(rel, path);
    }
    let bytes = w.finish().unwrap().into_inner();
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xlsx).unwrap();
    assert!(doc.plain_text().contains("cell"));
    assert_full_metadata(&doc.to_ir().metadata, "xlsx read");
}

/// `Metadata::properties` is what `office-oxide info` prints: every set
/// property, author/dates/subject/keywords included, in a fixed order.
#[test]
fn test_metadata_summary_lists_every_set_property() {
    let doc =
        Document::from_reader(Cursor::new(docx_with_properties()), DocumentFormat::Docx).unwrap();
    let props = doc.to_ir().metadata.properties();
    let labels: Vec<&str> = props.iter().map(|(l, _)| *l).collect();
    assert_eq!(
        labels,
        [
            "Title",
            "Author",
            "Keywords",
            "Category",
            "Company",
            "Manager",
            "Created",
            "Last modified by",
            "Revision",
            "Content status",
            "Language",
        ]
    );
    assert_eq!(props[2].1, "annual report, budget");
}

// ---------------------------------------------------------------------------
// Legacy property sets
// ---------------------------------------------------------------------------

/// A `VT_LPSTR` TypedPropertyValue ([MS-OLEPS] §2.15, §2.5 CodePageString).
fn lpstr(s: &str) -> Vec<u8> {
    let mut v = 0x001Eu16.to_le_bytes().to_vec();
    v.extend_from_slice(&0u16.to_le_bytes());
    v.extend_from_slice(&((s.len() + 1) as u32).to_le_bytes());
    v.extend_from_slice(s.as_bytes());
    v.push(0);
    while !v.len().is_multiple_of(4) {
        v.push(0);
    }
    v
}

/// A one-property-set PropertySetStream ([MS-OLEPS] §2.21) holding
/// `props` as `(property id, TypedPropertyValue)`.
fn property_set_stream(props: &[(u32, Vec<u8>)]) -> Vec<u8> {
    let mut out = 0xFFFEu16.to_le_bytes().to_vec();
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&[0u8; 16]);
    out.extend_from_slice(&1u32.to_le_bytes());
    out.extend_from_slice(&[0u8; 16]); // FMTID
    out.extend_from_slice(&48u32.to_le_bytes()); // set offset
    let base = out.len();
    out.extend_from_slice(&0u32.to_le_bytes()); // size, patched below
    out.extend_from_slice(&(props.len() as u32).to_le_bytes());
    let mut at = 8 + props.len() * 8;
    for (id, value) in props {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(at as u32).to_le_bytes());
        at += value.len();
    }
    for (_, value) in props {
        out.extend_from_slice(value);
    }
    let size = (out.len() - base) as u32;
    out[base..base + 4].copy_from_slice(&size.to_le_bytes());
    out
}

#[test]
fn test_legacy_document_summary_information_reaches_the_ir() {
    // SummaryInformation: title (2), last author (8), revision (9).
    let si = property_set_stream(&[
        (2, lpstr("Legacy Deck")),
        (8, lpstr("Saver")),
        (9, lpstr("4")),
    ]);
    // DocumentSummaryInformation: category (2), manager (14), company (15).
    let dsi = property_set_stream(&[
        (0x02, lpstr("Sales")),
        (0x0E, lpstr("Head Office")),
        (0x0F, lpstr("Contoso Ltd")),
    ]);
    let bytes = ppt_builder::build_ppt(
        &["Slide"],
        &[
            ("\u{5}SummaryInformation", &si),
            ("\u{5}DocumentSummaryInformation", &dsi),
            ("_xmlsignatures/", &[]),
        ],
    );
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt).unwrap();
    let m = doc.to_ir().metadata;
    assert_eq!(m.title.as_deref(), Some("Legacy Deck"));
    assert_eq!(m.last_modified_by.as_deref(), Some("Saver"));
    assert_eq!(m.revision.as_deref(), Some("4"));
    assert_eq!(m.category.as_deref(), Some("Sales"));
    assert_eq!(m.manager.as_deref(), Some("Head Office"));
    assert_eq!(m.company.as_deref(), Some("Contoso Ltd"));
    assert!(m.has_digital_signature, "an _xmlsignatures storage marks a signed file");

    let unsigned = ppt_builder::build_ppt(&["Slide"], &[("\u{5}SummaryInformation", &si)]);
    let doc = Document::from_reader(Cursor::new(unsigned), DocumentFormat::Ppt).unwrap();
    let m = doc.to_ir().metadata;
    assert!(!m.has_digital_signature);
    assert_eq!(m.company, None);
}
