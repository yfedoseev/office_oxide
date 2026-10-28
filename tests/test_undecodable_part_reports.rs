//! A part the zip cannot decode (a compression method the reader does not
//! implement, or an encrypted entry) is a damaged part, not an absent one:
//! every reader reports it and the text says it is incomplete. Before, such a
//! part looked absent, so a workbook whose shared strings could not be
//! decoded read as a workbook of empty cells with nothing to say so.
//!
//! Each test builds a small package in code, then rewrites one entry's local
//! and central headers (ZIP appnote 4.4.4 general purpose flag, 4.4.5
//! compression method) so that the entry cannot be decoded.

use std::io::{Cursor, Write};

use office_oxide::{Document, DocumentFormat};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

/// How an entry is made undecodable.
#[derive(Clone, Copy, Debug)]
enum Damage {
    /// Compression method 99 (AES marker): no decoder for it here.
    UnsupportedMethod,
    /// General purpose flag bit 0 (entry is encrypted).
    Encrypted,
}

const DAMAGES: [Damage; 2] = [Damage::UnsupportedMethod, Damage::Encrypted];

fn zip_of(parts: &[(&str, String)]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in parts {
        zip.start_file(*name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn u16_at(b: &[u8], at: usize) -> usize {
    u16::from_le_bytes([b[at], b[at + 1]]) as usize
}

/// `bytes` with the entry called `name` damaged in both its headers.
fn damage(mut bytes: Vec<u8>, name: &str, how: Damage) -> Vec<u8> {
    let mut patched = 0;
    for at in 0..bytes.len().saturating_sub(46) {
        let (flags, method, name_at, name_len) = match &bytes[at..at + 4] {
            // Local file header: flags at 6, method at 8, name at 30.
            [0x50, 0x4b, 0x03, 0x04] => (at + 6, at + 8, at + 30, u16_at(&bytes, at + 26)),
            // Central directory header: flags at 8, method at 10, name at 46.
            [0x50, 0x4b, 0x01, 0x02] => (at + 8, at + 10, at + 46, u16_at(&bytes, at + 28)),
            _ => continue,
        };
        if bytes.get(name_at..name_at + name_len) != Some(name.as_bytes()) {
            continue;
        }
        match how {
            Damage::UnsupportedMethod => {
                bytes[method..method + 2].copy_from_slice(&99u16.to_le_bytes())
            },
            Damage::Encrypted => bytes[flags] |= 1,
        }
        patched += 1;
    }
    assert_eq!(patched, 2, "{name} has a local and a central header");
    bytes
}

fn rels(entries: &str) -> String {
    format!(r#"<Relationships xmlns="{PKG_REL}">{entries}</Relationships>"#)
}

fn root_rels(target: &str) -> String {
    rels(&format!(
        r#"<Relationship Id="rId1" Type="{R}/officeDocument" Target="{target}"/>"#
    ))
}

fn types(overrides: &str) -> String {
    format!(
        r#"<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>{overrides}</Types>"#
    )
}

// ---- xlsx -----------------------------------------------------------------

fn xlsx() -> Vec<u8> {
    zip_of(&[
        (
            "[Content_Types].xml",
            types(
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
            ),
        ),
        ("_rels/.rels", root_rels("xl/workbook.xml")),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels(&format!(
                r#"<Relationship Id="rId1" Type="{R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{R}/sharedStrings" Target="sharedStrings.xml"/>"#
            )),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#
            ),
        ),
        (
            "xl/sharedStrings.xml",
            format!(r#"<sst xmlns="{NS}" count="1" uniqueCount="1"><si><t>HELLO</t></si></sst>"#),
        ),
        (
            "xl/styles.xml",
            format!(
                r#"<styleSheet xmlns="{NS}"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#
            ),
        ),
        ("xl/theme/theme1.xml", format!(r#"<a:theme xmlns:a="{DML}" name="t"/>"#)),
    ])
}

fn open(bytes: Vec<u8>, format: DocumentFormat) -> Document {
    Document::from_reader(Cursor::new(bytes), format).expect("opens, with the loss reported")
}

fn reported(doc: &Document) -> Vec<String> {
    let pairs = |parts: &[(String, String)]| {
        parts
            .iter()
            .map(|(p, why)| format!("{p}: {why}"))
            .collect::<Vec<_>>()
    };
    if let Some(d) = doc.as_xlsx() {
        pairs(&d.unreadable_parts)
    } else if let Some(d) = doc.as_docx() {
        d.unreadable_parts.to_vec()
    } else {
        pairs(&doc.as_pptx().unwrap().unreadable_parts)
    }
}

/// `part` is among the `found` reports, and its reason is not "missing".
fn assert_reported_damaged(found: &[String], part: &str, ctx: &str) {
    let entry = found.iter().find(|f| f.contains(part));
    let entry = entry.unwrap_or_else(|| panic!("{ctx}: {part} is not reported: {found:?}"));
    assert!(!entry.contains("missing"), "{ctx}: reported as absent: {entry}");
}

#[test]
fn test_a_workbook_part_that_cannot_be_decoded_is_reported_not_taken_for_absent() {
    for part in [
        "xl/sharedStrings.xml",
        "xl/styles.xml",
        "xl/_rels/workbook.xml.rels",
    ] {
        for how in DAMAGES {
            let doc = open(damage(xlsx(), part, how), DocumentFormat::Xlsx);
            let found = reported(&doc);
            assert_reported_damaged(&found, part, &format!("{how:?}"));
            assert!(doc.plain_text().contains(part), "{part} ({how:?}) is not in the text");
            assert!(doc.to_ir().metadata.text_truncated, "{part} ({how:?})");
        }
    }
}

#[test]
fn test_a_part_missing_from_the_package_is_still_absent_not_damaged() {
    // The same workbook without its shared strings, styles and theme.
    let mut parts = vec![];
    let full = xlsx();
    let mut archive = zip::ZipArchive::new(Cursor::new(full)).unwrap();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        if name.contains("sharedStrings") || name.contains("styles") || name.contains("theme") {
            continue;
        }
        let mut body = String::new();
        std::io::Read::read_to_string(&mut f, &mut body).unwrap();
        parts.push((name, body));
    }
    let parts: Vec<(&str, String)> = parts.iter().map(|(n, b)| (n.as_str(), b.clone())).collect();
    let doc = open(zip_of(&parts), DocumentFormat::Xlsx);
    assert!(reported(&doc).is_empty(), "{:?}", reported(&doc));
    assert!(!doc.to_ir().metadata.text_truncated);
}

// ---- docx -----------------------------------------------------------------

fn docx() -> Vec<u8> {
    zip_of(&[
        (
            "[Content_Types].xml",
            types(
                r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#,
            ),
        ),
        ("_rels/.rels", root_rels("word/document.xml")),
        (
            "word/_rels/document.xml.rels",
            rels(&format!(r#"<Relationship Id="rId2" Type="{R}/header" Target="header1.xml"/>"#)),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document xmlns:w="{WML}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>"#
            ),
        ),
        (
            "word/header1.xml",
            format!(r#"<w:hdr xmlns:w="{WML}"><w:p><w:r><w:t>HEAD</w:t></w:r></w:p></w:hdr>"#),
        ),
    ])
}

#[test]
fn test_a_docx_part_that_cannot_be_decoded_is_reported() {
    for how in DAMAGES {
        let doc = open(damage(docx(), "word/header1.xml", how), DocumentFormat::Docx);
        let found = reported(&doc);
        assert_reported_damaged(&found, "header1.xml", &format!("{how:?}"));
        let ir = doc.to_ir();
        assert!(ir.metadata.text_truncated, "{how:?}");
        assert!(
            ir.metadata
                .warnings
                .iter()
                .any(|w| w.contains("header1.xml")),
            "{how:?}: {:?}",
            ir.metadata.warnings
        );
    }
}

#[test]
fn test_a_docx_relationships_part_that_cannot_be_decoded_still_opens_with_the_loss_reported() {
    // The document opened when such a part was taken for absent; refusing it
    // now would turn a readable document into an error. It opens, with no
    // relationships, and says so.
    for how in DAMAGES {
        let bytes = damage(docx(), "word/_rels/document.xml.rels", how);
        let doc = open(bytes, DocumentFormat::Docx);
        assert_reported_damaged(&reported(&doc), "document.xml.rels", &format!("{how:?}"));
        assert!(doc.plain_text().contains("body"), "{how:?}");
        assert!(doc.to_ir().metadata.text_truncated, "{how:?}");
    }
}

#[test]
fn test_a_package_relationships_part_that_cannot_be_decoded_still_opens_with_the_loss_reported() {
    // The main part is found through `[Content_Types].xml` all the same.
    for how in DAMAGES {
        let doc = open(damage(docx(), "_rels/.rels", how), DocumentFormat::Docx);
        assert_reported_damaged(&reported(&doc), ".rels", &format!("{how:?}"));
        assert!(doc.plain_text().contains("body"), "{how:?}");
        let doc = open(damage(pptx(), "_rels/.rels", how), DocumentFormat::Pptx);
        assert_reported_damaged(&reported(&doc), ".rels", &format!("{how:?}"));
    }
}

#[test]
fn test_a_damaged_theme_is_a_warning_and_not_a_loss_of_text() {
    // A theme holds colours and fonts, no text: the document reads exactly as
    // an intact one, with no notice in the text and no truncation; the
    // metadata still lists the part.
    let intact = Document::from_reader(Cursor::new(xlsx()), DocumentFormat::Xlsx).unwrap();
    for how in DAMAGES {
        let doc = open(damage(xlsx(), "xl/theme/theme1.xml", how), DocumentFormat::Xlsx);
        assert_eq!(doc.plain_text(), intact.plain_text(), "{how:?}");
        assert_eq!(doc.to_markdown(), intact.to_markdown(), "{how:?}");
        let ir = doc.to_ir();
        assert!(!ir.metadata.text_truncated, "{how:?}");
        assert!(
            ir.metadata
                .warnings
                .iter()
                .any(|w| w.contains("theme1.xml")),
            "{how:?}: {:?}",
            ir.metadata.warnings
        );
    }
}

// ---- pptx -----------------------------------------------------------------

fn pptx() -> Vec<u8> {
    zip_of(&[
        (
            "[Content_Types].xml",
            types(
                r#"<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            ),
        ),
        ("_rels/.rels", root_rels("ppt/presentation.xml")),
        (
            "ppt/_rels/presentation.xml.rels",
            rels(&format!(
                r#"<Relationship Id="rId1" Type="{R}/slide" Target="slides/slide1.xml"/>"#
            )),
        ),
        (
            "ppt/presentation.xml",
            format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>"#
            ),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            rels(&format!(
                r#"<Relationship Id="rId1" Type="{R}/notesSlide" Target="../notesSlides/notes1.xml"/>"#
            )),
        ),
        (
            "ppt/slides/slide1.xml",
            format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>SLIDE</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
            ),
        ),
        (
            "ppt/notesSlides/notes1.xml",
            format!(
                r#"<p:notes xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="N"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>NOTE</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#
            ),
        ),
    ])
}

#[test]
fn test_a_pptx_part_that_cannot_be_decoded_is_reported() {
    for part in [
        "ppt/notesSlides/notes1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
    ] {
        for how in DAMAGES {
            let doc = open(damage(pptx(), part, how), DocumentFormat::Pptx);
            let found = reported(&doc);
            let leaf = part.rsplit('/').next().unwrap();
            assert_reported_damaged(&found, leaf, &format!("{part} {how:?}"));
            assert!(doc.to_ir().metadata.text_truncated, "{part} ({how:?})");
        }
    }
}
