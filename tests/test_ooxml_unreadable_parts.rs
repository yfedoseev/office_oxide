//! The report of the parts a reader could not use is bounded like any other
//! output the file shapes: a part many references name is tried once and
//! reported once, the list is capped, and the rest is counted.

mod common;

use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::rc::Rc;

use common::LimitScope;
use office_oxide::docx::DocxDocument;
use office_oxide::pptx::PptxDocument;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

/// A reader that counts the bytes read through it.
struct Counting {
    inner: Cursor<Vec<u8>>,
    read: Rc<std::cell::Cell<u64>>,
}

impl Read for Counting {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(buf)?;
        self.read.set(self.read.get() + n as u64);
        Ok(n)
    }
}

impl Seek for Counting {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        self.inner.seek(pos)
    }
}

fn rels(entries: &str) -> String {
    format!(r#"<Relationships xmlns="{PKG_REL}">{entries}</Relationships>"#)
}

/// A zip of `parts`, stored (so that bytes read count what is decompressed).
fn zip_of(parts: Vec<(String, String)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A malformed header: the root never closes properly.
fn malformed_header() -> String {
    format!(
        r#"<w:hdr xmlns:w="{WML}"><w:p><w:r><w:t>x</w:t></w:p></w:r>{}</w:hdr"#,
        " ".repeat(9_000)
    )
}

/// A `.docx` of `sections` paragraphs, each ending a section whose header
/// reference is `rId(2 + i % headers)`; header `i` is a malformed part.
fn docx_with_headers(sections: usize, headers: usize, header: &str) -> Vec<u8> {
    let mut content_types = format!(
        r#"<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#
    );
    let mut main_rels = String::new();
    let mut parts: Vec<(String, String)> = Vec::new();
    for h in 0..headers {
        content_types.push_str(&format!(
            r#"<Override PartName="/word/header{h}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#
        ));
        main_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{REL}/header" Target="header{h}.xml"/>"#,
            h + 2
        ));
        parts.push((format!("word/header{h}.xml"), header.to_string()));
    }
    content_types.push_str("</Types>");
    let mut body = String::new();
    for i in 0..sections {
        body.push_str(&format!(
            r#"<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId{}"/></w:sectPr></w:pPr><w:r><w:t>p</w:t></w:r></w:p>"#,
            2 + i % headers
        ));
    }
    let mut all = vec![
        ("[Content_Types].xml".to_string(), content_types),
        (
            "_rels/.rels".to_string(),
            rels(&format!(
                r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>"#
            )),
        ),
        ("word/_rels/document.xml.rels".to_string(), rels(&main_rels)),
        (
            "word/document.xml".to_string(),
            format!(
                r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}"><w:body>{body}</w:body></w:document>"#
            ),
        ),
    ];
    all.extend(parts);
    zip_of(all)
}

/// Open `bytes` as `.docx`, returning the document and the bytes read.
fn open_docx_counting(bytes: Vec<u8>) -> (DocxDocument, u64) {
    let read = Rc::new(std::cell::Cell::new(0));
    let doc = DocxDocument::from_reader(Counting {
        inner: Cursor::new(bytes),
        read: read.clone(),
    })
    .expect("a malformed header does not fail the file");
    (doc, read.get())
}

fn skipped_warnings(ir: &office_oxide::ir::DocumentIR) -> usize {
    ir.metadata
        .warnings
        .iter()
        .filter(|w| w.contains("skipped unreadable part"))
        .count()
}

#[test]
fn test_one_malformed_header_named_by_every_section_is_read_and_reported_once() {
    let _limits = common::default_limits();
    let bytes = docx_with_headers(20_000, 1, &malformed_header());
    let size = bytes.len() as u64;
    let (doc, read) = open_docx_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 1, "{:?}", doc.unreadable_parts.len());
    assert_eq!(doc.unreadable_parts_omitted, 0);
    // The package is read once for its parts; the 9 KB header is not read
    // again for each of 20,000 sections.
    assert!(read < 3 * size, "{read} bytes read of a {size} byte package");
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(skipped_warnings(&ir), 1, "{:?}", ir.metadata.warnings.len());
    assert!(ir.metadata.text_truncated);
}

#[test]
fn test_many_malformed_headers_are_listed_up_to_a_cap_and_the_rest_counted() {
    let _limits = common::default_limits();
    let bytes = docx_with_headers(250, 250, &malformed_header());
    let (doc, _) = open_docx_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 100);
    assert_eq!(doc.unreadable_parts_omitted, 150);
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .unwrap()
        .to_ir();
    assert_eq!(skipped_warnings(&ir), 100);
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w == "and 150 more unreadable parts not listed"),
        "{:?}",
        ir.metadata.warnings.len()
    );
}

#[test]
fn test_one_header_that_cannot_be_read_is_tried_once_for_all_sections() {
    // The header is larger than the package byte budget allows, so reading
    // it fails (as a corrupt zip entry would).
    let mut scope = LimitScope::new();
    let bytes = docx_with_headers(2_000, 1, &"x".repeat(3_000_000));
    scope.set(
        office_oxide::limits::max_package_bytes,
        office_oxide::limits::set_max_package_bytes,
        2_000_000,
    );
    let (doc, read) = open_docx_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 1, "{}", doc.unreadable_parts.len());
    assert!(read < 3 * bytes.len() as u64, "{read} bytes read");
}

/// A `.pptx` of `slides` slides. Slide `i` relates to notes part
/// `notes{i % notes}.xml` (malformed) and, when `comments`, to one comments
/// part `comments.xml` of `comments_size` bytes.
fn pptx_with_notes(slides: usize, notes: usize, comments_size: Option<usize>) -> Vec<u8> {
    let mut content_types = format!(
        r#"<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#
    );
    let mut pres_rels = String::new();
    let mut ids = String::new();
    let mut parts: Vec<(String, String)> = Vec::new();
    for i in 0..slides {
        pres_rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{REL}/slide" Target="slides/slide{i}.xml"/>"#,
            i + 1
        ));
        ids.push_str(&format!(r#"<p:sldId id="{}" r:id="rId{}"/>"#, 256 + i, i + 1));
        let mut srels = format!(
            r#"<Relationship Id="rId1" Type="{REL}/notesSlide" Target="../notesSlides/notes{}.xml"/>"#,
            i % notes
        );
        if comments_size.is_some() {
            srels.push_str(&format!(
                r#"<Relationship Id="rId2" Type="{REL}/comments" Target="../comments/comments.xml"/>"#
            ));
        }
        parts.push((format!("ppt/slides/_rels/slide{i}.xml.rels"), rels(&srels)));
        parts.push((
            format!("ppt/slides/slide{i}.xml"),
            format!(r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>s{i}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#),
        ));
        content_types.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#
        ));
    }
    for n in 0..notes {
        parts.push((
            format!("ppt/notesSlides/notes{n}.xml"),
            format!(r#"<p:notes xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld><a:p><a:t>x</a:p></a:t>{}</p:cSld</p:notes"#, " ".repeat(9_000)),
        ));
    }
    if let Some(size) = comments_size {
        parts.push(("ppt/comments/comments.xml".into(), "x".repeat(size)));
    }
    content_types.push_str("</Types>");
    let mut all = vec![
        ("[Content_Types].xml".to_string(), content_types),
        (
            "_rels/.rels".to_string(),
            rels(&format!(
                r#"<Relationship Id="rId1" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/>"#
            )),
        ),
        ("ppt/_rels/presentation.xml.rels".to_string(), rels(&pres_rels)),
        (
            "ppt/presentation.xml".to_string(),
            format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst>{ids}</p:sldIdLst></p:presentation>"#
            ),
        ),
    ];
    all.extend(parts);
    zip_of(all)
}

fn open_pptx_counting(bytes: Vec<u8>) -> (PptxDocument, u64) {
    let read = Rc::new(std::cell::Cell::new(0));
    let doc = PptxDocument::from_reader(Counting {
        inner: Cursor::new(bytes),
        read: read.clone(),
    })
    .expect("an unreadable auxiliary part does not fail the deck");
    (doc, read.get())
}

#[test]
fn test_one_malformed_notes_part_named_by_many_slides_is_reported_once() {
    let _limits = common::default_limits();
    let bytes = pptx_with_notes(3_000, 1, None);
    let (doc, _) = open_pptx_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 1, "{}", doc.unreadable_parts.len());
    assert_eq!(doc.unreadable_parts_omitted, 0);
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_ir();
    assert_eq!(skipped_warnings(&ir), 1);
}

#[test]
fn test_many_malformed_notes_parts_are_listed_up_to_a_cap_and_the_rest_counted() {
    let _limits = common::default_limits();
    let bytes = pptx_with_notes(250, 250, None);
    let (doc, _) = open_pptx_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 100);
    assert_eq!(doc.unreadable_parts_omitted, 150);
    let ir = Document::from_reader(Cursor::new(bytes), DocumentFormat::Pptx)
        .unwrap()
        .to_ir();
    assert_eq!(skipped_warnings(&ir), 100);
    assert!(ir.metadata.text_truncated);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w == "and 150 more unreadable parts not listed")
    );
}

#[test]
fn test_one_comments_part_that_cannot_be_read_is_tried_once_for_all_slides() {
    let mut scope = LimitScope::new();
    scope.set(
        office_oxide::limits::max_package_bytes,
        office_oxide::limits::set_max_package_bytes,
        2_000_000,
    );
    let bytes = pptx_with_notes(1_500, 1, Some(3_000_000));
    let (doc, read) = open_pptx_counting(bytes.clone());
    let named = doc
        .unreadable_parts
        .iter()
        .filter(|(p, _)| p.contains("comments"))
        .count();
    assert_eq!(named, 1, "{}", doc.unreadable_parts.len());
    assert!(read < 2 * bytes.len() as u64, "{read} bytes read of {}", bytes.len());
}

#[test]
fn test_a_comments_part_that_is_not_well_formed_is_reported() {
    let _limits = common::default_limits();
    let comments = format!(
        r#"<p:cmLst xmlns:p="{PML}"><p:cm authorId="0"><p:text>hi</p:text></p:cm><p:cm authorId="0"#
    );
    let mut zip_bytes = pptx_with_notes(1, 1, None);
    // Add the part and its relationship by rebuilding the package.
    let mut archive = zip::ZipArchive::new(Cursor::new(std::mem::take(&mut zip_bytes))).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        if name == "ppt/slides/_rels/slide0.xml.rels" {
            body = body.replace(
                "</Relationships>",
                &format!(r#"<Relationship Id="rId2" Type="{REL}/comments" Target="../comments/comments.xml"/></Relationships>"#),
            );
        }
        parts.push((name, body));
    }
    parts.push(("ppt/comments/comments.xml".into(), comments));
    let (doc, _) = open_pptx_counting(zip_of(parts));
    assert!(
        doc.unreadable_parts
            .iter()
            .any(|(p, _)| p == "/ppt/comments/comments.xml"),
        "{:?}",
        doc.unreadable_parts
    );
}

#[test]
fn test_a_theme_part_that_is_not_well_formed_is_reported_and_the_file_still_opens() {
    let _limits = common::default_limits();
    let bytes = docx_with_headers(1, 1, "<w:hdr/>");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        if name == "word/_rels/document.xml.rels" {
            body = body.replace(
                "</Relationships>",
                &format!(r#"<Relationship Id="rId90" Type="{REL}/theme" Target="theme/theme1.xml"/></Relationships>"#),
            );
        }
        parts.push((name, body));
    }
    parts.push(("word/theme/theme1.xml".into(), "<a:theme><a:x></a:theme>".into()));
    let (doc, _) = open_docx_counting(zip_of(parts));
    assert!(
        doc.unreadable_parts
            .iter()
            .any(|p| p.starts_with("/word/theme/theme1.xml")),
        "{:?}",
        doc.unreadable_parts
    );
}

/// Every rendering of `doc`: plain text, Markdown, HTML and the Markdown and
/// HTML of its IR.
fn renderings(bytes: &[u8], format: DocumentFormat) -> Vec<(&'static str, String)> {
    let doc = Document::from_reader(Cursor::new(bytes.to_vec()), format).unwrap();
    let ir = doc.to_ir();
    vec![
        ("plain_text", doc.plain_text()),
        ("to_markdown", doc.to_markdown()),
        ("to_html", doc.to_html()),
        ("ir plain", ir.plain_text()),
        ("ir markdown", ir.to_markdown()),
        ("ir html", ir.to_html()),
    ]
}

#[test]
fn test_a_docx_header_that_cannot_be_read_is_reported_in_every_rendering() {
    let _limits = common::default_limits();
    let bad = docx_with_headers(1, 1, &malformed_header());
    let good = docx_with_headers(1, 1, &format!(r#"<w:hdr xmlns:w="{WML}"/>"#));
    for (what, text) in renderings(&bad, DocumentFormat::Docx) {
        assert!(
            text.contains("skipped unreadable part /word/header0.xml"),
            "{what} does not report the lost header: {text}"
        );
    }
    for (what, text) in renderings(&good, DocumentFormat::Docx) {
        assert!(!text.contains("unreadable"), "{what} reports a loss on a good file: {text}");
    }
    // The notice is all a lost part adds: the rest reads as the good file.
    let good_plain = Document::from_reader(Cursor::new(good), DocumentFormat::Docx)
        .unwrap()
        .plain_text();
    let bad_plain = Document::from_reader(Cursor::new(bad), DocumentFormat::Docx)
        .unwrap()
        .plain_text();
    assert_eq!(bad_plain.split("\n\n[skipped").next().unwrap(), good_plain);
}

#[test]
fn test_a_deck_part_that_cannot_be_read_is_reported_in_every_rendering() {
    let _limits = common::default_limits();
    let bad = pptx_with_notes(2, 1, None);
    for (what, text) in renderings(&bad, DocumentFormat::Pptx) {
        assert!(
            text.contains("skipped unreadable part /ppt/notesSlides/notes0.xml"),
            "{what} does not report the lost notes: {text}"
        );
    }
    // A deck whose parts all read has no report.
    let mut parts = Vec::new();
    let mut archive = zip::ZipArchive::new(Cursor::new(bad)).unwrap();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        if name == "ppt/notesSlides/notes0.xml" {
            body = format!(r#"<p:notes xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld/></p:notes>"#);
        }
        parts.push((name, body));
    }
    for (what, text) in renderings(&zip_of(parts), DocumentFormat::Pptx) {
        assert!(!text.contains("unreadable"), "{what} reports a loss on a good file: {text}");
    }
}

#[test]
fn test_the_report_in_the_text_is_bounded_to_the_list_and_the_count() {
    let _limits = common::default_limits();
    let docx = docx_with_headers(250, 250, &malformed_header());
    let pptx = pptx_with_notes(250, 250, None);
    for (bytes, format) in [(docx, DocumentFormat::Docx), (pptx, DocumentFormat::Pptx)] {
        for (what, text) in renderings(&bytes, format) {
            assert_eq!(text.matches("skipped unreadable part").count(), 100, "{what}");
            assert!(text.contains("and 150 more unreadable parts not listed"), "{what}");
        }
    }
}

#[test]
fn test_a_report_the_text_budget_cannot_pay_for_is_left_out_and_the_budget_says_so() {
    let mut scope = LimitScope::new();
    scope.set(
        office_oxide::limits::max_text_chars,
        office_oxide::limits::set_max_text_chars,
        40,
    );
    let docx = docx_with_headers(1, 1, &malformed_header());
    let pptx = pptx_with_notes(1, 1, None);
    for (bytes, format) in [(docx, DocumentFormat::Docx), (pptx, DocumentFormat::Pptx)] {
        let doc = Document::from_reader(Cursor::new(bytes), format).unwrap();
        let ir = doc.to_ir();
        for (what, text) in [
            ("plain_text", doc.plain_text()),
            ("to_markdown", doc.to_markdown()),
            ("ir markdown", ir.to_markdown()),
        ] {
            assert!(!text.contains("skipped unreadable part"), "{what}: {text}");
            assert!(text.contains("output truncated"), "{what}: {text}");
        }
        // The metadata still lists what was lost.
        assert!(ir.metadata.text_truncated);
        assert_eq!(skipped_warnings(&ir), 1);
    }
}

/// A part name the file shapes: link and markup syntax, a code span, and a
/// line break that would start a list item.
const HOSTILE_NAME: &str = "h](http:/evil) <img src=x onerror=1> `t`\n\n- Heading";

/// `bytes` with every part named `from` renamed to `HOSTILE_NAME` followed by
/// `.xml`: raw in the zip, percent-encoded where a relationship names it.
fn hostile_part(bytes: &[u8], from: &str) -> Vec<u8> {
    let encoded = HOSTILE_NAME
        .replace(' ', "%20")
        .replace('<', "%3C")
        .replace('>', "%3E")
        .replace('`', "%60")
        .replace('\n', "%0A");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f
            .name()
            .unwrap()
            .replace(from, &format!("{HOSTILE_NAME}.xml"));
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        parts.push((name, body.replace(from, &format!("{encoded}.xml"))));
    }
    zip_of(parts)
}

#[test]
fn test_a_part_name_the_file_shapes_stays_one_inert_line_in_every_rendering() {
    let _limits = common::default_limits();
    for (bytes, format) in [
        (
            hostile_part(&docx_with_headers(1, 1, &malformed_header()), "header0"),
            DocumentFormat::Docx,
        ),
        (hostile_part(&pptx_with_notes(1, 1, None), "notes0"), DocumentFormat::Pptx),
    ] {
        let doc = Document::from_reader(Cursor::new(bytes), format).unwrap();
        let ir = doc.to_ir();
        // The metadata keeps the name as the file gave it.
        assert!(
            ir.metadata
                .warnings
                .iter()
                .any(|w| w.contains(HOSTILE_NAME)),
            "{:?}",
            ir.metadata.warnings
        );
        for (what, text) in [
            ("plain_text", doc.plain_text()),
            ("to_markdown", doc.to_markdown()),
            ("to_html", doc.to_html()),
            ("ir plain", ir.plain_text()),
            ("ir markdown", ir.to_markdown()),
            ("ir html", ir.to_html()),
        ] {
            let reports: Vec<&str> = text
                .lines()
                .filter(|l| l.contains("skipped unreadable part"))
                .collect();
            assert_eq!(reports.len(), 1, "{what}: {text}");
            // The name's line break did not start a block of its own.
            assert!(
                !text
                    .lines()
                    .any(|l| l.trim_start().starts_with("- Heading")),
                "{what}: {text}"
            );
            if what.contains("html") {
                assert!(!text.contains("<img"), "{what}: {text}");
            } else if what.contains("markdown") {
                // A tag, a link and an image opener are escaped.
                assert_eq!(text.matches("<img").count(), text.matches("\\<img").count(), "{what}");
                assert_eq!(text.matches("](").count(), text.matches("\\](").count(), "{what}");
                assert!(!text.contains("`t`"), "{what}: {text}");
            }
        }
    }
}

/// A deck whose parts all read (`pptx_with_notes` with a good notes part), with
/// `edit` applied to each part as `(name, body)`.
fn clean_deck(edit: impl Fn(&str, String) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(pptx_with_notes(1, 1, None))).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        if name == "ppt/notesSlides/notes0.xml" {
            body = format!(r#"<p:notes xmlns:p="{PML}" xmlns:a="{DML}"><p:cSld/></p:notes>"#);
        }
        let body = edit(&name, body);
        parts.push((name, body));
    }
    zip_of(parts)
}

fn add_rel(body: String, rel: &str) -> String {
    body.replace("</Relationships>", &format!("{rel}</Relationships>"))
}

/// Nothing was lost, and nothing says so.
fn assert_nothing_reported(bytes: Vec<u8>, format: DocumentFormat) {
    let doc = Document::from_reader(Cursor::new(bytes), format).unwrap();
    let ir = doc.to_ir();
    assert!(!ir.metadata.text_truncated, "{:?}", ir.metadata.warnings);
    assert!(
        !ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("unreadable")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(!doc.plain_text().contains("unreadable"));
}

/// A relationship marked external names a place outside the package, whatever
/// its target looks like: here a target that, read as a part name, is a part
/// the package holds and cannot read. It is not that part, so nothing is lost.
#[test]
fn test_a_deck_relationship_to_an_external_target_is_not_read_as_a_part() {
    let _limits = common::default_limits();
    let bad = "/ppt/ext/bad.xml";
    for (part, rel) in [
        ("ppt/slides/_rels/slide0.xml.rels", "comments"),
        ("ppt/slides/_rels/slide0.xml.rels", "slideLayout"),
        ("ppt/_rels/presentation.xml.rels", "theme"),
    ] {
        let bytes = clean_deck(|name, body| {
            if name == "ppt/ext/bad.xml" {
                return "<broken".to_string();
            }
            if name == part {
                return add_rel(
                    body,
                    &format!(
                        r#"<Relationship Id="rId8" Type="{REL}/{rel}" Target="{bad}" TargetMode="External"/>"#
                    ),
                );
            }
            body
        });
        // The part is in the package to be read, were the target taken for it.
        let mut parts = zip_entries(&bytes);
        parts.push(("ppt/ext/bad.xml".into(), "<broken".into()));
        assert_nothing_reported(zip_of(parts), DocumentFormat::Pptx);
    }
    // The notes relationship itself pointing outside the package.
    let mut parts = zip_entries(&clean_deck(|name, body| {
        if name == "ppt/slides/_rels/slide0.xml.rels" {
            body.replace(
                r#"Target="../notesSlides/notes0.xml""#,
                &format!(r#"Target="{bad}" TargetMode="External""#),
            )
        } else {
            body
        }
    }));
    parts.push(("ppt/ext/bad.xml".into(), "<broken".into()));
    assert_nothing_reported(zip_of(parts), DocumentFormat::Pptx);
}

/// A chart relationship marked external is not read as the part its target
/// spells: the package is read as if it were not there.
#[test]
fn test_a_deck_chart_relationship_to_an_external_target_is_not_read() {
    let _limits = common::default_limits();
    let with = |rel: &str| {
        let mut parts = zip_entries(&clean_deck(|name, body| {
            if name == "ppt/slides/_rels/slide0.xml.rels" {
                add_rel(body, rel)
            } else {
                body
            }
        }));
        parts.push(("ppt/ext/chart.xml".into(), " ".repeat(200_000)));
        let read = Rc::new(std::cell::Cell::new(0));
        PptxDocument::from_reader(Counting {
            inner: Cursor::new(zip_of(parts)),
            read: read.clone(),
        })
        .unwrap();
        read.get()
    };
    let internal = with(&format!(
        r#"<Relationship Id="rId8" Type="{REL}/chart" Target="/ppt/ext/chart.xml"/>"#
    ));
    let external = with(&format!(
        r#"<Relationship Id="rId8" Type="{REL}/chart" Target="/ppt/ext/chart.xml" TargetMode="External"/>"#
    ));
    assert!(
        internal >= external + 200_000,
        "{internal} bytes read with the part named, {external} without"
    );
}

/// The entries of a zip, as `(name, body)`.
fn zip_entries(bytes: &[u8]) -> Vec<(String, String)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        parts.push((name, body));
    }
    parts
}

#[test]
fn test_a_dangling_numbering_relationship_is_not_a_lost_part() {
    let _limits = common::default_limits();
    let mut archive = zip::ZipArchive::new(Cursor::new(docx_with_headers(
        1,
        1,
        &format!(r#"<w:hdr xmlns:w="{WML}"/>"#),
    )))
    .unwrap();
    let mut parts = Vec::new();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().unwrap().to_string();
        let mut body = String::new();
        f.read_to_string(&mut body).unwrap();
        if name == "word/_rels/document.xml.rels" {
            body = add_rel(
                body,
                &format!(
                    r#"<Relationship Id="rId90" Type="{REL}/numbering" Target="numbering.xml"/>"#
                ),
            );
        }
        parts.push((name, body));
    }
    assert_nothing_reported(zip_of(parts), DocumentFormat::Docx);
}

#[test]
fn test_a_deck_with_no_readable_slide_names_a_slide_not_the_first_part_that_failed() {
    let _limits = common::default_limits();
    let mut parts = zip_entries(&clean_deck(|name, body| {
        match name {
            // The theme fails first, as the reader meets it before any slide.
            "ppt/_rels/presentation.xml.rels" => add_rel(
                body,
                &format!(
                    r#"<Relationship Id="rId9" Type="{REL}/theme" Target="theme/theme1.xml"/>"#
                ),
            ),
            "ppt/slides/slide0.xml" => "<p:sld".to_string(),
            _ => body,
        }
    }));
    parts.push(("ppt/theme/theme1.xml".into(), "<a:theme><a:x></a:theme>".into()));
    let err = match PptxDocument::from_reader(Cursor::new(zip_of(parts))) {
        Ok(_) => panic!("a deck with no readable slide opened"),
        Err(e) => e.to_string(),
    };
    assert!(err.contains("no slide could be read"), "{err}");
    assert!(err.contains("/ppt/slides/slide0.xml"), "{err}");
    assert!(!err.contains("theme"), "{err}");
}
