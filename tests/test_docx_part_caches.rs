//! A part many relationships name is read once, however the relationships
//! spell its name, and a relationship to a part the package does not hold is
//! not a loss: the reader counts the bytes it inflates and reports what it
//! could not read.

use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::rc::Rc;

use office_oxide::docx::DocxDocument;

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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

/// A zip of `parts`, stored (so the bytes read count what is inflated).
fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, body) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(&body).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A `.docx` whose body is `body`, whose main part has the relationships
/// `rels` and whose package also holds `parts`.
fn docx_with(body: &str, rels: &str, parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="docx" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#
    );
    let mut all = vec![
        ("[Content_Types].xml".to_string(), content_types.into_bytes()),
        (
            "_rels/.rels".to_string(),
            format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/document.xml.rels".to_string(),
            format!(r#"<Relationships xmlns="{PKG_REL}">{rels}</Relationships>"#).into_bytes(),
        ),
        (
            "word/document.xml".to_string(),
            format!(
                r#"<w:document xmlns:w="{WML}" xmlns:r="{REL}" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>{body}</w:body></w:document>"#
            )
            .into_bytes(),
        ),
    ];
    all.extend(parts);
    zip_of(all)
}

/// Open `bytes`, returning the document and the bytes read.
fn open_counting(bytes: Vec<u8>) -> (DocxDocument, u64) {
    let read = Rc::new(std::cell::Cell::new(0));
    let doc = DocxDocument::from_reader(Counting {
        inner: Cursor::new(bytes),
        read: read.clone(),
    })
    .expect("the document opens");
    (doc, read.get())
}

/// `n` relationships `rId100..`, of `kind`, each naming `target`.
fn rels_to(n: usize, kind: &str, target: impl Fn(usize) -> String) -> String {
    (0..n)
        .map(|i| {
            format!(
                r#"<Relationship Id="rId{}" Type="{REL}/{kind}" Target="{}"/>"#,
                100 + i,
                target(i)
            )
        })
        .collect()
}

const COPIES: usize = 40;
const PART_BYTES: usize = 100_000;

/// A chart part of about `PART_BYTES` bytes.
fn chart_part() -> Vec<u8> {
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>{}"#,
        " ".repeat(PART_BYTES)
    )
    .into_bytes()
}

fn assert_read_once(bytes: &[u8], read: u64) {
    assert!(
        read < 3 * bytes.len() as u64,
        "{read} bytes read of a {} byte package: the part was read again for each relationship",
        bytes.len()
    );
}

#[test]
fn test_an_image_part_many_relationships_name_is_read_once() {
    let bytes = docx_with(
        "<w:p/>",
        &rels_to(COPIES, "image", |_| "media/image1.png".into()),
        vec![("word/media/image1.png".into(), vec![7u8; PART_BYTES])],
    );
    let (doc, read) = open_counting(bytes.clone());
    assert_eq!(doc.images.len(), COPIES);
    assert_read_once(&bytes, read);
}

#[test]
fn test_a_chart_part_many_relationships_name_is_read_once() {
    let bytes = docx_with(
        "<w:p/>",
        &rels_to(COPIES, "chart", |_| "charts/chart1.xml".into()),
        vec![("word/charts/chart1.xml".into(), chart_part())],
    );
    let (_, read) = open_counting(bytes.clone());
    assert_read_once(&bytes, read);
}

#[test]
fn test_a_diagram_data_part_many_relationships_name_is_read_once() {
    let bytes = docx_with(
        "<w:p/>",
        &rels_to(COPIES, "diagramData", |_| "diagrams/data1.xml".into()),
        vec![("word/diagrams/data1.xml".into(), chart_part())],
    );
    let (_, read) = open_counting(bytes.clone());
    assert_read_once(&bytes, read);
}

/// A body holding one embedded-object reference to each of the relationships
/// `rId100..`.
fn ole_body(n: usize) -> String {
    (0..n)
        .map(|i| {
            format!(
                r#"<w:p><w:r><w:pict><o:OLEObject Type="Embed" r:id="rId{}"/></w:pict></w:r></w:p>"#,
                100 + i
            )
        })
        .collect()
}

#[test]
fn test_an_embedded_package_many_objects_name_is_opened_once() {
    // The embedded part is not a package: opening it fails, and the failure
    // is remembered as a success is.
    let bytes = docx_with(
        &ole_body(COPIES),
        &rels_to(COPIES, "package", |_| "embeddings/e.docx".into()),
        vec![("word/embeddings/e.docx".into(), vec![b'x'; PART_BYTES])],
    );
    let (doc, read) = open_counting(bytes.clone());
    assert_eq!(doc.unreadable_parts.len(), 1, "{:?}", doc.unreadable_parts);
    assert_read_once(&bytes, read);
}

#[test]
fn test_an_embedded_package_that_opens_is_opened_once() {
    let inner = docx_with(
        "<w:p><w:r><w:t>inner text</w:t></w:r></w:p>",
        "",
        vec![("word/pad.xml".into(), vec![b' '; PART_BYTES])],
    );
    let bytes = docx_with(
        &ole_body(COPIES),
        &rels_to(COPIES, "package", |_| "embeddings/e.docx".into()),
        vec![("word/embeddings/e.docx".into(), inner)],
    );
    let (doc, read) = open_counting(bytes.clone());
    assert!(doc.plain_text().contains("inner text"));
    assert_read_once(&bytes, read);
}
