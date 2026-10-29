//! The slide list is bounded by the package entry limit. This test lowers
//! the process-wide limit for the length of the test.

use std::io::{Cursor, Write};

use office_oxide::limits::{max_package_entries, set_max_package_entries};
use office_oxide::pptx::PptxDocument;

mod common;
use common::LimitScope;

#[test]
fn test_slide_list_longer_than_the_entry_limit_is_refused() {
    let pml = "http://schemas.openxmlformats.org/presentationml/2006/main";
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let entries: String = (0..40)
        .map(|i| format!(r#"<p:sldId id="{}" r:id="rId2"/>"#, 256 + i))
        .collect();
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>"#.to_string()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#)),
        ("ppt/presentation.xml", format!(r#"<p:presentation xmlns:p="{pml}" xmlns:r="{rel}"><p:sldIdLst>{entries}</p:sldIdLst></p:presentation>"#)),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = zip::write::SimpleFileOptions::default();
    for (name, data) in parts {
        zip.start_file(name, opts).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();

    let mut limits = LimitScope::new();
    limits.set(max_package_entries, set_max_package_entries, 10);
    let err = PptxDocument::from_reader(Cursor::new(bytes))
        .expect_err("40 slides listed under a 10-entry limit");
    assert!(err.to_string().contains("slide list"), "{err}");
}
