//! Package-wide resource bounds, exercised through the public API.
//!
//! These tests lower the process-wide limits, so they live in their own
//! test binary: nothing else runs in this process to be affected.

use std::io::Cursor;

use office_oxide::core::Error as CoreError;
use office_oxide::core::editable::EditablePackage;
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::{Document, DocumentFormat, OfficeError};

/// A DOCX whose three media parts each hold `part_len` bytes.
fn docx_with_media(part_len: usize) -> Vec<u8> {
    let mut w = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let doc = PartName::new("/word/document.xml").unwrap();
    w.add_part(
        &doc,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#,
    )
    .unwrap();
    w.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    for i in 0..3 {
        let media = PartName::new(&format!("/word/media/blob{i}.bin")).unwrap();
        w.add_part(&media, "application/octet-stream", &vec![b'z'; part_len])
            .unwrap();
    }
    w.finish().unwrap().into_inner()
}

fn is_package_limit(e: &CoreError) -> bool {
    matches!(e, CoreError::PackageLimit(_))
}

#[test]
fn test_package_limits_bound_the_edit_and_read_paths() {
    // Each media part fits the per-part cap; together they do not fit a
    // 2.5-part package budget.
    let bytes = docx_with_media(1 << 20);
    office_oxide::limits::set_max_package_bytes((5 << 20) / 2);

    // The edit path loads every part eagerly.
    let err = EditablePackage::from_reader(Cursor::new(bytes.clone()))
        .err()
        .expect("the edit path must refuse the package");
    assert!(is_package_limit(&err), "{err:?}");

    // The read path does not touch the media and still opens.
    let doc = Document::from_reader(Cursor::new(bytes.clone()), DocumentFormat::Docx)
        .expect("reading the text does not read the media");
    assert_eq!(doc.plain_text().trim(), "x");

    office_oxide::limits::set_max_package_bytes(u64::MAX);
    EditablePackage::from_reader(Cursor::new(bytes.clone())).expect("unbounded again");

    // Entry count: the package has 3 media + document + rels + content types.
    office_oxide::limits::set_max_package_entries(3);
    let err = Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
        .err()
        .expect("too many entries");
    assert!(
        matches!(err, OfficeError::Docx(_) | OfficeError::Core(_))
            && err.to_string().contains("package limit"),
        "{err}"
    );
    office_oxide::limits::set_max_package_entries(
        office_oxide::limits::DEFAULT_MAX_PACKAGE_ENTRIES,
    );
    office_oxide::limits::set_max_package_bytes(office_oxide::limits::DEFAULT_MAX_PACKAGE_BYTES);
}
