//! The size of one decompressed part, and the ratio of that size to the
//! bytes that encode it.
//!
//! A part is refused (an `Err`, never a crash or an allocation of the
//! claimed size) when it inflates past `max_part_bytes`, or when it is large
//! (more than `INFLATE_RATIO_FLOOR_BYTES`) and inflates more than
//! `max_inflate_ratio` times its stored size.

mod common;

use std::io::Cursor;

use common::LimitScope;
use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;
use office_oxide::limits;
use office_oxide::{Document, DocumentFormat};

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// A `.docx` whose body repeats one empty paragraph until it is `bytes` long.
fn docx_with_body_of(bytes: usize) -> Vec<u8> {
    let head = format!(r#"<w:document xmlns:w="{WML}"><w:body>"#);
    let unit = "<w:p><w:r><w:t>x</w:t></w:r></w:p>";
    let mut xml = String::with_capacity(bytes + 64);
    xml.push_str(&head);
    while xml.len() < bytes {
        xml.push_str(unit);
    }
    xml.push_str("</w:body></w:document>");
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            xml.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    writer.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> office_oxide::Result<Document> {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx)
}

#[test]
fn test_a_part_larger_than_the_part_limit_is_refused() {
    let bytes = docx_with_body_of(1 << 20);
    let mut scope = LimitScope::new();
    assert!(open(bytes.clone()).is_ok(), "the default accepts a 1 MiB part");
    scope.set(limits::max_part_bytes, limits::set_max_part_bytes, 512 * 1024);
    let err = match open(bytes) {
        Ok(_) => panic!("a 1 MiB part is over a 512 KiB limit"),
        Err(e) => e,
    };
    assert!(err.to_string().contains("word/document.xml"), "{err}");
}

#[test]
fn test_a_large_part_with_an_extreme_ratio_is_refused() {
    // 70 MiB of one repeated element deflates to about 70 KiB (ratio ~1,000).
    let bytes = docx_with_body_of(70 << 20);
    assert!(bytes.len() < 1 << 20, "{} bytes stored", bytes.len());
    let mut scope = LimitScope::new();
    scope.set(limits::max_inflate_ratio, limits::set_max_inflate_ratio, 100);
    let err = match open(bytes) {
        Ok(_) => panic!("a 70 MiB part stored in under 1 MiB is over a ratio of 100"),
        Err(e) => e,
    };
    assert!(err.to_string().contains("word/document.xml"), "{err}");
}

#[test]
fn test_a_small_part_is_never_refused_for_its_ratio() {
    // 8 MiB stored in a few KiB is a ratio of ~1,000, but a small part is
    // harmless; the ratio limit looks only at parts past the floor.
    let bytes = docx_with_body_of(8 << 20);
    let mut scope = LimitScope::new();
    scope.set(limits::max_inflate_ratio, limits::set_max_inflate_ratio, 2);
    assert!(open(bytes).is_ok());
}

/// Raise the compressed size the central directory declares for `name`
/// (offset 20 of its header) without touching the data.
fn forge_compressed_size(bytes: &mut [u8], name: &str, size: u32) {
    let mut at = 0;
    while let Some(found) = bytes[at..].windows(4).position(|w| w == b"PK\x01\x02") {
        let header = at + found;
        let name_at = header + 46;
        if bytes[name_at..].starts_with(name.as_bytes()) {
            bytes[header + 20..header + 24].copy_from_slice(&size.to_le_bytes());
            return;
        }
        at = header + 4;
    }
    panic!("no central directory entry for {name}");
}

#[test]
fn test_a_header_that_declares_a_larger_stored_size_does_not_raise_the_ratio_cap() {
    // The same 70 MiB part as above, but the central directory claims it is
    // stored in 1 MiB, which would put its ratio under the limit. The file
    // cannot hold that many bytes between this entry and the next, so the
    // claim is not believed.
    let mut bytes = docx_with_body_of(70 << 20);
    let mut scope = LimitScope::new();
    scope.set(limits::max_inflate_ratio, limits::set_max_inflate_ratio, 100);
    forge_compressed_size(&mut bytes, "word/document.xml", 1 << 20);
    let err = match open(bytes) {
        Ok(_) => panic!("a forged stored size let a bomb through the ratio check"),
        Err(e) => e,
    };
    assert!(err.to_string().contains("word/document.xml"), "{err}");
}
