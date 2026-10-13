//! Source compatibility of the IR constructors that downstream crates use.
//!
//! A converter in another crate builds `DocumentIR` with a literal that names
//! every field. Adding a public field to that struct in a patch release is a
//! compile error for every such crate, so `DocumentIR` must stay exactly
//! `{ metadata, sections }`; new document-level data belongs on `Metadata`
//! (which callers build with `..Default::default()`).

use office_oxide::DocumentIR;
use office_oxide::ir::Metadata;

#[test]
fn test_document_ir_is_constructible_from_metadata_and_sections_alone() {
    let ir = DocumentIR {
        metadata: Metadata::default(),
        sections: Vec::new(),
    };
    assert!(ir.sections.is_empty());
    assert!(ir.metadata.defined_names.is_empty());
}
