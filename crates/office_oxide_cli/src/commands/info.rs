use std::fmt::Write as _;
use std::io::Write as _;

use office_oxide::{Document, DocumentIR};

pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open(file)?;
    let ir = doc.to_ir();
    let size = std::fs::metadata(file).ok().map(|m| m.len());
    let mut out = std::io::stdout().lock();
    out.write_all(render(&ir, size).as_bytes())?;
    out.flush()?;
    Ok(())
}

/// The `info` report for a parsed document.
///
/// Every set document property is listed via [`Metadata::properties`], the
/// one list the CLI and MCP share, so a property added to `Metadata` is
/// shown by both. Author, subject, keywords and the dates were parsed all
/// along but never shown.
///
/// [`Metadata::properties`]: office_oxide::ir::Metadata::properties
fn render(ir: &DocumentIR, file_size: Option<u64>) -> String {
    let mut s = String::new();
    let m = &ir.metadata;
    let _ = writeln!(s, "Format: {:?}", m.format);
    if let Some(size) = file_size {
        let _ = writeln!(s, "Size: {size} bytes");
    }
    for (label, value) in m.properties() {
        let _ = writeln!(s, "{label}: {value}");
    }
    for p in &m.custom_properties {
        let _ = writeln!(s, "Custom property: {} = {}", p.name, p.value);
    }
    if m.has_macros {
        let _ = writeln!(s, "Macros: yes");
    }
    if m.has_digital_signature {
        let _ = writeln!(s, "Digitally signed: yes");
    }
    if m.thumbnail.is_some() {
        let _ = writeln!(s, "Thumbnail: yes");
    }
    if m.text_truncated {
        let _ = writeln!(
            s,
            "Warning: text extraction is incomplete — part of the document could not be \
             recovered safely"
        );
    }
    // What the reader worked around, one line each: a skipped part, a
    // flattened structure, a container whose counts disagree.
    for w in &m.warnings {
        let _ = writeln!(s, "Warning: {w}");
    }
    let _ = writeln!(s, "Sections: {}", ir.sections.len());

    for (i, section) in ir.sections.iter().enumerate() {
        let title = section.title.as_deref().unwrap_or("(untitled)");
        let _ = writeln!(s, "  [{i}] {title} — {} elements", section.elements.len());
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::{Metadata, Section};

    /// Author, subject, keywords, description and both dates were in
    /// `Metadata` but `info` printed only the format and title.
    #[test]
    fn test_info_shows_every_populated_document_property() {
        // Field by field, so a property added to `Metadata` later does not
        // break this test's construction.
        let mut metadata = Metadata {
            format: DocumentFormat::Docx,
            ..Default::default()
        };
        metadata.title = Some("Quarterly".into());
        metadata.author = Some("Ada".into());
        metadata.subject = Some("Numbers".into());
        metadata.keywords = vec!["annual report".into(), "q3".into()];
        metadata.created = Some("2024-01-02T03:04:05Z".into());
        metadata.modified = Some("2024-02-03T04:05:06Z".into());
        metadata.description = Some("Summary".into());
        metadata.has_macros = true;
        metadata.text_truncated = true;
        metadata.warnings = vec!["skipped unreadable part /word/header1.xml: CRC".into()];
        let ir = DocumentIR {
            metadata,
            sections: vec![Section::default()],
            defined_names: Vec::new(),
        };
        let out = render(&ir, Some(1234));
        for needle in [
            "Format: Docx",
            "Size: 1234 bytes",
            "Title: Quarterly",
            "Author: Ada",
            "Subject: Numbers",
            "Keywords: annual report, q3",
            "Created: 2024-01-02T03:04:05Z",
            "Modified: 2024-02-03T04:05:06Z",
            "Description: Summary",
            "Macros: yes",
            "Warning: text extraction is incomplete",
            "Warning: skipped unreadable part /word/header1.xml: CRC",
            "Sections: 1",
        ] {
            assert!(out.contains(needle), "missing {needle:?} in:\n{out}");
        }
    }

    /// Absent properties are left out rather than printed empty.
    #[test]
    fn test_info_omits_absent_properties() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Xlsx,
                ..Default::default()
            },
            sections: Vec::new(),
            defined_names: Vec::new(),
        };
        let out = render(&ir, None);
        assert_eq!(out, "Format: Xlsx\nSections: 0\n");
    }
}
