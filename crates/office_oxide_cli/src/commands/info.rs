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
/// Every document property the IR carries is listed, read off the IR's own
/// serde form so a property added to `Metadata` shows up here without this
/// command having to be kept in step by hand. Author, subject, keywords and
/// the dates were parsed all along but never shown.
fn render(ir: &DocumentIR, file_size: Option<u64>) -> String {
    let mut s = String::new();
    let m = &ir.metadata;
    let _ = writeln!(s, "Format: {:?}", m.format);
    if let Some(size) = file_size {
        let _ = writeln!(s, "Size: {size} bytes");
    }
    if let Some(ref title) = m.title {
        let _ = writeln!(s, "Title: {title}");
    }
    if let Ok(serde_json::Value::Object(fields)) = serde_json::to_value(m) {
        for (key, value) in fields {
            // Shown on their own lines, in their own words.
            if matches!(key.as_str(), "format" | "title" | "has_macros" | "text_truncated") {
                continue;
            }
            let shown = match value {
                serde_json::Value::String(v) if !v.is_empty() => v,
                serde_json::Value::Array(items) if !items.is_empty() => items
                    .iter()
                    .map(|v| v.as_str().map_or_else(|| v.to_string(), str::to_string))
                    .collect::<Vec<_>>()
                    .join(", "),
                serde_json::Value::Number(n) => n.to_string(),
                serde_json::Value::Bool(true) => "yes".to_string(),
                _ => continue,
            };
            let _ = writeln!(s, "{}: {shown}", label(&key));
        }
    }
    if m.has_macros {
        let _ = writeln!(s, "Macros: yes");
    }
    if m.text_truncated {
        let _ = writeln!(
            s,
            "Warning: text extraction is incomplete — the source file's own structure disagrees \
             with itself about how much text there is, and the gap could not be safely recovered"
        );
    }
    let _ = writeln!(s, "Sections: {}", ir.sections.len());

    for (i, section) in ir.sections.iter().enumerate() {
        let title = section.title.as_deref().unwrap_or("(untitled)");
        let _ = writeln!(s, "  [{i}] {title} — {} elements", section.elements.len());
    }
    s
}

/// `last_modified_by` → `Last modified by`.
fn label(key: &str) -> String {
    let words = key.replace('_', " ");
    let mut chars = words.chars();
    match chars.next() {
        Some(first) => first.to_uppercase().chain(chars).collect(),
        None => String::new(),
    }
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
