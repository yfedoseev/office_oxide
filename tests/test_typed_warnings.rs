//! Every document-level warning is recorded with its kind.
//!
//! `Document::damage` and strict mode tell a loss from a repair by the kind a
//! warning was raised with. A warning written straight into
//! `Metadata::warnings` would reach the IR with no kind and be misfiled, so
//! the converters write that field only through `notice::report`, and the
//! readers keep theirs as typed notices. This test fails when a new site
//! bypasses that, whatever the warning says.

use std::path::Path;

fn rust_sources(dir: &Path, out: &mut Vec<std::path::PathBuf>) {
    for entry in std::fs::read_dir(dir).unwrap() {
        let path = entry.unwrap().path();
        if path.is_dir() {
            rust_sources(&path, out);
        } else if path.extension().is_some_and(|e| e == "rs") {
            out.push(path);
        }
    }
}

#[test]
fn test_metadata_warnings_are_written_only_with_their_kind() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("src");
    let mut files = Vec::new();
    rust_sources(&root, &mut files);
    let mut bad = Vec::new();
    for file in &files {
        let rel = file
            .strip_prefix(&root)
            .unwrap()
            .to_string_lossy()
            .replace('\\', "/");
        // An IR built from Markdown belongs to no document and has no
        // damage report.
        if rel == "ir_from_markdown.rs" {
            continue;
        }
        let text = std::fs::read_to_string(file).unwrap();
        let compact: String = text.split_whitespace().collect::<Vec<_>>().join(" ");
        for (what, needle) in [
            ("push", "metadata.warnings.push("),
            ("extend", "metadata.warnings.extend("),
            ("append", "metadata.warnings.append("),
            ("insert", "metadata.warnings.insert("),
        ] {
            if compact.contains(needle) {
                bad.push(format!("{rel}: metadata.warnings.{what}"));
            }
        }
        for (i, _) in compact.match_indices("metadata.warnings = ") {
            if !compact[i..].starts_with("metadata.warnings = crate::core::notice::report(") {
                bad.push(format!("{rel}: metadata.warnings assigned without notice::report"));
            }
        }
        // The `warnings` field of every `Metadata { .. }` literal, written out
        // or in shorthand, whatever the layout: its value must be built by
        // `notice::report`.
        for value in metadata_warnings_fields(&compact) {
            if !value.contains("notice::report(") {
                bad.push(format!("{rel}: Metadata {{ warnings: {value} }}"));
            }
        }
    }
    assert!(bad.is_empty(), "warnings written without a kind:\n{}", bad.join("\n"));
}

/// The value of the `warnings` field in each `Metadata { .. }` struct literal
/// of `code` (whitespace already collapsed): the expression after `warnings:`,
/// or `warnings` itself for the shorthand. The struct's own definition is
/// not a literal and is skipped.
fn metadata_warnings_fields(code: &str) -> Vec<String> {
    let bytes = code.as_bytes();
    let mut out = Vec::new();
    for (at, _) in code.match_indices("Metadata {") {
        let before = &code[..at];
        let whole_word = !before.ends_with(|c: char| c.is_alphanumeric() || c == '_');
        if !whole_word || before.ends_with("struct ") || before.ends_with("for ") {
            continue;
        }
        // The literal's body, to its matching brace.
        let start = at + "Metadata {".len();
        let (mut depth, mut end) = (1i32, start);
        while end < bytes.len() && depth > 0 {
            match bytes[end] {
                b'{' | b'(' | b'[' => depth += 1,
                b'}' | b')' | b']' => depth -= 1,
                _ => {},
            }
            end += 1;
        }
        let body = &code[start..end.saturating_sub(1)];
        // Its fields, split at the commas outside any bracket.
        let (mut depth, mut from) = (0i32, 0);
        let mut fields = Vec::new();
        for (i, b) in body.bytes().enumerate() {
            match b {
                b'{' | b'(' | b'[' => depth += 1,
                b'}' | b')' | b']' => depth -= 1,
                b',' if depth == 0 => {
                    fields.push(&body[from..i]);
                    from = i + 1;
                },
                _ => {},
            }
        }
        fields.push(&body[from..]);
        for field in fields {
            let field = field.trim();
            if field == "warnings" {
                out.push("warnings".to_string());
            } else if let Some(value) = field.strip_prefix("warnings:") {
                out.push(value.trim().to_string());
            }
        }
    }
    out
}

/// The scan finds the field however it is written, and only in a literal.
#[test]
fn test_the_scan_finds_the_warnings_field_of_a_metadata_literal() {
    let found = |code: &str| metadata_warnings_fields(code);
    assert_eq!(found("Metadata { title: None, warnings: w.clone(), ..x }"), ["w.clone()"]);
    assert_eq!(found("Metadata { warnings, title }"), ["warnings"]);
    assert_eq!(
        found("Metadata { warnings: { let n = f(a, b); crate::core::notice::report(n) }, x }"),
        ["{ let n = f(a, b); crate::core::notice::report(n) }"]
    );
    assert!(found("pub struct Metadata { pub warnings: Vec<String> }").is_empty());
    assert!(found("SheetIr { warnings, x }").is_empty());
}
