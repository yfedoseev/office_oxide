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
        // Inside a `Metadata { .. }` literal of a converter.
        if rel.starts_with("convert_") {
            for line in text.lines() {
                let t = line.trim_start();
                if t.starts_with("warnings:")
                    && !t.contains("notice::report(")
                    && !t.contains("Notices")
                    && t != "warnings: {"
                {
                    bad.push(format!("{rel}: `{t}`"));
                }
            }
        }
    }
    assert!(bad.is_empty(), "warnings written without a kind:\n{}", bad.join("\n"));
}
