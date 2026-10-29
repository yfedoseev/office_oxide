//! `--strict` and `--limit NAME=VALUE`: per-run options for every command
//! that reads a document, and `info` listing what is wrong with a file.

use std::io::Cursor;
use std::process::{Command, Output};

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;

/// A `.docx` whose body is cut off part-way: the first paragraph is kept.
fn cut_off_docx(name: &str) -> std::path::PathBuf {
    let body = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p><w:p><w:r><w:t>cut"#;
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    writer
        .add_part(
            &PartName::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            body.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let dir = std::env::temp_dir().join(format!("office_oxide_cliopts_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join(name);
    std::fs::write(&path, writer.finish().unwrap().into_inner()).unwrap();
    path
}

fn run(args: &[&str], path: &std::path::Path) -> Output {
    Command::new(env!("CARGO_BIN_EXE_office-oxide"))
        .args(args)
        .arg(path)
        .output()
        .unwrap()
}

fn stderr(out: &Output) -> String {
    String::from_utf8_lossy(&out.stderr).into_owned()
}

#[test]
fn test_an_incomplete_file_is_read_by_default_and_refused_with_strict() {
    let path = cut_off_docx("cut.docx");
    for sub in ["text", "markdown", "html", "ir", "info"] {
        let plain = run(&[sub], &path);
        assert!(plain.status.success(), "{sub}: {}", stderr(&plain));

        let strict = run(&["--strict", sub], &path);
        assert!(!strict.status.success(), "{sub} accepted an incomplete file");
        assert!(stderr(&strict).contains("incomplete"), "{sub}: {}", stderr(&strict));
        assert!(stderr(&strict).contains("word/document.xml"), "{sub}: {}", stderr(&strict));
    }
}

#[test]
fn test_info_says_what_is_missing() {
    let path = cut_off_docx("info.docx");
    let out = run(&["info"], &path);
    let text = String::from_utf8_lossy(&out.stdout);
    assert!(text.contains("Incomplete:") && text.contains("word/document.xml"), "{text}");
}

#[test]
fn test_a_limit_is_set_by_name_for_the_run() {
    let path = cut_off_docx("limit.docx");
    let refused = run(&["--limit", "part_bytes=64", "text"], &path);
    assert!(!refused.status.success());
    let allowed = run(&["--limit", "part_bytes=max", "text"], &path);
    assert!(allowed.status.success(), "{}", stderr(&allowed));
}

#[test]
fn test_a_bad_limit_is_a_usage_error_naming_the_choices() {
    let path = cut_off_docx("bad.docx");
    for bad in ["no_such_limit=1", "text_chars", "text_chars=lots"] {
        let out = run(&["--limit", bad, "text"], &path);
        assert_eq!(out.status.code(), Some(2), "{bad}");
        assert!(stderr(&out).contains("error:"), "{bad}: {}", stderr(&out));
    }
    let out = run(&["--limit", "no_such_limit=1", "text"], &path);
    assert!(stderr(&out).contains("text_chars"), "{}", stderr(&out));
}
