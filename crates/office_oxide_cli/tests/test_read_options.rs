//! `--include-hidden` and `--include-deleted` reach every command that reads
//! a document, and are off unless given.

use std::io::Cursor;
use std::process::Command;

use office_oxide::core::opc::{OpcWriter, PartName};
use office_oxide::core::relationships::rel_types;

fn fixture(name: &str) -> std::path::PathBuf {
    let body = r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
<w:r><w:t>shown </w:t></w:r>
<w:r><w:rPr><w:vanish/></w:rPr><w:t>hiddenword </w:t></w:r>
<w:del w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>deletedword</w:delText></w:r></w:del>
</w:p></w:body></w:document>"#;
    let mut writer = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
    let part = PartName::new("/word/document.xml").unwrap();
    writer
        .add_part(
            &part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            body.as_bytes(),
        )
        .unwrap();
    writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
    let dir = std::env::temp_dir().join(format!("office_oxide_readopts_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join(name);
    std::fs::write(&path, writer.finish().unwrap().into_inner()).unwrap();
    path
}

fn run(args: &[&str], path: &std::path::Path) -> String {
    let out = Command::new(env!("CARGO_BIN_EXE_office-oxide"))
        .args(args)
        .arg(path)
        .output()
        .unwrap();
    assert!(out.status.success(), "{args:?}: {}", String::from_utf8_lossy(&out.stderr));
    String::from_utf8_lossy(&out.stdout).into_owned()
}

#[test]
fn test_hidden_and_deleted_text_are_off_unless_asked_for() {
    let path = fixture("flags.docx");
    for sub in ["text", "markdown", "html", "ir"] {
        let plain = run(&[sub], &path);
        assert!(plain.contains("shown"), "{sub}");
        assert!(!plain.contains("hiddenword") && !plain.contains("deletedword"), "{sub}");

        let hidden = run(&[sub, "--include-hidden"], &path);
        assert!(hidden.contains("hiddenword") && !hidden.contains("deletedword"), "{sub}");

        let deleted = run(&[sub, "--include-deleted"], &path);
        assert!(!deleted.contains("hiddenword") && deleted.contains("deletedword"), "{sub}");

        let both = run(&["--include-hidden", "--include-deleted", sub], &path);
        assert!(both.contains("hiddenword") && both.contains("deletedword"), "{sub}");
    }
}
