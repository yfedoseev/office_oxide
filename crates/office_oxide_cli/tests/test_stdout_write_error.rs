//! The binary's behaviour when stdout cannot be written.

use std::process::{Command, Stdio};

/// `office-oxide text f.docx > /dev/full` panicked in `print!` ("failed
/// printing to stdout: No space left on device") and exited 101. A write
/// error other than a closed pipe is an ordinary failure: `error: …` on
/// stderr and exit status 1, like every other failure this CLI reports.
#[test]
#[cfg(target_os = "linux")]
fn test_a_stdout_write_error_is_reported_not_panicked() {
    let dir = std::env::temp_dir().join(format!("office_oxide_devfull_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("doc.docx");
    let mut w = office_oxide::docx::write::DocxWriter::new();
    w.add_heading("Heading", 1);
    w.add_paragraph("Some body text.");
    w.save(&path).unwrap();

    for sub in ["text", "markdown", "html", "info", "ir"] {
        let full = std::fs::OpenOptions::new()
            .write(true)
            .open("/dev/full")
            .unwrap();
        let out = Command::new(env!("CARGO_BIN_EXE_office-oxide"))
            .arg(sub)
            .arg(&path)
            .stdout(Stdio::from(full))
            .stderr(Stdio::piped())
            .output()
            .unwrap();
        let stderr = String::from_utf8_lossy(&out.stderr);
        assert!(!stderr.contains("panicked"), "{sub}: stderr: {stderr}");
        assert_eq!(out.status.code(), Some(1), "{sub}: stderr: {stderr}");
        assert!(stderr.starts_with("error: "), "{sub}: stderr: {stderr}");
    }
    std::fs::remove_dir_all(&dir).ok();
}
