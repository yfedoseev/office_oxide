//! The binary's behaviour when its reader goes away.

use std::io::{Read, Write};
use std::process::{Command, Stdio};

/// Build the smallest document whose `text` output is longer than a pipe
/// buffer, so the writer is still writing when the reader closes.
fn big_docx() -> Vec<u8> {
    let mut w = office_oxide::docx::write::DocxWriter::new();
    for i in 0..40_000 {
        w.add_paragraph(&format!(
            "Paragraph {i} with enough words to fill a pipe buffer many times over."
        ));
    }
    let mut out = Vec::new();
    w.write_to(&mut std::io::Cursor::new(&mut out)).unwrap();
    out
}

/// `office-oxide text f.docx | head` used to panic with "failed printing
/// to stdout: Broken pipe" and exit 101, because Rust starts with
/// `SIGPIPE` ignored. A closed pipe now ends the process the way it ends
/// `cat`: quietly, by the signal.
#[test]
#[cfg(unix)]
fn test_closed_stdout_pipe_ends_quietly_without_a_panic() {
    let dir = std::env::temp_dir().join(format!("office_oxide_pipe_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("big.docx");
    std::fs::write(&path, big_docx()).unwrap();

    let mut child = Command::new(env!("CARGO_BIN_EXE_office-oxide"))
        .arg("text")
        .arg(&path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    // Read a little, then close the read end while the child still has
    // megabytes to write.
    let mut stdout = child.stdout.take().unwrap();
    let mut first = [0u8; 64];
    let n = stdout.read(&mut first).unwrap();
    assert!(n > 0);
    drop(stdout);
    let status = child.wait().unwrap();
    let mut stderr = String::new();
    child
        .stderr
        .take()
        .unwrap()
        .read_to_string(&mut stderr)
        .unwrap();
    std::fs::remove_dir_all(&dir).ok();

    assert!(!stderr.contains("panicked"), "stderr: {stderr}");
    assert_ne!(status.code(), Some(101), "a panic exit: {stderr}");
    let _ = std::io::stderr().flush();
}
