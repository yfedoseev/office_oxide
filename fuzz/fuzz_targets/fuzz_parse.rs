#![no_main]

use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

// Office documents are untrusted input: OOXML (Docx/Xlsx/Pptx) is a zip of XML,
// the legacy formats (Doc/Xls/Ppt) are CFB compound files. Feed arbitrary bytes
// to every format's parser — none may panic, overflow, or hang; malformed input
// must surface as `Err`.
fuzz_target!(|data: &[u8]| {
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
        DocumentFormat::Doc,
        DocumentFormat::Xls,
        DocumentFormat::Ppt,
    ] {
        let _ = Document::from_reader(Cursor::new(data.to_vec()), format);
    }
});
