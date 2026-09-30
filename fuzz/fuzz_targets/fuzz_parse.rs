#![no_main]

use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

// Office documents are untrusted input: OOXML (Docx/Xlsx/Pptx) is a zip of XML,
// the legacy formats (Doc/Xls/Ppt) are CFB compound files. Feed arbitrary bytes
// to every format's parser — none may panic, overflow, or hang; malformed input
// must surface as `Err`. Parsing is only half the surface: historical crashes
// lived in the renderers and the IR converter, so a successful parse is
// exercised through every output path as well.
fuzz_target!(|data: &[u8]| {
    for format in [
        DocumentFormat::Docx,
        DocumentFormat::Xlsx,
        DocumentFormat::Pptx,
        DocumentFormat::Doc,
        DocumentFormat::Xls,
        DocumentFormat::Ppt,
    ] {
        if let Ok(doc) = Document::from_reader(Cursor::new(data.to_vec()), format) {
            let _ = doc.plain_text();
            let _ = doc.to_markdown();
            let _ = doc.to_html();
            let _ = doc.to_ir();
        }
    }
});
