#![no_main]

// The opt-in `default-password` feature decrypts what Office protects with its
// default password: the Standard and Agile encryption of an OOXML package (an
// `EncryptionInfo` and an `EncryptedPackage` stream in a compound file) and the
// RC4 workbook stream of an `.xls` (a `FILEPASS` record). Whole-file fuzzing
// seldom builds the container around the encryption structures, so this target
// builds a valid compound file in code and hands the fuzzer's bytes in as the
// streams / the record. None may panic, overflow or hang; a header the reader
// cannot use must be refused as before.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::{CfbNode, cfb_tree};
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

fn record(rt: u16, data: &[u8]) -> Vec<u8> {
    let data = &data[..data.len().min(8224)];
    let mut out = rt.to_le_bytes().to_vec();
    out.extend_from_slice(&(data.len() as u16).to_le_bytes());
    out.extend_from_slice(data);
    out
}

fuzz_target!(|data: &[u8]| {
    let Some((&which, rest)) = data.split_first() else {
        return;
    };
    match which % 3 {
        // An encrypted OOXML package: the fuzzer's bytes split into the two streams.
        0 | 1 => {
            let cut = usize::from(rest.first().copied().unwrap_or(0)) * 8;
            let (info, package) = rest.split_at(cut.min(rest.len()));
            let file = cfb_tree(vec![
                CfbNode::Stream("EncryptionInfo".into(), info.to_vec()),
                CfbNode::Stream("EncryptedPackage".into(), package.to_vec()),
            ]);
            let format = if which % 3 == 0 { DocumentFormat::Docx } else { DocumentFormat::Xlsx };
            let _ = Document::from_reader(Cursor::new(file), format);
        },
        // An `.xls` whose FILEPASS record and the records after it are the fuzzer's.
        _ => {
            let cut = usize::from(rest.first().copied().unwrap_or(0)) % 64;
            let (filepass, tail) = rest.split_at(cut.min(rest.len()));
            let mut stream = record(0x0809, &[0, 6, 5, 0, 0, 0, 0, 0]);
            stream.extend(record(0x002F, filepass));
            stream.extend_from_slice(tail);
            let file = cfb_tree(vec![CfbNode::Stream("Workbook".into(), stream)]);
            let _ = Document::from_reader(Cursor::new(file), DocumentFormat::Xls);
        },
    }
});
