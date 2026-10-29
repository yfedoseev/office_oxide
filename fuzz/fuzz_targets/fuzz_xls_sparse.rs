#![no_main]

// The legacy `.xls` reader stores each row up to its last used column, so a
// few cells in far columns pad their rows with empty slots, charged to a
// per-sheet and a per-document allowance. Whole-file fuzzing rarely reaches
// that behind a valid CFB container and BIFF8 globals substream. This target
// builds a well-formed workbook in code from `(sheet, row, column)` triples the
// fuzzer chooses, under small allowances, and checks the padding stays within
// them. None may panic, overflow or hang.
#[path = "../../tests/common/mod.rs"]
mod common;

use std::io::Cursor;

use common::sparse_xls;
use libfuzzer_sys::fuzz_target;
use office_oxide::{Document, DocumentFormat};

// The padding allowances are shares of the memory budget (a 64th of it in
// cells for the document, a 128th for a sheet): 2,000 and 1,000 cells.
const MEMORY_BYTES: u64 = 128_000;
const DOCUMENT_SLOTS: usize = 2_000;

fuzz_target!(|data: &[u8]| {
    office_oxide::limits::set_max_memory_bytes(MEMORY_BYTES);
    let mut sheets: Vec<Vec<(u16, u16)>> = vec![Vec::new(); 4];
    for c in data.chunks_exact(4).take(2_000) {
        let sheet = usize::from(c[0] % 4);
        let row = u16::from(c[1]) | (u16::from(c[0] >> 6) << 8);
        sheets[sheet].push((row, u16::from(c[2]) | (u16::from(c[3] & 1) << 8)));
    }
    let bytes = sparse_xls(&sheets);
    if let Ok(doc) = Document::from_reader(Cursor::new(bytes), DocumentFormat::Xls) {
        let ir = doc.to_ir();
        let cells: usize = ir
            .sections
            .iter()
            .flat_map(|s| &s.elements)
            .filter_map(|e| match e {
                office_oxide::ir::Element::Table(t) => {
                    Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>())
                },
                _ => None,
            })
            .sum();
        // The cells that exist, plus the padding the allowances pay for.
        assert!(cells <= data.len() / 4 * 512 + DOCUMENT_SLOTS + 8 * 512);
    }
});
