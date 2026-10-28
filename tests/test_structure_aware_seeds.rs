//! Seed corpus for the structure-aware fuzz targets `fuzz_ppt_slide_list` and
//! `fuzz_xls_sparse`: the same builders, driven by a fixed generator, so the
//! decisions the targets exist to reach (persist-directory aliasing,
//! container claims, padding allowances) are exercised by `cargo test` too.

mod common;

use std::io::Cursor;

use common::LimitScope;
use common::ppt::aliasing_deck;
use common::sparse_xls;
use office_oxide::ir::Element;
use office_oxide::limits::{
    max_grid_padding_cells, max_sheet_grid_padding_cells, set_max_grid_padding_cells,
    set_max_sheet_grid_padding_cells,
};
use office_oxide::{Document, DocumentFormat};

fn generator(seed: u64) -> impl FnMut() -> u64 {
    let mut state = seed;
    move || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    }
}

#[test]
fn test_ppt_entries_mapped_anywhere_never_multiply_the_text() {
    let _defaults = common::default_limits();
    let mut next = generator(0x9E37_79B9_7F4A_7C15);
    for round in 0..80 {
        let slides = 1 + (next() % 8) as usize;
        let texts: Vec<String> = (0..slides)
            .map(|i| format!("slide {i} {}", "t".repeat(200)))
            .collect();
        let entries = (next() % 300) as usize;
        let targets: Vec<u8> = (0..entries).map(|_| next() as u8).collect();
        let nest = next().is_multiple_of(2);
        let bytes = aliasing_deck(&texts, nest, &targets);
        let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt)
            .unwrap_or_else(|e| panic!("round {round}: {e}"));
        let ir = doc.to_ir();
        assert!(
            ir.plain_text().len() < 64 * slides * 400 + 4_096,
            "round {round}: {} slides, {entries} entries, nest={nest}: {} bytes",
            slides,
            ir.plain_text().len()
        );
        // Every entry is a slide, left out, or one of the two empty
        // placeholders (an unresolvable id, and an offset that is no slide).
        assert!(ir.sections.len() <= slides + 2, "round {round}: {} sections", ir.sections.len());
    }
}

#[test]
fn test_ppt_overlapping_containers_are_read_once_and_reported() {
    let _defaults = common::default_limits();
    let texts: Vec<String> = (0..4).map(|i| format!("slide {i}")).collect();
    // Four entries, one per container; the last three sit inside the first.
    let bytes = aliasing_deck(&texts, true, &[0, 1, 2, 3]);
    let doc = Document::from_reader(Cursor::new(bytes), DocumentFormat::Ppt).expect("opens");
    let ir = doc.to_ir();
    assert_eq!(ir.sections.len(), 1, "the nested containers are left out, not kept empty");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("3 slide or notes containers overlap")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_xls_sparse_sheets_stay_within_the_padding_allowances() {
    let mut limits = LimitScope::new();
    limits.set(max_grid_padding_cells, set_max_grid_padding_cells, 3_000);
    limits.set(max_sheet_grid_padding_cells, set_max_sheet_grid_padding_cells, 1_000);
    let mut next = generator(0xD1B5_4A32_D192_ED03);
    for round in 0..60 {
        let mut sheets: Vec<Vec<(u16, u16)>> = vec![Vec::new(); 4];
        let cells = (next() % 400) as usize;
        for _ in 0..cells {
            let sheet = (next() % 4) as usize;
            sheets[sheet].push(((next() % 300) as u16, (next() % 256) as u16));
        }
        let doc = Document::from_reader(Cursor::new(sparse_xls(&sheets)), DocumentFormat::Xls)
            .unwrap_or_else(|e| panic!("round {round}: {e}"));
        let ir = doc.to_ir();
        let table_cells: usize = ir
            .sections
            .iter()
            .flat_map(|s| &s.elements)
            .filter_map(|e| match e {
                Element::Table(t) => Some(t.rows.iter().map(|r| r.cells.len()).sum::<usize>()),
                _ => None,
            })
            .sum();
        // The cells that exist, the padding the allowances pay for, and a
        // row of slack for the one that is refused.
        assert!(
            table_cells <= cells + 3_000 + 4 * 256,
            "round {round}: {table_cells} table cells from {cells} cells"
        );
    }
}
