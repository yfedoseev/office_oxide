//! A Word 97 file whose pieces are piled on the same bytes must not make the
//! paragraph and character-run builders do runs x pieces work. Built in code.

mod common;

use common::{FibTweaks, Para, Subdocs, build_doc_full, cell_grpprl, open_doc, row_grpprl};
use office_oxide::ir::Element;

fn table_paras(rows: usize) -> Vec<Para> {
    let mut paras = Vec::new();
    for _ in 0..rows {
        paras.push(Para {
            text: "A",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        });
        paras.push(Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl(&[0, 1000], &[0]),
        });
    }
    paras
}

fn tables(ir: &office_oxide::ir::DocumentIR) -> usize {
    ir.sections
        .iter()
        .flat_map(|s| s.elements.iter())
        .filter(|e| matches!(e, Element::Table(_)))
        .count()
}

#[test]
fn test_word97_pieces_piled_on_the_same_bytes_do_not_get_a_structured_view() {
    let tweaks = FibTweaks {
        piled_pieces: 800,
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&table_paras(40), &Subdocs::default(), tweaks));
    let ir = doc.to_ir();
    assert_eq!(tables(&ir), 0, "this layout must not be walked structurally");
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("pieces")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(doc.plain_text().contains('A'), "the text survives");
}

#[test]
fn test_word97_a_single_piece_keeps_its_tables() {
    let doc = open_doc(&build_doc_full(&table_paras(3), &Subdocs::default(), FibTweaks::default()));
    assert_eq!(tables(&doc.to_ir()), 1);
}

#[test]
fn test_word97_pieces_that_share_a_few_bytes_are_still_walked() {
    let tweaks = FibTweaks {
        piled_pieces: 2,
        ..Default::default()
    };
    let doc = open_doc(&build_doc_full(&table_paras(3), &Subdocs::default(), tweaks));
    assert_eq!(tables(&doc.to_ir()), 1);
}
