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

/// Four prose paragraphs of 99 `x`, with 300 one-character runs over the start
/// of the text that all carry `sprmCFVanish` (0x083C) = 1: hidden text.
fn hidden_text_doc(piled_pieces: usize) -> office_oxide::Document {
    let text: &'static str = Box::leak("x".repeat(99).into_boxed_str());
    let paras: Vec<Para> = (0..4)
        .map(|_| Para {
            text,
            terminator: '\r',
            grpprl: Vec::new(),
        })
        .collect();
    let tweaks = FibTweaks {
        piled_pieces,
        chpx_runs: 300,
        chpx_grpprl: vec![0x3C, 0x08, 0x01],
        ..Default::default()
    };
    open_doc(&build_doc_full(&paras, &Subdocs::default(), tweaks))
}

#[test]
fn test_word97_character_runs_over_piled_pieces_are_not_read_and_the_text_is_kept() {
    // 300 character runs over 800 pieces on the same bytes is runs x pieces
    // work however it is searched. The runs are not read: the text stays whole
    // (nothing is hidden or deleted by them) and the reason is reported.
    let doc = hidden_text_doc(800);
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("character formatting")), "{warnings:?}");
    assert_eq!(doc.plain_text().matches('x').count(), 4 * 99);
}

#[test]
fn test_word97_hidden_text_is_still_excluded_when_the_character_runs_are_cheap() {
    // The same runs over an ordinary piece table are read: the hidden text is
    // left out, with no warning. This is what the guard above gives up.
    let doc = hidden_text_doc(1);
    let warnings = doc.to_ir().metadata.warnings;
    assert!(!warnings.iter().any(|w| w.contains("character formatting")), "{warnings:?}");
    assert!(doc.plain_text().matches('x').count() < 4 * 99);
}
