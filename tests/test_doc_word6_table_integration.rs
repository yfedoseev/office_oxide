//! Tables in Word 6.0/95 `.doc` files.
//!
//! The Word 6/95 reader was text-only: it read no paragraph properties, so the
//! `\x07` cell and row marks of a table came out as tabs and line breaks inside
//! ordinary paragraphs and every table was lost. Its paragraph properties are
//! in a PAPX FKP in the `WordDocument` stream, in the Word 6 encoding (one-byte
//! sprm numbers, 7-byte FKP entries, 2-byte bin-table entries).
//!
//! Documents are built in code with the synthetic writer in `tests/common`.

mod common;

use common::{build_word6_doc_with_papx, open_doc, w6_in_table, w6_row_end};
use office_oxide::ir::{DocumentIR, Element, Table, TableCell};

const WIDENT: u16 = 0xA5DC; // Word 6.0/95 for Windows

fn tables(ir: &DocumentIR) -> Vec<&Table> {
    ir.sections
        .iter()
        .flat_map(|s| &s.elements)
        .filter_map(|e| match e {
            Element::Table(t) => Some(t),
            _ => None,
        })
        .collect()
}

fn text_of(cell: &TableCell) -> String {
    use office_oxide::ir::InlineContent;
    cell.content
        .iter()
        .filter_map(|e| match e {
            Element::Paragraph(p) => Some(
                p.content
                    .iter()
                    .filter_map(|c| match c {
                        InlineContent::Text(t) => Some(t.text.as_str()),
                        _ => None,
                    })
                    .collect::<String>(),
            ),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join("|")
}

/// A | B ; C | D, then a prose paragraph.
fn two_by_two(header_first_row: bool) -> Vec<u8> {
    let cell = |t: &'static [u8]| (t, w6_in_table());
    build_word6_doc_with_papx(
        WIDENT,
        &[
            cell(b"A\x07"),
            cell(b"B\x07"),
            (b"\x07", w6_row_end(2, header_first_row)),
            cell(b"C\x07"),
            cell(b"D\x07"),
            (b"\x07", w6_row_end(2, false)),
            (b"After the table.\r", vec![]),
        ],
    )
}

#[test]
fn test_word6_table_is_a_table_with_its_rows_and_cells() {
    let doc = open_doc(&two_by_two(false));
    let ir = doc.to_ir();
    let tables = tables(&ir);
    assert_eq!(tables.len(), 1, "one table, not tab-separated lines");
    let rows: Vec<Vec<String>> = tables[0]
        .rows
        .iter()
        .map(|r| r.cells.iter().map(text_of).collect())
        .collect();
    assert_eq!(rows, vec![vec!["A", "B"], vec!["C", "D"]]);
    assert!(!tables[0].rows[0].is_header);
    // The paragraph after the table is still prose.
    assert!(doc.plain_text().contains("After the table."));
}

#[test]
fn test_word6_header_row_flag_reaches_the_row() {
    let doc = open_doc(&two_by_two(true));
    let ir = doc.to_ir();
    let t = tables(&ir)[0];
    assert!(t.rows[0].is_header && t.rows[0].repeat_as_header);
    assert!(!t.rows[1].is_header);
}

#[test]
fn test_word6_without_a_papx_still_opens_with_its_text() {
    // No bin table at all (`lcbPlcfbtePapx` 0): the text-only path.
    let bytes = common::build_word6_doc(WIDENT, b"Plain text.\rSecond paragraph.\r");
    let doc = open_doc(&bytes);
    assert!(doc.plain_text().contains("Second paragraph."));
    assert!(tables(&doc.to_ir()).is_empty());
}

#[test]
fn test_word6_unreadable_papx_is_reported_not_fatal() {
    // A bin table that points outside the stream: the text survives and the
    // lost structure is reported, not silently dropped.
    let doc = common::open_word6_with_bad_papx_pointer(WIDENT);
    assert!(doc.plain_text().contains('B'), "the text survives");
    let warnings = doc.to_ir().metadata.warnings;
    assert!(warnings.iter().any(|w| w.contains("PAPX")), "{warnings:?}");
}

#[test]
fn test_word6_unknown_sprm_in_a_table_cell_is_reported_and_the_text_kept() {
    // An opcode the table does not know cannot be stepped over: the list is
    // read as far as it goes, the cell keeps its text, and the loss is reported.
    let doc = open_doc(&build_word6_doc_with_papx(
        WIDENT,
        &[
            (&b"Odd cell\x07"[..], vec![24, 1, 250, 1, 2, 3]),
            (&b"\x07"[..], w6_row_end(1, false)),
        ],
    ));
    let ir = doc.to_ir();
    assert_eq!(text_of(&tables(&ir)[0].rows[0].cells[0]), "Odd cell");
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("could not be read")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_word6_text_after_the_last_papx_run_is_kept_with_default_properties() {
    // The FKP page describes the table but stops before the text that follows
    // it. Text no run covers has default paragraph properties; it is not
    // dropped.
    let cell = |t: &'static [u8]| (t, w6_in_table());
    let paras = [
        cell(b"A\x07"),
        cell(b"B\x07"),
        (&b"\x07"[..], w6_row_end(2, false)),
        (&b"Tail one.\r"[..], vec![]),
        (&b"Tail two.\r"[..], vec![]),
    ];
    let doc = open_doc(&common::build_word6_doc_with_papx_covering(WIDENT, &paras, 3));
    let ir = doc.to_ir();
    assert_eq!(tables(&ir).len(), 1, "the table the FKP does describe is still a table");
    let json = serde_json::to_string(&ir).unwrap();
    assert!(
        json.contains("Tail one.") && json.contains("Tail two."),
        "text after the last run was dropped: {json}"
    );
}

#[test]
fn test_word6_text_in_a_row_end_paragraph_is_kept_not_dropped() {
    // A row-end paragraph is empty in a well-formed file. When a damaged
    // property list marks a paragraph that holds text as the row end, the text
    // is a cell's content and stays in the table.
    let cell = |t: &'static [u8]| (t, w6_in_table());
    let doc = open_doc(&build_word6_doc_with_papx(
        WIDENT,
        &[
            cell(b"A\x07"),
            (&b"stray words\x07"[..], w6_row_end(1, false)),
            (&b"After.\r"[..], vec![]),
        ],
    ));
    let json = serde_json::to_string(&doc.to_ir()).unwrap();
    assert!(json.contains("stray words"), "row-end paragraph text was dropped: {json}");
}

#[test]
fn test_word97_text_in_a_row_end_paragraph_is_kept_not_dropped() {
    use common::{Para, build_doc, cell_grpprl, row_grpprl};
    let paras = vec![
        Para {
            text: "A",
            terminator: '\u{7}',
            grpprl: cell_grpprl(),
        },
        Para {
            text: "stray words",
            terminator: '\u{7}',
            grpprl: row_grpprl(&[0, 1000], &[0]),
        },
    ];
    let doc = open_doc(&build_doc(&paras));
    let json = serde_json::to_string(&doc.to_ir()).unwrap();
    assert!(json.contains("stray words"), "row-end paragraph text was dropped: {json}");
}

#[test]
fn test_word6_file_without_a_table_keeps_the_text_only_view() {
    // Paragraph properties are read for the tables they describe. A file with
    // none is left exactly as the text-only reader produced it: no stray line
    // break where a section break (0x0C) sat, no change of any kind.
    let text: &[u8] = b"One.\r\x0cTwo, after a section break.\r";
    let plain = open_doc(&common::build_word6_doc(WIDENT, text)).to_ir();
    let with_papx = open_doc(&build_word6_doc_with_papx(
        WIDENT,
        &[(&text[..5], vec![]), (&text[5..], vec![])],
    ))
    .to_ir();
    assert_eq!(
        serde_json::to_string(&plain).unwrap(),
        serde_json::to_string(&with_papx).unwrap(),
        "a Word 6 file with no table must not change"
    );
}

fn paragraph_texts(ir: &DocumentIR) -> Vec<String> {
    use office_oxide::ir::InlineContent;
    fn collect(els: &[Element], out: &mut Vec<String>) {
        for e in els {
            match e {
                Element::Paragraph(p) => out.push(
                    p.content
                        .iter()
                        .filter_map(|c| match c {
                            InlineContent::Text(t) => Some(t.text.as_str()),
                            InlineContent::LineBreak => Some("\n"),
                            _ => None,
                        })
                        .collect(),
                ),
                Element::Table(t) => {
                    for c in t.rows.iter().flat_map(|r| &r.cells) {
                        collect(&c.content, out);
                    }
                },
                _ => {},
            }
        }
    }
    let mut out = Vec::new();
    for s in &ir.sections {
        collect(&s.elements, &mut out);
    }
    out
}

fn row_of(a: &'static [u8], b: &'static [u8]) -> Vec<(&'static [u8], Vec<u8>)> {
    vec![
        (a, w6_in_table()),
        (b, w6_in_table()),
        (&b"\x07"[..], w6_row_end(2, false)),
    ]
}

#[test]
fn test_word6_section_break_in_a_file_with_a_table_still_ends_the_paragraph() {
    // A section break (0x0C) replaces the paragraph mark of a section's last
    // paragraph. Without the section table it was read as a line break inside
    // one paragraph, merging the two on either side of it.
    let mut paras = row_of(b"A\x07", b"B\x07");
    paras.push((&b"End of one.\x0c"[..], vec![]));
    paras.push((&b"Start of two.\r"[..], vec![]));
    let mut spec = common::Word6Spec::new(vec![paras]);
    // "A\a" "B\a" "\a" = 5 characters, "End of one.\x0c" = 12, "Start of two.\r" = 14.
    spec.section_ends = vec![17, 31];
    let ir = open_doc(&spec.build()).to_ir();
    let texts = paragraph_texts(&ir);
    assert!(
        !texts
            .iter()
            .any(|t| t.contains("End of one.") && t.contains("Start of two.")),
        "the paragraphs on either side of a section break were merged: {texts:?}"
    );
}

#[test]
fn test_word6_bin_table_that_lists_fewer_pages_than_exist_does_not_lose_a_table() {
    // `cpnBtePap` says how many PAPX pages there are; the bin table can list
    // fewer. The pages after the last listed one follow it in order.
    let page0 = vec![(&b"Intro.\r"[..], vec![])];
    let page1 = row_of(b"A\x07", b"B\x07");
    let mut spec = common::Word6Spec::new(vec![page0, page1]);
    spec.listed_pages = 1;
    let ir = open_doc(&spec.build()).to_ir();
    assert_eq!(tables(&ir).len(), 1, "the table on the unlisted page was lost");
}

#[test]
fn test_word6_table_spread_over_two_fkp_pages_is_one_table() {
    let mut row1 = row_of(b"A\x07", b"B\x07");
    let row2 = row_of(b"C\x07", b"D\x07");
    let page1 = row1.split_off(0);
    let spec = common::Word6Spec::new(vec![page1, row2]);
    let ir = open_doc(&spec.build()).to_ir();
    let tables = tables(&ir);
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].rows.len(), 2);
}

#[test]
fn test_word6_page_named_twice_in_the_bin_table_is_read_once() {
    let mut spec = common::Word6Spec::new(vec![row_of(b"A\x07", b"B\x07")]);
    spec.duplicate_listed = true;
    let ir = open_doc(&spec.build()).to_ir();
    let tables = tables(&ir);
    assert_eq!((tables.len(), tables[0].rows.len()), (1, 1), "rows were duplicated");
}

#[test]
fn test_word6_table_in_every_variant_of_the_family() {
    // The Macintosh builds share the layout; `fExtChar` stores two-byte text;
    // a fast-saved file describes its text with a CLX.
    for (name, wident, ext_char, complex) in [
        ("mac 0xA697", 0xA697u16, false, false),
        ("mac 0xA699", 0xA699, false, false),
        ("fExtChar", 0xA5DC, true, false),
        ("fast-saved", 0xA5DC, false, true),
        ("fast-saved fExtChar", 0xA5DC, true, true),
    ] {
        let mut spec = common::Word6Spec::new(vec![row_of(b"A\x07", b"B\x07")]);
        spec.wident = wident;
        spec.ext_char = ext_char;
        spec.complex = complex;
        let ir = open_doc(&spec.build()).to_ir();
        let tables = tables(&ir);
        assert_eq!(tables.len(), 1, "{name}: no table");
        assert_eq!(text_of(&tables[0].rows[0].cells[1]), "B", "{name}");
    }
}

#[test]
fn test_word6_row_definition_that_cannot_be_used_is_reported() {
    // More columns than a table has: the row still ends, but its column
    // structure is gone, and that is said.
    let doc = open_doc(&build_word6_doc_with_papx(
        WIDENT,
        &[
            (&b"A\x07"[..], w6_in_table()),
            (&b"\x07"[..], w6_row_end(64, false)),
        ],
    ));
    let ir = doc.to_ir();
    assert_eq!(tables(&ir).len(), 1);
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("row definition")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// `pages` pages of table rows: `rows_per_page` one-cell rows on each.
fn many_runs(pages: usize, rows_per_page: usize) -> Vec<Vec<(&'static [u8], Vec<u8>)>> {
    (0..pages)
        .map(|_| {
            (0..rows_per_page)
                .flat_map(|_| {
                    [
                        (&b"A\x07"[..], w6_in_table()),
                        (&b"\x07"[..], w6_row_end(1, false)),
                    ]
                })
                .collect()
        })
        .collect()
}

#[test]
fn test_word6_pieces_piled_on_the_same_bytes_do_not_get_a_structured_view() {
    // The paragraph builder intersects every property run with every piece it
    // touches. Pieces of a fast-saved file that all start at the same bytes
    // multiply against the runs there. A real file reuses a byte or two; a pile
    // of hundreds is not an edit history. For one, the text-only view stands
    // and the reason is reported.
    let mut spec = common::Word6Spec::new(many_runs(8, 3));
    // 600 pieces of about 70 bytes each, all starting where the table is.
    let tail = vec![b'x'; 44_000];
    spec.tail = &tail;
    spec.complex = true;
    spec.overlapping_pieces = 600;
    let doc = open_doc(&spec.build());
    let ir = doc.to_ir();
    assert!(tables(&ir).is_empty(), "this layout must not be walked structurally");
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("pieces")),
        "{:?}",
        ir.metadata.warnings
    );
    assert!(doc.plain_text().contains('A'), "the text survives");
}

#[test]
fn test_word6_one_long_piece_with_many_small_ones_inside_is_not_walked_structurally() {
    // Only two pieces are ever over a byte here, but the builder starts its
    // search for a run at the first span whose running maximum end is past the
    // run, and the long piece keeps that maximum high: every run revisits every
    // small piece. What is bounded is the work the builder does, not the depth.
    // 700 bytes of text first, so 600 distinct one-byte pieces can sit inside
    // the long one there, and the table's runs come after them.
    let pad = vec![b'x'; 700];
    let mut pages = many_runs(8, 3);
    pages[0].insert(0, (&pad[..], vec![]));
    let mut spec = common::Word6Spec::new(pages);
    spec.complex = true;
    spec.nested_pieces = 600;
    let doc = open_doc(&spec.build());
    let ir = doc.to_ir();
    assert!(tables(&ir).is_empty(), "this layout must not be walked structurally");
    assert!(
        ir.metadata.warnings.iter().any(|w| w.contains("pieces")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_word6_pieces_that_share_a_few_bytes_are_still_walked() {
    // A guard for the check above. Real fast-saved files do share bytes between
    // pieces (a piece can sit inside another); that must not cost the tables.
    let page: Vec<(&'static [u8], Vec<u8>)> =
        vec![(b"A\x07", w6_in_table()), (b"\x07", w6_row_end(1, false))];
    let mut spec = common::Word6Spec::new(vec![page]);
    spec.complex = true;
    spec.overlapping_pieces = 2;
    let ir = open_doc(&spec.build()).to_ir();
    assert_eq!(tables(&ir).len(), 1);
}

/// A one-row, one-cell table, with `prefix` ahead of each paragraph's own table
/// sprms (a property list as Word writes it, other sprms first).
fn one_cell_table(prefix: &[u8]) -> Vec<(&'static [u8], Vec<u8>)> {
    let with = |tail: Vec<u8>| [prefix, &tail].concat();
    vec![
        (b"Cell text\x07", with(w6_in_table())),
        (b"\x07", with(w6_row_end(1, false))),
    ]
}

#[test]
fn test_word6_piece_modifier_is_not_read_as_a_word97_property_list() {
    // A fast-saved file's pieces can name a property modifier. In Word 6 it is
    // a list of one-byte sprms; read as Word 97's two-byte opcodes, these bytes
    // are `sprmCFRMarkDel` (0x0800) = 1 and hide the text of every piece.
    let mut spec = common::Word6Spec::new(vec![one_cell_table(&[])]);
    spec.complex = true;
    spec.piece_prc = vec![0x00, 0x08, 0x01];
    let ir = open_doc(&spec.build()).to_ir();
    let t = tables(&ir);
    assert_eq!(t.len(), 1);
    assert_eq!(text_of(&t[0].rows[0].cells[0]), "Cell text");
}

#[test]
fn test_word6_macintosh_font_sprms_are_two_byte_operands_not_lengths() {
    // On the Macintosh variants sprms 111-113 carry a 2-byte operand. Read as a
    // length byte, 0x7F skips past the rest of the list and the table sprms
    // that follow are never seen.
    let page = one_cell_table(&[111, 0x7F, 0x00]);
    for wident in [0xA697u16, 0xA698, 0xA699] {
        let mut spec = common::Word6Spec::new(vec![page.clone()]);
        spec.wident = wident;
        let ir = open_doc(&spec.build()).to_ir();
        assert_eq!(tables(&ir).len(), 1, "wIdent {wident:#X}");
    }
}

#[test]
fn test_word6_pages_past_the_file_are_reported_as_lost_tables() {
    // `cpnBtePap` claims far more pages than the file holds. The warning must
    // say what is lost: for a Word 6/95 file the paragraphs have no structure
    // to lose, the tables are what cannot be found.
    let mut spec = common::Word6Spec::new(vec![one_cell_table(&[])]);
    spec.claimed_pages = Some(300);
    let ir = open_doc(&spec.build()).to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("lie outside the file") && w.contains("tables")),
        "{:?}",
        ir.metadata.warnings
    );
}
