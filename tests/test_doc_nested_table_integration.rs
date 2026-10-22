//! Legacy `.doc` tables that contain tables.
//!
//! A paragraph's `sprmPItap` (0x6649) is its table nesting depth: 1 is the
//! outermost table, and the cells of a table at depth N+1 sit inside a cell of
//! the table at depth N, before that cell's own end mark. The reader used to
//! build one flat table run, so a nested table's rows became rows of the outer
//! table and the outer row that held it was cut apart.
//!
//! Documents are built in code with the synthetic `.doc` writer.

mod common;

use common::{
    Para, build_doc, cell_grpprl_at, inner_cell_grpprl, inner_row_grpprl, open_doc, row_grpprl_at,
};
use office_oxide::ir::{DocumentIR, Element, Table, TableCell};

/// The last paragraph of a cell. Word ends an outermost cell with `\x07`, and a
/// cell of a nested table with an ordinary paragraph mark flagged as such.
fn cell(text: &'static str, itap: u8) -> Para {
    if itap <= 1 {
        Para {
            text,
            terminator: '\u{7}',
            grpprl: cell_grpprl_at(1),
        }
    } else {
        Para {
            text,
            terminator: '\r',
            grpprl: inner_cell_grpprl(itap),
        }
    }
}

/// A paragraph inside a cell that is not the cell's last (ends in `\r`).
fn line(text: &'static str, itap: u8) -> Para {
    Para {
        text,
        terminator: '\r',
        grpprl: cell_grpprl_at(itap),
    }
}

/// The paragraph that ends a row at depth `itap`, `cols` equal columns wide.
fn row_end(itap: u8, cols: usize) -> Para {
    let centers: Vec<i16> = (0..=cols).map(|c| (c * 1000) as i16).collect();
    let rgfs = vec![0; cols];
    if itap <= 1 {
        Para {
            text: "",
            terminator: '\u{7}',
            grpprl: row_grpprl_at(1, &centers, &rgfs),
        }
    } else {
        Para {
            text: "",
            terminator: '\r',
            grpprl: inner_row_grpprl(itap, &centers, &rgfs),
        }
    }
}

fn top_tables(ir: &DocumentIR) -> Vec<&Table> {
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

fn nested_table(cell: &TableCell) -> Option<&Table> {
    cell.content.iter().find_map(|e| match e {
        Element::Table(t) => Some(t),
        _ => None,
    })
}

/// How deep tables nest inside `t` (a table with no table in a cell is 1).
fn depth(t: &Table) -> usize {
    1 + t
        .rows
        .iter()
        .flat_map(|r| &r.cells)
        .filter_map(nested_table)
        .map(depth)
        .max()
        .unwrap_or(0)
}

/// All the words in a table tree.
fn words(t: &Table, out: &mut Vec<String>) {
    for c in t.rows.iter().flat_map(|r| &r.cells) {
        out.extend(
            text_of(c)
                .split(['|', ' '])
                .filter(|w| !w.is_empty())
                .map(str::to_owned),
        );
        if let Some(n) = nested_table(c) {
            words(n, out);
        }
    }
}

/// A | B | C ; D | (E, then a 2x2 table) | G ; H | I | J
fn table_with_a_nested_table() -> Vec<Para> {
    vec![
        cell("A", 1),
        cell("B", 1),
        cell("C", 1),
        row_end(1, 3),
        cell("D", 1),
        line("E", 1),
        cell("1", 2),
        cell("2", 2),
        row_end(2, 2),
        cell("3", 2),
        cell("4", 2),
        row_end(2, 2),
        cell("", 1), // end mark of the outer cell that holds the nested table
        cell("G", 1),
        row_end(1, 3),
        cell("H", 1),
        cell("I", 1),
        cell("J", 1),
        row_end(1, 3),
    ]
}

#[test]
fn test_doc_nested_table_stays_inside_the_cell_that_holds_it() {
    let doc = open_doc(&build_doc(&table_with_a_nested_table()));
    let ir = doc.to_ir();
    let tables = top_tables(&ir);

    assert_eq!(tables.len(), 1, "one top-level table");
    let outer = tables[0];
    assert_eq!(outer.rows.len(), 3, "the outer table keeps its 3 rows");
    assert!(outer.rows.iter().all(|r| r.cells.len() == 3), "and 3 cells in each");
    assert_eq!(text_of(&outer.rows[0].cells[0]), "A");
    assert_eq!(text_of(&outer.rows[1].cells[2]), "G");
    assert_eq!(text_of(&outer.rows[2].cells[1]), "I");

    let e = &outer.rows[1].cells[1];
    assert_eq!(text_of(e), "E", "the text before the nested table is the cell's own");
    let inner = nested_table(e).expect("the nested table is in the cell's content");
    let rows: Vec<Vec<String>> = inner
        .rows
        .iter()
        .map(|r| r.cells.iter().map(text_of).collect())
        .collect();
    assert_eq!(rows, vec![vec!["1", "2"], vec!["3", "4"]]);

    assert!(
        !ir.metadata.warnings.iter().any(|w| w.contains("flattened")),
        "nothing is flattened any more: {:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_nested_table_header_flag_belongs_to_the_nested_table() {
    // `sprmTTableHeader` on the nested table's first row marks that row of the
    // nested table, and nothing in the outer table.
    let mut paras = table_with_a_nested_table();
    paras[8].grpprl.extend_from_slice(&[0x04, 0x34, 0x01]);
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let outer = top_tables(&ir)[0];
    assert!(outer.rows.iter().all(|r| !r.is_header), "no outer row is a header");
    let inner = nested_table(&outer.rows[1].cells[1]).unwrap();
    assert!(inner.rows[0].is_header && inner.rows[0].repeat_as_header);
    assert!(!inner.rows[1].is_header);
}

#[test]
fn test_doc_two_nested_tables_in_one_cell_stay_separate() {
    let paras = vec![
        line("top", 1),
        cell("a", 2),
        row_end(2, 1),
        line("between", 1),
        cell("b", 2),
        row_end(2, 1),
        cell("", 1),
        row_end(1, 1),
    ];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let outer = top_tables(&ir)[0];
    let content = &outer.rows[0].cells[0].content;
    let kinds: Vec<&str> = content
        .iter()
        .map(|e| match e {
            Element::Paragraph(_) => "p",
            Element::Table(_) => "t",
            _ => "?",
        })
        .collect();
    assert_eq!(kinds, vec!["p", "t", "p", "t"], "text, table, text, table");
}

#[test]
fn test_doc_nested_table_cut_off_by_the_end_of_the_text_is_kept() {
    let paras = vec![line("outer text", 1), cell("inner text", 2)];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let mut all = Vec::new();
    for t in top_tables(&ir) {
        words(t, &mut all);
    }
    assert!(
        all.contains(&"outer".to_owned()) && all.contains(&"inner".to_owned()),
        "{all:?}"
    );
}

#[test]
fn test_doc_a_cell_may_begin_with_a_table_that_itself_begins_with_a_table() {
    // The first cell of the depth-1 table starts with a depth-2 table, whose
    // first cell starts with a depth-3 table: the depth-3 paragraphs arrive
    // before any paragraph of the tables around them. The structure is kept,
    // and nothing is reported -- this is an ordinary shape, not a damaged one.
    let paras = vec![
        cell("deepest", 3),
        row_end(3, 1),
        cell("", 2),
        row_end(2, 1),
        cell("", 1),
        row_end(1, 1),
    ];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let tables = top_tables(&ir);
    assert_eq!(tables.len(), 1);
    assert_eq!(depth(tables[0]), 3);
    let mut all = Vec::new();
    words(tables[0], &mut all);
    assert!(all.contains(&"deepest".to_owned()), "{all:?}");
    assert!(
        !ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("nesting depth")),
        "{:?}",
        ir.metadata.warnings
    );
}

#[test]
fn test_doc_table_nesting_deeper_than_word_allows_is_bounded() {
    // 20 tables inside one another. Word nests at most 15; the rest is folded
    // into the 15th rather than growing the tree (or the stack) without bound.
    let mut paras = Vec::new();
    for d in 1..=20u8 {
        let t: &'static str = Box::leak(format!("L{d}").into_boxed_str());
        paras.push(line(t, d));
    }
    for d in (1..=20u8).rev() {
        paras.push(cell("", d));
        paras.push(row_end(d, 1));
    }
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let tables = top_tables(&ir);
    assert_eq!(tables.len(), 1);
    assert!(depth(tables[0]) <= 15, "tables nest {} deep", depth(tables[0]));
    let mut all = Vec::new();
    words(tables[0], &mut all);
    for d in 1..=20 {
        assert!(all.contains(&format!("L{d}")), "L{d} lost: {all:?}");
    }
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("nesting depth"))
    );
}

#[test]
fn test_doc_nested_table_renders_in_markdown_and_html() {
    // Distinctive words: a one-letter probe is found in any markup.
    let doc = open_doc(&build_doc(&nested_table_with_distinctive_words()));
    let md = doc.to_markdown();
    let html = doc.to_html();
    for out in [&md, &html] {
        for w in [
            "Dleft", "Emid", "Nalpha", "Nbeta", "Ngamma", "Ndelta", "Gright",
        ] {
            assert!(out.contains(w), "{w} missing from {out}");
        }
    }
}

#[test]
fn test_doc_nested_table_cells_end_on_the_inner_cell_flag_not_on_a_bell() {
    // Nested cells end in `\r`; an empty one must still be a cell, or the row
    // has fewer cells than its row definition and the grid is lost.
    let paras = vec![
        cell("", 2),
        cell("only", 2),
        row_end(2, 2),
        cell("", 1),
        row_end(1, 1),
    ];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let inner = nested_table(&top_tables(&ir)[0].rows[0].cells[0]).expect("nested table");
    assert_eq!(inner.rows.len(), 1);
    assert_eq!(inner.rows[0].cells.len(), 2, "an empty cell is still a cell");
    assert_eq!(text_of(&inner.rows[0].cells[1]), "only");
}

#[test]
fn test_doc_nested_row_mark_without_a_row_definition_still_ends_the_row() {
    // `sprmPFInnerTtp` marks the row end by itself; a damaged file that lost
    // the row definition keeps its rows apart.
    let mut end = row_end(2, 1);
    end.grpprl = vec![0x16, 0x24, 0x01, 0x4C, 0x24, 0x01, 0x49, 0x66, 2, 0, 0, 0];
    let paras = vec![
        cell("a", 2),
        end,
        cell("b", 2),
        row_end(2, 1),
        cell("", 1),
        row_end(1, 1),
    ];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    let inner = nested_table(&top_tables(&ir)[0].rows[0].cells[0]).expect("nested table");
    assert_eq!(inner.rows.len(), 2, "two rows, not one merged row");
}

#[test]
fn test_doc_table_depth_beyond_a_byte_is_reported_not_wrapped_to_a_shallow_depth() {
    // `sprmPItap` is a 4-byte integer. 257 wrapped to depth 1 and silently
    // became an ordinary outer paragraph; it is out of range and is reported.
    let mut c = line("far too deep", 1);
    c.grpprl = vec![0x16, 0x24, 0x01, 0x49, 0x66, 0x01, 0x01, 0x00, 0x00]; // itap = 257
    let paras = vec![c, cell("", 1), row_end(1, 1)];
    let doc = open_doc(&build_doc(&paras));
    let ir = doc.to_ir();
    assert!(
        ir.metadata
            .warnings
            .iter()
            .any(|w| w.contains("nesting depth")),
        "{:?}",
        ir.metadata.warnings
    );
}

/// D | (E, then a 2-row nested table of distinctive words) | G
fn nested_table_with_distinctive_words() -> Vec<Para> {
    vec![
        cell("Dleft", 1),
        line("Emid", 1),
        cell("Nalpha", 2),
        cell("Nbeta", 2),
        row_end(2, 2),
        cell("Ngamma", 2),
        cell("Ndelta", 2),
        row_end(2, 2),
        cell("", 1),
        cell("Gright", 1),
        row_end(1, 3),
    ]
}

#[test]
fn test_doc_nested_table_text_survives_conversion_to_every_ooxml_format() {
    let doc = open_doc(&build_doc(&nested_table_with_distinctive_words()));
    let dir = std::env::temp_dir().join(format!("oo-nested-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    for ext in ["docx", "xlsx", "pptx"] {
        let out = dir.join(format!("out.{ext}"));
        doc.save_as(&out)
            .unwrap_or_else(|e| panic!("doc -> {ext}: {e}"));
        let text = office_oxide::Document::open(&out).unwrap().plain_text();
        for w in [
            "Dleft", "Emid", "Nalpha", "Nbeta", "Ngamma", "Ndelta", "Gright",
        ] {
            assert!(text.contains(w), "{w} lost converting to {ext}: {text:?}");
        }
    }
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_doc_nested_table_does_not_add_columns_to_the_markdown_table() {
    // A nested table's own `|` must not become cell separators of the outer
    // table: every row keeps the outer column count and there is one separator.
    let doc = open_doc(&build_doc(&nested_table_with_distinctive_words()));
    let md = doc.to_markdown();
    let cols = |line: &str| line.matches('|').count() - line.matches("\\|").count();
    let rows: Vec<&str> = md.lines().filter(|l| l.starts_with('|')).collect();
    assert!(!rows.is_empty(), "{md}");
    assert!(rows.iter().all(|l| cols(l) == cols(rows[0])), "ragged table:\n{md}");
    assert_eq!(rows.iter().filter(|l| l.contains("---")).count(), 1, "one separator row:\n{md}");
    for w in ["Nalpha", "Nbeta", "Ngamma", "Ndelta"] {
        assert!(md.contains(w), "{w} missing:\n{md}");
    }
}
