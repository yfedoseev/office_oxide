//! Shared-formula group expansion.
//!
//! A `<f t="shared" ref="B2:B5" si="0">A2</f>` element carries the formula
//! text once, on the group's *master* cell. Every other cell in the group
//! holds only a bare `<f t="shared" si="0"/>` and the reader is expected to
//! reconstruct the formula by translating the master's relative references
//! by the row/column offset between the two cells — which is what Excel,
//! openpyxl, and every other serious reader do. office_oxide used to set
//! `Cell.formula = None` for every follower, making a formula cell
//! indistinguishable from a cell with no formula at all.

use std::collections::{HashMap, HashSet};

use super::cell::CellRef;
use super::worksheet::Row;
use crate::limits::{MAX_FORMULA_CHARS, TextBudget};

/// Excel's grid limits (ECMA-376 §18.3.1.73 / §18.3.1.13).
const MAX_COL: i64 = 16_383;
const MAX_ROW: i64 = 1_048_575;

/// A follower cell whose formula is still to be written: the group it names
/// and where its cell sits in the sheet's rows.
#[derive(Debug)]
struct Follower {
    si: u32,
    at: CellRef,
    row: usize,
    cell: usize,
}

/// What `SharedFormulas::resolve` could not give a formula to.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct SharedFormulaLoss {
    /// Masters longer than `MAX_FORMULA_CHARS`, and followers of such
    /// groups, plus followers past the document's text budget. Their cells
    /// keep their values and have no formula.
    pub dropped: usize,
}

/// Shared-formula masters and followers collected while parsing one
/// worksheet, keyed by the group's `si` index.
///
/// Parsing only records them. A follower stores its own copy of the master's
/// formula, so a long master shared by many cells multiplies; the copies are
/// made afterwards by `resolve`, one sheet at a time in the
/// workbook's order, against one document-wide budget. Sheets are parsed on
/// several threads, and charging as they went would make which sheet loses
/// its formulas depend on scheduling.
#[derive(Debug, Default)]
pub struct SharedFormulas {
    masters: HashMap<u32, (CellRef, String)>,
    /// Groups whose master was refused for its length.
    refused: HashSet<u32>,
    /// Masters refused, counted once each.
    refused_masters: usize,
    /// Every follower in document order. A group's master normally comes
    /// first, but not always — real files exist with the records
    /// deliberately reversed — so a follower is resolved only once the whole
    /// sheet has been read.
    followers: Vec<Follower>,
    /// Where the cell being parsed will sit: `(row, cell)` indexes into the
    /// sheet's rows.
    cursor: (usize, usize),
}

impl SharedFormulas {
    /// The row index last given to [`at_position`](Self::at_position).
    pub(crate) fn row(&self) -> usize {
        self.cursor.0
    }

    /// Note that the next cell parsed is cell `cell` of row `row`.
    pub(crate) fn at_position(&mut self, row: usize, cell: usize) {
        self.cursor = (row, cell);
    }

    /// Record a group's master cell and its formula text. A master longer
    /// than `MAX_FORMULA_CHARS` (Excel's own limit, ECMA-376 §18.17) is
    /// refused, as is its group: it cannot come from a spreadsheet
    /// application, and the text is copied into every follower. Returns
    /// `false` for a refused master, whose cell then has no formula.
    pub fn add_master(&mut self, si: u32, at: CellRef, formula: &str) -> bool {
        // A character is at most four bytes: only a short text is counted.
        let too_long =
            formula.len() > MAX_FORMULA_CHARS * 4 || formula.chars().count() > MAX_FORMULA_CHARS;
        if too_long {
            self.refused.insert(si);
            self.refused_masters += 1;
            return false;
        }
        self.masters.insert(si, (at, formula.to_string()));
        true
    }

    /// Record a follower cell (`<f t="shared" si="N"/>`) at the current
    /// position; its formula is written by `resolve`.
    pub fn follower(&mut self, si: u32, at: &CellRef) {
        self.followers.push(Follower {
            si,
            at: at.clone(),
            row: self.cursor.0,
            cell: self.cursor.1,
        });
    }

    /// Write every follower's formula into `rows`, charging each translated
    /// text to `budget` before it is kept. A follower whose group has no
    /// master stays without a formula, as before; one that does not fit the
    /// budget, or whose master was refused, is counted in the result and
    /// keeps its cell and value. Once the budget is spent no more is
    /// translated.
    pub(crate) fn resolve(self, rows: &mut [Row], budget: &mut TextBudget) -> SharedFormulaLoss {
        let mut dropped = self.refused_masters;
        for f in self.followers {
            if self.refused.contains(&f.si) {
                dropped += 1;
                continue;
            }
            let Some((master_at, formula)) = self.masters.get(&f.si) else {
                continue;
            };
            if budget.exhausted() {
                dropped += 1;
                continue;
            }
            let text = translate(formula, master_at, &f.at);
            if !budget.charge(text.len()) {
                dropped += 1;
                continue;
            }
            if let Some(cell) = rows.get_mut(f.row).and_then(|r| r.cells.get_mut(f.cell)) {
                cell.formula = Some(text);
            }
        }
        SharedFormulaLoss { dropped }
    }
}

/// Translate a formula written for cell `from` so it applies at cell `to`.
///
/// Only A1-style cell references are rewritten, and only the components a
/// `$` has not pinned. Everything else — function names, operators, string
/// literals, numbers, quoted sheet names — is copied through untouched. A
/// reference that would land off the grid becomes `#REF!`, as in Excel.
///
/// This is deliberately token-level rather than a formula parser: shifting
/// references is all a shared-formula group ever needs, and a full
/// expression AST would be a much larger surface to get wrong.
pub fn translate(formula: &str, from: &CellRef, to: &CellRef) -> String {
    let d_col = to.col as i64 - from.col as i64;
    let d_row = to.row as i64 - from.row as i64;
    if d_col == 0 && d_row == 0 {
        return formula.to_string();
    }

    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;

    while i < bytes.len() {
        let c = bytes[i];

        // Copy a string literal verbatim — a doubled quote escapes a quote.
        if c == b'"' {
            out.push('"');
            i += 1;
            while i < bytes.len() {
                out.push(bytes[i] as char);
                if bytes[i] == b'"' {
                    i += 1;
                    if i < bytes.len() && bytes[i] == b'"' {
                        out.push('"');
                        i += 1;
                        continue;
                    }
                    break;
                }
                i += 1;
            }
            continue;
        }

        // Copy a quoted sheet name verbatim ('My Sheet'!A1). Its contents
        // can look like anything, including a cell reference.
        if c == b'\'' {
            out.push('\'');
            i += 1;
            while i < bytes.len() {
                out.push(bytes[i] as char);
                let was_quote = bytes[i] == b'\'';
                i += 1;
                if was_quote {
                    break;
                }
            }
            continue;
        }

        if let Some((next, replacement)) = match_reference(formula, i, d_col, d_row) {
            out.push_str(&replacement);
            i = next;
            continue;
        }

        // Not the start of a reference: copy one byte. Formulas are ASCII
        // apart from string literals and names, both of which are copied
        // whole above, so this stays on a char boundary.
        let ch_len = utf8_len(c);
        out.push_str(&formula[i..i + ch_len]);
        i += ch_len;
    }

    out
}

fn utf8_len(first: u8) -> usize {
    match first {
        0x00..=0x7F => 1,
        0xC0..=0xDF => 2,
        0xE0..=0xEF => 3,
        _ => 4,
    }
}

/// Try to read an A1-style reference starting at `at`, returning the index
/// just past it and its translated text.
fn match_reference(s: &str, at: usize, d_col: i64, d_row: i64) -> Option<(usize, String)> {
    let b = s.as_bytes();

    // A reference can't continue an identifier or a number: `LOG10` and the
    // `E5` of `1.5E5` must not be mistaken for one.
    if at > 0 {
        let prev = b[at - 1];
        if prev.is_ascii_alphanumeric() || prev == b'_' || prev == b'.' || prev == b'$' {
            return None;
        }
    }

    let mut i = at;
    let col_abs = b.get(i) == Some(&b'$');
    if col_abs {
        i += 1;
    }

    let letters_start = i;
    while i < b.len() && b[i].is_ascii_alphabetic() {
        i += 1;
    }
    let letters = &s[letters_start..i];
    // XFD is the last column, so more than three letters is a name.
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }

    let row_abs = b.get(i) == Some(&b'$');
    if row_abs {
        i += 1;
    }

    let digits_start = i;
    while i < b.len() && b[i].is_ascii_digit() {
        i += 1;
    }
    if i == digits_start {
        return None;
    }

    // A name that merely looks like a reference: `LOG10(`, `SUM(`, a
    // defined name `Rate1` used as a value. Only the call form is
    // distinguishable, so that's the one guarded.
    if b.get(i) == Some(&b'(') {
        return None;
    }
    // `A1x` / `A1_` is part of a longer name, not a reference.
    if b.get(i)
        .is_some_and(|&n| n.is_ascii_alphanumeric() || n == b'_')
    {
        return None;
    }

    let col = CellRef::parse_col(letters)?;
    let row: u32 = s[digits_start..i].parse().ok()?;
    if row == 0 {
        return None;
    }

    let new_col = if col_abs {
        col as i64
    } else {
        col as i64 + d_col
    };
    // `row` is 1-based here, as written.
    let new_row = if row_abs {
        row as i64
    } else {
        row as i64 + d_row
    };

    if !(0..=MAX_COL).contains(&new_col) || !(1..=MAX_ROW + 1).contains(&new_row) {
        return Some((i, "#REF!".to_string()));
    }

    let mut out = String::with_capacity(i - at);
    if col_abs {
        out.push('$');
    }
    out.push_str(&CellRef::col_name(new_col as u32));
    if row_abs {
        out.push('$');
    }
    out.push_str(&new_row.to_string());
    Some((i, out))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn r(reference: &str) -> CellRef {
        CellRef::parse(reference).unwrap()
    }

    #[test]
    fn test_shared_formula_shifts_relative_references() {
        // B2's master `A2`, expanded down the group to B3/B4/B5.
        assert_eq!(translate("A2", &r("B2"), &r("B3")), "A3");
        assert_eq!(translate("A2", &r("B2"), &r("B5")), "A5");
        // And across columns.
        assert_eq!(translate("A2", &r("B2"), &r("C2")), "B2");
        assert_eq!(translate("SUM(A1:A10)", &r("B1"), &r("C3")), "SUM(B3:B12)");
    }

    #[test]
    fn test_shared_formula_respects_absolute_markers() {
        assert_eq!(translate("$A2", &r("B2"), &r("C3")), "$A3");
        assert_eq!(translate("A$2", &r("B2"), &r("C3")), "B$2");
        assert_eq!(translate("$A$2", &r("B2"), &r("C3")), "$A$2");
        assert_eq!(translate("A2+$B$2", &r("C2"), &r("C4")), "A4+$B$2");
    }

    #[test]
    fn test_shared_formula_leaves_non_references_alone() {
        // A function whose name ends in digits must not be read as a
        // reference, nor must an exponent, a string literal, or a quoted
        // sheet name's contents.
        assert_eq!(translate("LOG10(A1)", &r("B1"), &r("B2")), "LOG10(A2)");
        assert_eq!(translate("1.5E5*A1", &r("B1"), &r("B2")), "1.5E5*A2");
        assert_eq!(
            translate("CONCAT(\"A1 stays\",A1)", &r("B1"), &r("B2")),
            "CONCAT(\"A1 stays\",A2)"
        );
        assert_eq!(translate("'Sheet A1'!A1", &r("B1"), &r("B2")), "'Sheet A1'!A2");
        assert_eq!(translate("Sheet1!A1", &r("B1"), &r("B2")), "Sheet1!A2");
        // A zero delta is returned untouched.
        assert_eq!(translate("A1", &r("B1"), &r("B1")), "A1");
    }

    #[test]
    fn test_shared_formula_off_grid_reference_becomes_ref_error() {
        assert_eq!(translate("A1", &r("B2"), &r("B1")), "#REF!");
        assert_eq!(translate("A1", &r("B1"), &r("A1")), "#REF!");
    }

    fn rows_of(refs: &[&str]) -> Vec<Row> {
        refs.iter()
            .enumerate()
            .map(|(i, a)| Row {
                index: i as u32 + 1,
                cells: vec![crate::xlsx::Cell {
                    reference: r(a),
                    value: crate::xlsx::CellValue::Number(1.0),
                    style_index: None,
                    formula: None,
                    vm: None,
                    rich_runs: None,
                }],
            })
            .collect()
    }

    #[test]
    fn test_shared_formula_followers_before_master_resolve_in_second_pass() {
        let mut sf = SharedFormulas::default();
        let mut rows = rows_of(&["B3", "B2"]);
        // Follower first — the master record comes later in the file.
        sf.at_position(0, 0);
        sf.follower(0, &r("B3"));
        sf.at_position(1, 0);
        assert!(sf.add_master(0, r("B2"), "A2"));
        let loss = sf.resolve(&mut rows, &mut TextBudget::with_limit(100));
        assert_eq!(loss.dropped, 0);
        assert_eq!(rows[0].cells[0].formula.as_deref(), Some("A3"));
    }

    #[test]
    fn test_shared_formula_follower_without_any_master_is_dropped() {
        let mut sf = SharedFormulas::default();
        let mut rows = rows_of(&["B3"]);
        sf.at_position(0, 0);
        sf.follower(7, &r("B3"));
        let loss = sf.resolve(&mut rows, &mut TextBudget::with_limit(100));
        assert_eq!(loss.dropped, 0, "an orphan is not a loss of a formula that existed");
        assert_eq!(rows[0].cells[0].formula, None);
    }

    #[test]
    fn test_followers_past_the_budget_keep_their_cells_and_are_counted() {
        let mut sf = SharedFormulas::default();
        let mut rows = rows_of(&["B2", "B3", "B4"]);
        assert!(sf.add_master(0, r("B1"), "A1+A1"));
        for i in 0..3 {
            sf.at_position(i, 0);
            sf.follower(0, &rows[i].cells[0].reference.clone());
        }
        // "A2+A2" is 5 bytes: two fit in 12, the third does not.
        let loss = sf.resolve(&mut rows, &mut TextBudget::with_limit(12));
        assert_eq!(loss.dropped, 1);
        assert_eq!(rows[1].cells[0].formula.as_deref(), Some("A3+A3"));
        assert_eq!(rows[2].cells[0].formula, None);
    }

    #[test]
    fn test_a_master_over_excels_limit_drops_its_group_but_not_the_sheet() {
        let mut sf = SharedFormulas::default();
        let mut rows = rows_of(&["B2"]);
        assert!(!sf.add_master(0, r("B1"), &format!("A1+{}", "1+".repeat(MAX_FORMULA_CHARS))));
        sf.at_position(0, 0);
        sf.follower(0, &r("B2"));
        let loss = sf.resolve(&mut rows, &mut TextBudget::with_limit(1_000_000));
        assert_eq!(loss.dropped, 2);
        assert_eq!(rows[0].cells[0].formula, None);
    }
}
