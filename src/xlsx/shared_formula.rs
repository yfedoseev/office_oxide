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

use std::collections::HashMap;

use super::cell::CellRef;

/// Excel's grid limits (ECMA-376 §18.3.1.73 / §18.3.1.13).
const MAX_COL: i64 = 16_383;
const MAX_ROW: i64 = 1_048_575;

/// Shared-formula masters collected while parsing one worksheet, keyed by
/// the group's `si` index.
#[derive(Debug, Default)]
pub struct SharedFormulas {
    masters: HashMap<u32, (CellRef, String)>,
    /// Followers whose master hadn't been seen yet. A group's master
    /// normally comes first, but not always — calamine ships a fixture with
    /// the records deliberately reversed — so these are resolved in a second
    /// pass once the whole sheet has been read.
    pending: Vec<(u32, CellRef)>,
}

impl SharedFormulas {
    /// Record a group's master cell and its formula text.
    pub fn add_master(&mut self, si: u32, at: CellRef, formula: String) {
        self.masters.insert(si, (at, formula));
    }

    /// Resolve a follower cell, or `None` if its master hasn't been seen
    /// yet — in which case the follower is remembered for `resolve_pending`.
    pub fn follower(&mut self, si: u32, at: &CellRef) -> Option<String> {
        match self.masters.get(&si) {
            Some((master_at, formula)) => Some(translate(formula, master_at, at)),
            None => {
                self.pending.push((si, at.clone()));
                None
            },
        }
    }

    /// Whether any follower is still waiting on its master.
    pub fn has_pending(&self) -> bool {
        !self.pending.is_empty()
    }

    /// Resolve every follower that was parsed before its master, as
    /// `(cell_ref, formula_text)` pairs. Followers whose `si` never had a
    /// master stay unresolved and are dropped.
    pub fn resolve_pending(&self) -> Vec<(CellRef, String)> {
        self.pending
            .iter()
            .filter_map(|(si, at)| {
                let (master_at, formula) = self.masters.get(si)?;
                Some((at.clone(), translate(formula, master_at, at)))
            })
            .collect()
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

    #[test]
    fn test_shared_formula_followers_before_master_resolve_in_second_pass() {
        let mut sf = SharedFormulas::default();
        // Follower first — the master record comes later in the file.
        assert_eq!(sf.follower(0, &r("B3")), None);
        assert!(sf.has_pending());
        sf.add_master(0, r("B2"), "A2".to_string());
        assert_eq!(sf.resolve_pending(), vec![(r("B3"), "A3".to_string())]);
        // Once the master is known, followers resolve inline.
        assert_eq!(sf.follower(0, &r("B4")).as_deref(), Some("A4"));
    }

    #[test]
    fn test_shared_formula_follower_without_any_master_is_dropped() {
        let mut sf = SharedFormulas::default();
        assert_eq!(sf.follower(7, &r("B3")), None);
        assert!(sf.resolve_pending().is_empty());
    }
}
