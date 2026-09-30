//! Row-by-row lookup of rectangular cell ranges (merged cells, range
//! hyperlinks) without materialising every position a range covers.
//!
//! A single spec-legal range can cover the whole grid — `A1:XFD1048576` in
//! SpreadsheetML, rows 0..=65535 x cols 0..=255 in BIFF8 — so expanding each
//! range into a set of covered `(row, col)` positions costs work proportional
//! to the range's *area*, not to the document. The converters only ever ask
//! about positions that hold a materialised cell, and they visit rows in
//! order, so a sweep line over the ranges (sorted by first row) answers
//! "which ranges cover this row" and the caller clips each one to the
//! columns the row actually has.
//!
//! Overlapping ranges are not valid ([ECMA-376] §18.3.1.55 mergeCell,
//! [MS-XLS] §2.4.168 MergeCells), but nothing stops a file from declaring
//! millions of them over the same rows. Total work is therefore charged
//! against a fixed budget; once it is spent, [`RangeSweep::exhausted`]
//! reports it and no further ranges are returned, so the caller can state
//! the omission instead of hanging.

/// An inclusive, normalised (`lo <= hi`) rectangle of 0-based positions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellRange {
    pub row_lo: u32,
    pub row_hi: u32,
    pub col_lo: u32,
    pub col_hi: u32,
}

impl CellRange {
    /// Build a range from two corners given in any order.
    pub(crate) fn from_corners(r1: u32, c1: u32, r2: u32, c2: u32) -> Self {
        Self {
            row_lo: r1.min(r2),
            row_hi: r1.max(r2),
            col_lo: c1.min(c2),
            col_hi: c1.max(c2),
        }
    }

    /// Rows spanned, counted in u64 so a full-height range cannot overflow.
    pub(crate) fn row_span(&self) -> u64 {
        u64::from(self.row_hi - self.row_lo) + 1
    }

    /// Columns spanned.
    pub(crate) fn col_span(&self) -> u64 {
        u64::from(self.col_hi - self.col_lo) + 1
    }
}

/// Work units (range-row visits plus positions marked) one sheet may spend
/// before the sweep stops returning ranges. Ordinary sheets — thousands of
/// small merges over tens of thousands of rows — use a tiny fraction of it;
/// only a file declaring vast numbers of overlapping ranges gets near.
const WORK_BUDGET: u64 = 50_000_000;

/// Sweep over `ranges` in ascending row order.
pub(crate) struct RangeSweep<T> {
    ranges: Vec<(CellRange, T)>,
    /// Indices into `ranges`, sorted by `row_lo`.
    order: Vec<usize>,
    /// Next position in `order` not yet activated.
    next: usize,
    /// Indices into `ranges` whose rows include the current row, kept in
    /// ascending (declaration) order so "last declared wins" is a plain
    /// iteration.
    active: Vec<usize>,
    last_row: Option<u32>,
    work: u64,
    budget: u64,
    exhausted: bool,
}

impl<T> RangeSweep<T> {
    pub(crate) fn new(ranges: Vec<(CellRange, T)>) -> Self {
        Self::with_budget(ranges, WORK_BUDGET)
    }

    pub(crate) fn with_budget(ranges: Vec<(CellRange, T)>, budget: u64) -> Self {
        let mut order: Vec<usize> = (0..ranges.len()).collect();
        order.sort_by_key(|&i| ranges[i].0.row_lo);
        Self {
            ranges,
            order,
            next: 0,
            active: Vec::new(),
            last_row: None,
            work: 0,
            budget,
            exhausted: false,
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }

    /// Whether the work budget ran out; ranges were dropped from then on.
    pub(crate) fn exhausted(&self) -> bool {
        self.exhausted
    }

    /// Charge `units` of work; `false` once the budget is spent.
    pub(crate) fn charge(&mut self, units: u64) -> bool {
        if self.exhausted {
            return false;
        }
        self.work = self.work.saturating_add(units);
        if self.work > self.budget {
            self.exhausted = true;
            self.active.clear();
            return false;
        }
        true
    }

    /// The ranges covering `row`, in declaration order. Rows are expected in
    /// ascending order; a row that goes backwards restarts the sweep (its
    /// cost is charged like any other work).
    pub(crate) fn row(&mut self, row: u32) -> impl Iterator<Item = &(CellRange, T)> + '_ {
        self.advance(row);
        let ranges = &self.ranges;
        self.active.iter().map(move |&i| &ranges[i])
    }

    fn advance(&mut self, row: u32) {
        if self.exhausted {
            return;
        }
        if self.last_row.is_some_and(|last| row < last) {
            self.next = 0;
            self.active.clear();
            if !self.charge(self.ranges.len() as u64) {
                return;
            }
        }
        self.last_row = Some(row);
        let mut added = false;
        while let Some(&i) = self.order.get(self.next) {
            if self.ranges[i].0.row_lo > row {
                break;
            }
            self.next += 1;
            if self.ranges[i].0.row_hi >= row {
                self.active.push(i);
                added = true;
            }
        }
        if added {
            self.active.sort_unstable();
        }
        let ranges = &self.ranges;
        self.active.retain(|&i| ranges[i].0.row_hi >= row);
        let cost = self.active.len() as u64 + 1;
        self.charge(cost);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn r(row_lo: u32, row_hi: u32, col_lo: u32, col_hi: u32) -> CellRange {
        CellRange {
            row_lo,
            row_hi,
            col_lo,
            col_hi,
        }
    }

    #[test]
    fn test_sweep_returns_ranges_covering_each_row_in_declaration_order() {
        let mut s = RangeSweep::new(vec![(r(2, 4, 0, 1), 'b'), (r(0, 3, 5, 5), 'a')]);
        let at = |s: &mut RangeSweep<char>, row| s.row(row).map(|(_, t)| *t).collect::<String>();
        assert_eq!(at(&mut s, 0), "a");
        assert_eq!(at(&mut s, 2), "ba");
        assert_eq!(at(&mut s, 4), "b");
        assert_eq!(at(&mut s, 5), "");
        // Going backwards restarts rather than answering wrongly.
        assert_eq!(at(&mut s, 3), "ba");
    }

    #[test]
    fn test_full_grid_range_costs_one_unit_per_visited_row() {
        let mut s = RangeSweep::new(vec![(r(0, 1_048_575, 0, 16_383), ())]);
        for row in [0, 5, 1_048_575] {
            assert_eq!(s.row(row).count(), 1);
        }
        assert!(!s.exhausted());
    }

    #[test]
    fn test_budget_exhaustion_is_reported_and_stops_returning_ranges() {
        let ranges = (0..100).map(|_| (r(0, 10, 0, 0), ())).collect();
        let mut s = RangeSweep::with_budget(ranges, 250);
        assert_eq!(s.row(0).count(), 100);
        assert_eq!(s.row(1).count(), 100);
        assert_eq!(s.row(2).count(), 0);
        assert!(s.exhausted());
    }

    #[test]
    fn test_spans_of_a_full_grid_range_do_not_overflow() {
        let full = CellRange::from_corners(u32::MAX, u32::MAX, 0, 0);
        assert_eq!(full.row_span(), 1 << 32);
        assert_eq!(full.col_span(), 1 << 32);
    }
}
