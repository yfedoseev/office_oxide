//! `.ppt` table reconstruction from a grid of grouped shapes.
//!
//! Legacy binary PowerPoint (97–2003) has no native "table" shape record
//! at all — [MS-ODRAW]'s `MSOSPT` preset-shape enumeration has no table
//! entry, and grepping the crate confirms no table-parsing code ever
//! existed here. A PowerPoint 2000+ table is authored as an
//! ordinary `OfficeArtSpgrContainer` *group* whose member shapes are laid
//! out in a rectangular grid — each cell is just a shape with its own
//! text, positioned via `OfficeArtChildAnchor` ([MS-ODRAW] §2.2.16 group
//! member anchor: `xLeft`/`yTop`/`xRight`/`yBottom`, all `i32`, group-local
//! coordinates). Without this module that text still reaches the IR (the
//! generic shape-recursion fallback walks any container) but as a flat,
//! disconnected sequence of paragraphs with no row/column structure.
//!
//! This is a *heuristic reconstruction*, not a spec-defined record parse
//! — unlike this session's other PPT work, there is no ground-truth byte
//! layout for "a group is a table," only "a group whose children happen
//! to form a rectangular grid of shapes." The detector is deliberately
//! conservative: it requires an exact, fully-populated rectangular grid
//! (no merged cells, no missing cells) before committing to
//! `Element::Table`. Anything less certain is left to the existing
//! flat-paragraph fallback — a safe "no regression" default, since a
//! failed detection never loses text, only the row/column structure it
//! already didn't have.

use super::text::TextRun;

/// One grid cell: its group-local bounding box plus whatever text runs
/// its shape's own text body produced.
#[derive(Debug, Clone)]
pub struct TableCellData {
    pub left: i32,
    pub top: i32,
    pub runs: Vec<TextRun>,
}

/// A reconstructed table: `rows[r][c]` is that cell's text runs.
#[derive(Debug, Clone, Default)]
pub struct TableBlock {
    /// `rows[r][c]` is that cell's text runs, in row-major order.
    pub rows: Vec<Vec<Vec<TextRun>>>,
}

/// Cluster a set of positions along one axis into row/column bands.
///
/// Real "Insert Table" grids align cells exactly, but this tolerates
/// small jitter: sorts the values, then starts a new cluster whenever the
/// gap to the previous value exceeds a tolerance derived from the overall
/// span (`span / 50`, floor 1) rather than a hardcoded absolute unit —
/// legacy `.ppt` anchors aren't in a fixed unit the crate has verified
/// (unlike EMU in OOXML), so an adaptive tolerance avoids baking in a
/// wrong guess. Returns one representative center per cluster, in
/// ascending order.
fn cluster_positions(mut values: Vec<i32>) -> Vec<i32> {
    if values.is_empty() {
        return Vec::new();
    }
    values.sort_unstable();
    let span = (values[values.len() - 1] - values[0]).max(0);
    let tolerance = (span / 50).max(1);

    let mut clusters: Vec<Vec<i32>> = vec![vec![values[0]]];
    for &v in &values[1..] {
        let last = clusters.last().unwrap();
        if v - last[last.len() - 1] <= tolerance {
            clusters.last_mut().unwrap().push(v);
        } else {
            clusters.push(vec![v]);
        }
    }
    clusters
        .into_iter()
        .map(|c| c.iter().sum::<i32>() / c.len() as i32)
        .collect()
}

/// Index of the cluster center nearest to `v`.
fn nearest_cluster(centers: &[i32], v: i32) -> usize {
    centers
        .iter()
        .enumerate()
        .min_by_key(|(_, c)| (**c - v).abs())
        .map(|(i, _)| i)
        .unwrap()
}

/// Build a `TableBlock` from a flat list of positioned cells, or `None`
/// if they don't form a clean, fully-populated rectangular grid.
///
/// Requires at least a 2x2 grid (a single row or column is just an
/// ordinary shape line-up, not worth promoting to `Element::Table`), and
/// requires `cells.len() == rows * cols` with no two cells landing in the
/// same (row, col) slot — any ambiguity (merged cells, an irregular
/// layout, overlapping shapes) bails out to the existing flat-paragraph
/// behavior rather than guessing.
pub fn build_table(cells: &[TableCellData]) -> Option<TableBlock> {
    if cells.len() < 4 {
        return None;
    }
    let row_centers = cluster_positions(cells.iter().map(|c| c.top).collect());
    let col_centers = cluster_positions(cells.iter().map(|c| c.left).collect());
    if row_centers.len() < 2 || col_centers.len() < 2 {
        return None;
    }
    if cells.len() != row_centers.len() * col_centers.len() {
        return None;
    }

    let mut grid: Vec<Vec<Option<Vec<TextRun>>>> =
        vec![vec![None; col_centers.len()]; row_centers.len()];
    for cell in cells {
        let r = nearest_cluster(&row_centers, cell.top);
        let c = nearest_cluster(&col_centers, cell.left);
        if grid[r][c].is_some() {
            return None; // two cells claim the same slot — not a clean grid
        }
        grid[r][c] = Some(cell.runs.clone());
    }
    if grid.iter().flatten().any(Option::is_none) {
        return None; // a slot with no cell at all — not fully populated
    }

    Some(TableBlock {
        rows: grid
            .into_iter()
            .map(|row| row.into_iter().map(Option::unwrap).collect())
            .collect(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(left: i32, top: i32) -> TableCellData {
        TableCellData {
            left,
            top,
            runs: Vec::new(),
        }
    }

    #[test]
    fn test_cluster_positions_groups_exact_duplicates() {
        let c = cluster_positions(vec![100, 100, 100, 500, 500, 500]);
        assert_eq!(c, vec![100, 500]);
    }

    #[test]
    fn test_cluster_positions_tolerates_small_jitter() {
        // Real files can be off by a handful of units per cell due to
        // rounding in whatever authored them; 3/4/5 must still cluster
        // with a far-away 500 group as two bands, not four.
        let c = cluster_positions(vec![100, 103, 98, 500, 502]);
        assert_eq!(c.len(), 2);
    }

    #[test]
    fn test_cluster_positions_empty_input() {
        assert!(cluster_positions(vec![]).is_empty());
    }

    #[test]
    fn test_build_table_recognizes_a_clean_2x2_grid() {
        let cells = vec![cell(0, 0), cell(100, 0), cell(0, 50), cell(100, 50)];
        let table = build_table(&cells).expect("should detect a 2x2 grid");
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].len(), 2);
    }

    #[test]
    fn test_build_table_recognizes_a_3x2_grid_regardless_of_input_order() {
        // 3 rows, 2 columns, shuffled input order.
        let cells = vec![
            cell(100, 100),
            cell(0, 0),
            cell(0, 100),
            cell(100, 0),
            cell(0, 200),
            cell(100, 200),
        ];
        let table = build_table(&cells).expect("should detect a 3x2 grid");
        assert_eq!(table.rows.len(), 3);
        assert_eq!(table.rows[0].len(), 2);
    }

    #[test]
    fn test_build_table_rejects_a_single_row() {
        // 1x3: a line-up of shapes, not a table.
        let cells = vec![cell(0, 0), cell(100, 0), cell(200, 0)];
        assert!(build_table(&cells).is_none());
    }

    #[test]
    fn test_build_table_rejects_a_single_column() {
        let cells = vec![cell(0, 0), cell(0, 100), cell(0, 200)];
        assert!(build_table(&cells).is_none());
    }

    #[test]
    fn test_build_table_rejects_a_grid_with_a_missing_cell() {
        // 2x2 grid minus its bottom-right cell: 3 shapes, not 4.
        let cells = vec![cell(0, 0), cell(100, 0), cell(0, 50)];
        assert!(build_table(&cells).is_none());
    }

    #[test]
    fn test_build_table_rejects_a_merged_cell_collision() {
        // A wide top cell spanning both columns collides with both
        // bottom cells' column slot once clustered — two cells can't
        // land in the same (row, col) slot, so this must bail rather
        // than silently pick one and drop the other's text.
        let cells = vec![
            cell(0, 0),   // "merged" top-left-ish
            cell(100, 0), // clusters into the same row as above; if its
            // left also clusters near 0 this is a collision case
            cell(0, 50),
            cell(100, 50),
            cell(50, 0), // an extra shape whose left is ambiguous between the two column clusters
        ];
        // 5 cells can never satisfy rows*cols exactly for any 2-cluster
        // split, so this must be rejected outright.
        assert!(build_table(&cells).is_none());
    }

    #[test]
    fn test_build_table_rejects_fewer_than_four_cells() {
        assert!(build_table(&[cell(0, 0), cell(1, 1), cell(2, 2)]).is_none());
    }

    #[test]
    fn test_build_table_preserves_cell_text_by_position() {
        let mut a = cell(0, 0);
        a.runs = vec![TextRun {
            text: "A".to_string(),
            ..Default::default()
        }];
        let mut b = cell(100, 0);
        b.runs = vec![TextRun {
            text: "B".to_string(),
            ..Default::default()
        }];
        let mut c = cell(0, 50);
        c.runs = vec![TextRun {
            text: "C".to_string(),
            ..Default::default()
        }];
        let mut d = cell(100, 50);
        d.runs = vec![TextRun {
            text: "D".to_string(),
            ..Default::default()
        }];
        let table = build_table(&[a, b, c, d]).unwrap();
        assert_eq!(table.rows[0][0][0].text, "A");
        assert_eq!(table.rows[0][1][0].text, "B");
        assert_eq!(table.rows[1][0][0].text, "C");
        assert_eq!(table.rows[1][1][0].text, "D");
    }
}
