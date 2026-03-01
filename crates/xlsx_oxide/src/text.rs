use crate::cell::{Cell, CellValue};
use crate::date;
use crate::worksheet::Row;
use crate::XlsxDocument;

impl XlsxDocument {
    /// Extract all text as a plain string (one sheet per section, tab-separated cells).
    pub fn plain_text(&self) -> String {
        let mut parts = Vec::new();
        for i in 0..self.worksheets.len() {
            if let Some(text) = self.sheet_plain_text(i) {
                if !text.is_empty() {
                    parts.push(text);
                }
            }
        }
        parts.join("\n\n")
    }

    /// Extract a single sheet as plain text.
    pub fn sheet_plain_text(&self, sheet_index: usize) -> Option<String> {
        let ws = self.worksheets.get(sheet_index)?;
        let mut lines = Vec::new();
        for row in &ws.rows {
            let cells: Vec<String> = row
                .cells
                .iter()
                .map(|c| self.format_cell_value(c))
                .collect();
            lines.push(cells.join("\t"));
        }
        Some(lines.join("\n"))
    }

    /// Convert to CSV string (default: first sheet).
    pub fn to_csv(&self) -> String {
        self.sheet_to_csv(0).unwrap_or_default()
    }

    /// Convert specific sheet to CSV (RFC 4180 compliant).
    pub fn sheet_to_csv(&self, sheet_index: usize) -> Option<String> {
        let ws = self.worksheets.get(sheet_index)?;
        let col_count = compute_column_count(&ws.rows);
        let mut lines = Vec::new();

        for row in &ws.rows {
            let mut fields: Vec<String> = Vec::with_capacity(col_count);
            for cell in &row.cells {
                fields.push(csv_escape(&self.format_cell_value(cell)));
            }
            // Pad to column count
            while fields.len() < col_count {
                fields.push(String::new());
            }
            lines.push(fields.join(","));
        }

        Some(lines.join("\r\n"))
    }

    /// Convert to markdown (pipe-delimited tables).
    pub fn to_markdown(&self) -> String {
        let mut parts = Vec::new();
        for i in 0..self.worksheets.len() {
            if let Some(md) = self.sheet_to_markdown(i) {
                if !md.is_empty() {
                    parts.push(md);
                }
            }
        }
        parts.join("\n\n")
    }

    /// Convert specific sheet to markdown.
    pub fn sheet_to_markdown(&self, sheet_index: usize) -> Option<String> {
        let ws = self.worksheets.get(sheet_index)?;
        if ws.rows.is_empty() {
            return Some(String::new());
        }

        let col_count = compute_column_count(&ws.rows);
        if col_count == 0 {
            return Some(String::new());
        }

        let mut lines = Vec::new();

        // Sheet name as heading
        lines.push(format!("## {}", ws.name));
        lines.push(String::new());

        // First row as header
        let header_row = &ws.rows[0];
        let header_cells: Vec<String> = (0..col_count)
            .map(|i| {
                header_row
                    .cells
                    .get(i)
                    .map(|c| self.format_cell_value(c))
                    .unwrap_or_default()
            })
            .collect();
        lines.push(format!("| {} |", header_cells.join(" | ")));

        // Separator row
        let sep: Vec<&str> = vec!["---"; col_count];
        lines.push(format!("| {} |", sep.join(" | ")));

        // Data rows
        for row in ws.rows.iter().skip(1) {
            let cells: Vec<String> = (0..col_count)
                .map(|i| {
                    row.cells
                        .get(i)
                        .map(|c| self.format_cell_value(c))
                        .unwrap_or_default()
                })
                .collect();
            lines.push(format!("| {} |", cells.join(" | ")));
        }

        Some(lines.join("\n"))
    }

    /// Format a cell value to a display string, applying date detection.
    pub fn format_cell_value(&self, cell: &Cell) -> String {
        match &cell.value {
            CellValue::Empty => String::new(),
            CellValue::Number(n) => {
                // Check if this should be a date
                if date::is_date_cell(cell.style_index, self.styles.as_ref()) {
                    if let Some(dt) = date::DateTimeValue::from_serial(
                        *n,
                        self.workbook.date1904,
                    ) {
                        return dt.to_iso_string();
                    }
                }
                format_number(*n)
            }
            CellValue::String(s) => s.clone(),
            CellValue::SharedString(idx) => self
                .shared_strings
                .get(*idx)
                .unwrap_or("")
                .to_string(),
            CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CellValue::Error(e) => e.clone(),
            CellValue::Date(dt) => dt.to_iso_string(),
        }
    }
}

/// Format a number, trimming unnecessary trailing zeros.
fn format_number(n: f64) -> String {
    if n == n.trunc() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

/// Compute the maximum number of columns across all rows.
fn compute_column_count(rows: &[Row]) -> usize {
    rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
}

/// Escape a field for CSV (RFC 4180).
fn csv_escape(field: &str) -> String {
    if field.contains(',') || field.contains('"') || field.contains('\n') || field.contains('\r') {
        let escaped = field.replace('"', "\"\"");
        format!("\"{escaped}\"")
    } else {
        field.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn csv_escape_plain() {
        assert_eq!(csv_escape("hello"), "hello");
    }

    #[test]
    fn csv_escape_with_comma() {
        assert_eq!(csv_escape("a,b"), "\"a,b\"");
    }

    #[test]
    fn csv_escape_with_quotes() {
        assert_eq!(csv_escape("say \"hi\""), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn csv_escape_with_newline() {
        assert_eq!(csv_escape("line1\nline2"), "\"line1\nline2\"");
    }

    #[test]
    fn format_number_integer() {
        assert_eq!(format_number(42.0), "42");
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(-10.0), "-10");
    }

    #[test]
    fn format_number_float() {
        assert_eq!(format_number(3.14), "3.14");
        assert_eq!(format_number(0.5), "0.5");
    }
}
