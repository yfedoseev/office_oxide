use super::cell::{Cell, CellValue};
use super::date;
use super::worksheet::Row;
use super::XlsxDocument;

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
        let mut buf = String::with_capacity(ws.rows.len() * 64);
        for (row_idx, row) in ws.rows.iter().enumerate() {
            if row_idx > 0 {
                buf.push('\n');
            }
            for (col_idx, cell) in row.cells.iter().enumerate() {
                if col_idx > 0 {
                    buf.push('\t');
                }
                self.write_cell_value(cell, &mut buf);
            }
        }
        Some(buf)
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
        let mut buf = String::new();
        self.write_cell_value(cell, &mut buf);
        buf
    }

    /// Write a cell value directly to a buffer (avoids allocation for shared strings).
    pub fn write_cell_value(&self, cell: &Cell, buf: &mut String) {
        match &cell.value {
            CellValue::Empty => {}
            CellValue::Number(n) => {
                if date::is_date_cell(cell.style_index, self.styles.as_ref()) {
                    if let Some(dt) = date::DateTimeValue::from_serial(
                        *n,
                        self.workbook.date1904,
                    ) {
                        buf.push_str(&dt.to_iso_string());
                        return;
                    }
                }
                write_number(*n, buf);
            }
            CellValue::String(s) => buf.push_str(s),
            CellValue::SharedString(idx) => {
                let s = self.shared_strings.get(*idx).unwrap_or("");
                // Truncate to prevent DoS from crafted shared strings
                if s.len() <= 32_768 {
                    buf.push_str(s);
                } else {
                    let mut end = 32_768;
                    while !s.is_char_boundary(end) && end > 0 { end -= 1; }
                    buf.push_str(&s[..end]);
                }
            }
            CellValue::Boolean(b) => buf.push_str(if *b { "TRUE" } else { "FALSE" }),
            CellValue::Error(e) => buf.push_str(e),
            CellValue::Date(dt) => buf.push_str(&dt.to_iso_string()),
        }
    }
}

/// Write a formatted number directly to a buffer.
fn write_number(n: f64, buf: &mut String) {
    use std::fmt::Write;
    if n == n.trunc() && n.abs() < 1e15 {
        write!(buf, "{}", n as i64).ok();
    } else {
        write!(buf, "{}", n).ok();
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

    fn fmt_num(n: f64) -> String {
        let mut buf = String::new();
        write_number(n, &mut buf);
        buf
    }

    #[test]
    fn format_number_integer() {
        assert_eq!(fmt_num(42.0), "42");
        assert_eq!(fmt_num(0.0), "0");
        assert_eq!(fmt_num(-10.0), "-10");
    }

    #[test]
    fn format_number_float() {
        assert_eq!(fmt_num(3.15), "3.15");
        assert_eq!(fmt_num(0.5), "0.5");
    }
}
