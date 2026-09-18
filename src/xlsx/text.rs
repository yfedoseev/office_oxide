use super::XlsxDocument;
use super::cell::{Cell, CellValue};
use super::date;
use super::numfmt;
use super::worksheet::Row;

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
        // `to_markdown()` already surfaces chart text (axis titles, series
        // names); `plain_text()` silently dropped it entirely (issue #331).
        for text in &self.chart_text {
            if !text.trim().is_empty() {
                parts.push(text.trim().to_string());
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
        // Charts: emit each chart's extracted text under a "## Chart N" heading
        // so its words appear in markdown / search / PDF without needing a
        // graphical chart renderer.
        for (i, text) in self.chart_text.iter().enumerate() {
            if !text.trim().is_empty() {
                parts.push(format!("## Chart {}\n\n{}", i + 1, text));
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

        // If the sheet is effectively single-column with prose-length cells
        // (notes, single-column reports), emit each cell as its own paragraph
        // instead of wrapping every line in a 1-column GFM table. The table
        // form looks awful when rendered (tall, narrow, hard to read) and
        // round-trips badly through markdown→IR→office.
        if col_count == 1
            && ws.rows.iter().any(|r| {
                r.cells
                    .first()
                    .map(|c| self.format_cell_value(c).chars().count() > 20)
                    .unwrap_or(false)
            })
        {
            let mut out = String::new();
            out.push_str(&format!("## {}\n\n", ws.name));
            for row in &ws.rows {
                if let Some(cell) = row.cells.first() {
                    let text = self.format_cell_value(cell);
                    if !text.trim().is_empty() {
                        out.push_str(text.trim());
                        out.push_str("\n\n");
                    }
                }
            }
            return Some(out.trim_end().to_string());
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
            // A formula cell with no cached `<v>` (the default output shape
            // of closedxml and similar writers) rendered as a blank cell
            // indistinguishable from a genuinely empty one, and the formula
            // text never reached any consumer at all (issue #279).
            CellValue::Empty => {
                if let Some(f) = &cell.formula {
                    buf.push('=');
                    buf.push_str(f);
                }
            },
            CellValue::Number(n) => {
                if date::is_date_cell(cell.style_index, self.styles.as_ref()) {
                    if let Some(dt) = date::DateTimeValue::from_serial(*n, self.workbook.date1904) {
                        buf.push_str(&dt.to_iso_string());
                        return;
                    }
                }
                if let Some(idx) = cell.style_index {
                    if let Some(styles) = self.styles.as_ref() {
                        if let Some(fmt_id) = styles.number_format_id_for(idx) {
                            if fmt_id != 0 {
                                // The explicit declaration only: apply_format's fmt_str branch is
                                // for custom codes, and feeding it a resolved
                                // built-in makes apply_custom mangle it
                                // (id 47 "mm:ss.0" rendered as "mm:ss0.6").
                                let fmt_str = styles.number_format_override_for(idx);
                                let formatted = numfmt::apply_format(*n, fmt_id, fmt_str);
                                buf.push_str(&formatted);
                                return;
                            }
                        }
                    }
                }
                write_number(*n, buf);
            },
            CellValue::String(s) => buf.push_str(s),
            CellValue::SharedString(idx) => {
                let s = self.shared_strings.get(*idx).unwrap_or("");
                // Truncate to prevent DoS from crafted shared strings
                if s.len() <= 32_768 {
                    buf.push_str(s);
                } else {
                    let mut end = 32_768;
                    while !s.is_char_boundary(end) && end > 0 {
                        end -= 1;
                    }
                    buf.push_str(&s[..end]);
                }
            },
            CellValue::Boolean(b) => buf.push_str(if *b { "TRUE" } else { "FALSE" }),
            CellValue::Error(e) => buf.push_str(e),
            CellValue::Date(dt) => buf.push_str(&dt.to_iso_string()),
        }
    }

    /// Pre-compute the set of style indices that map to date formats.
    /// Call once before iterating many cells; use with `write_cell_value_fast`.
    pub fn date_style_indices(&self) -> std::collections::HashSet<u32> {
        let Some(styles) = self.styles.as_ref() else {
            return Default::default();
        };
        (0..styles.cell_formats.len() as u32)
            .filter(|&idx| {
                let Some(fmt_id) = styles.number_format_id_for(idx) else {
                    return false;
                };
                // An explicit <numFmt> wins over the built-in meaning of its
                // id — [ECMA-376] §18.8.30 lets a workbook redefine ids
                // 0-163. Same precedence as `date::is_date_cell`; testing the
                // id first made `0.00000E+0` declared under id 50 render as a
                // 1900 date (#207, #225).
                if let Some(fmt_str) = styles.number_format_override_for(idx) {
                    return date::is_date_format_string(fmt_str);
                }
                date::is_date_format_id(fmt_id)
            })
            .collect()
    }

    /// Like `write_cell_value` but uses a pre-computed date style set instead
    /// of calling `is_date_cell()` (which re-scans format strings) per cell.
    pub fn write_cell_value_fast(
        &self,
        cell: &Cell,
        buf: &mut String,
        date_indices: &std::collections::HashSet<u32>,
    ) {
        match &cell.value {
            // See `write_cell_value`'s identical arm (issue #279).
            CellValue::Empty => {
                if let Some(f) = &cell.formula {
                    buf.push('=');
                    buf.push_str(f);
                }
            },
            CellValue::Number(n) => {
                let is_date = cell.style_index.is_some_and(|i| date_indices.contains(&i));
                if is_date {
                    if let Some(dt) = date::DateTimeValue::from_serial(*n, self.workbook.date1904) {
                        buf.push_str(&dt.to_iso_string());
                        return;
                    }
                }
                // Apply number format (thousands, decimals, %, currency, etc.)
                if let Some(idx) = cell.style_index {
                    if let Some(styles) = self.styles.as_ref() {
                        if let Some(fmt_id) = styles.number_format_id_for(idx) {
                            if fmt_id != 0 {
                                // The explicit declaration only: apply_format's fmt_str branch is
                                // for custom codes, and feeding it a resolved
                                // built-in makes apply_custom mangle it
                                // (id 47 "mm:ss.0" rendered as "mm:ss0.6").
                                let fmt_str = styles.number_format_override_for(idx);
                                let formatted = numfmt::apply_format(*n, fmt_id, fmt_str);
                                buf.push_str(&formatted);
                                return;
                            }
                        }
                    }
                }
                write_number(*n, buf);
            },
            CellValue::String(s) => buf.push_str(s),
            CellValue::SharedString(idx) => {
                let s = self.shared_strings.get(*idx).unwrap_or("");
                if s.len() <= 32_768 {
                    buf.push_str(s);
                } else {
                    let mut end = 32_768;
                    while !s.is_char_boundary(end) && end > 0 {
                        end -= 1;
                    }
                    buf.push_str(&s[..end]);
                }
            },
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

    /// #225 / #207 — `date_style_indices` backs `to_ir()`'s cell renderer
    /// and tested the built-in meaning of a `numFmtId` before the workbook's
    /// own `<numFmt>` override of that id, so `0.00000E+0` declared under id
    /// 50 was still treated as a date.
    #[test]
    fn test_date_style_indices_honours_numfmt_override_over_builtin_id() {
        let styles = br#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="50" formatCode="0.00000E+0"/>
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <cellXfs count="3">
    <xf numFmtId="50" applyNumberFormat="1"/>
    <xf numFmtId="164" applyNumberFormat="1"/>
    <xf numFmtId="14" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;
        let ss = super::super::styles::StyleSheet::parse(styles).expect("styles parse");
        let doc = XlsxDocument {
            workbook: super::super::WorkbookInfo {
                sheets: Vec::new(),
                defined_names: Vec::new(),
                date1904: false,
            },
            worksheets: Vec::new(),
            shared_strings: super::super::SharedStringTable::empty(),
            styles: Some(ss),
            theme: None,
            chart_text: Vec::new(),
            embedded_fonts: Vec::new(),
            core_properties: None,
            app_properties: None,
            has_macros: false,
            styles_data: None,
            theme_data: None,
        };
        let idx = doc.date_style_indices();
        assert!(!idx.contains(&0), "id 50 overridden to a numeric code is not a date");
        assert!(idx.contains(&1), "a custom yyyy-mm-dd code is a date");
        assert!(idx.contains(&2), "an un-overridden built-in date id is a date");
    }

    /// issue #331 — `to_markdown()` already surfaced chart text; `plain_text()`
    /// silently dropped it, so the CLI's default `text` output (and anything
    /// built on `plain_text()`, like PDF export) lost every chart's words.
    #[test]
    fn test_plain_text_includes_chart_text() {
        let doc = XlsxDocument {
            workbook: super::super::WorkbookInfo {
                sheets: Vec::new(),
                defined_names: Vec::new(),
                date1904: false,
            },
            worksheets: Vec::new(),
            shared_strings: super::super::SharedStringTable::empty(),
            styles: None,
            theme: None,
            chart_text: vec!["Title: Rotated Title".to_string()],
            embedded_fonts: Vec::new(),
            core_properties: None,
            app_properties: None,
            has_macros: false,
            styles_data: None,
            theme_data: None,
        };
        assert!(
            doc.plain_text().contains("Rotated Title"),
            "plain_text() must include chart text, same as to_markdown(): {:?}",
            doc.plain_text()
        );
    }
}
