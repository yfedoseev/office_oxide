use super::XlsxDocument;
use super::cell::{Cell, CellValue};
use super::date;
use super::numfmt;
use super::worksheet::Row;
use crate::limits::TextBudget;

/// `B2 (Author)` — the same marker `to_ir()` puts on a comment's endnote.
pub(crate) fn comment_marker(cell_ref: &str, author: Option<&str>) -> String {
    match author {
        Some(a) => format!("{cell_ref} ({a})"),
        None => cell_ref.to_string(),
    }
}

/// Per-render cell formatting state: the date styles found once
/// (`date_style_indices`) and each custom `<numFmt>` compiled once, rather
/// than a date re-check of the format code and a re-parse of it for every
/// cell. Every direct renderer and `to_ir()` format through this.
pub(crate) struct CellFormatter<'a> {
    doc: &'a XlsxDocument,
    date_indices: std::collections::HashSet<u32>,
    compiled: std::collections::HashMap<u32, Option<numfmt::CompiledFormat>>,
}

impl<'a> CellFormatter<'a> {
    pub(crate) fn new(doc: &'a XlsxDocument) -> Self {
        Self {
            doc,
            date_indices: doc.date_style_indices(),
            compiled: std::collections::HashMap::new(),
        }
    }

    pub(crate) fn date_indices(&self) -> &std::collections::HashSet<u32> {
        &self.date_indices
    }

    /// Append `cell`'s display text to `buf`.
    pub(crate) fn write(&mut self, cell: &Cell, buf: &mut String) {
        let doc = self.doc;
        let date_indices = &self.date_indices;
        let compiled = &mut self.compiled;
        doc.write_value_with(
            cell,
            buf,
            |idx| date_indices.contains(&idx),
            |n, idx, fmt_id| {
                let styles = doc.styles.as_ref();
                let code = compiled.entry(fmt_id).or_insert_with(|| {
                    styles
                        .and_then(|s| s.number_format_override_for(idx))
                        .and_then(numfmt::compile_custom)
                });
                numfmt::apply_format_compiled(n, fmt_id, code.as_ref())
            },
        );
    }

    /// `cell`'s display text.
    pub(crate) fn text(&mut self, cell: &Cell) -> String {
        let mut buf = String::new();
        self.write(cell, &mut buf);
        buf
    }
}

impl XlsxDocument {
    /// Extract all text as a plain string: each sheet's name on its own
    /// line, then its rows with tab-separated cells.
    ///
    /// The sheet name is emitted for every sheet, as the `.xls` reader,
    /// `to_markdown()` and every other spreadsheet reader (POI, openpyxl,
    /// xlrd, calamine) do — the same workbook used to name its sheets in
    /// one format and not the other.
    pub fn plain_text(&self) -> String {
        let mut parts = Vec::new();
        // One text budget for the whole document: a shared string
        // referenced from every cell is rendered once per cell, and
        // nothing else bounds that product (see `crate::limits`).
        let mut budget = TextBudget::new();
        for (i, ws) in self.worksheets.iter().enumerate() {
            let mut sheet = ws.name.clone();
            // Page headers above the cells and footers below, as the IR
            // renderers place a section's headers and footers.
            let [fh, oh, eh, ff, of, ef] = ws.header_footer.active(&ws.name);
            for h in [fh, oh, eh].into_iter().flatten() {
                sheet.push('\n');
                sheet.push_str(&h);
            }
            if let Some(text) = self.sheet_plain_text_within(i, &mut budget) {
                if !text.is_empty() {
                    sheet.push('\n');
                    sheet.push_str(&text);
                }
            }
            for f in [ff, of, ef].into_iter().flatten() {
                sheet.push('\n');
                sheet.push_str(&f);
            }
            if budget.exhausted() {
                sheet.push('\n');
                sheet.push_str(&budget.notice());
                parts.push(sheet);
                break;
            }
            // Cell comments are document content; `to_ir()` carries them
            // as endnotes, and this direct renderer dropped them.
            for c in &ws.comments {
                sheet.push('\n');
                sheet.push_str(&format!(
                    "{}: {}",
                    comment_marker(&c.cell_ref, c.author.as_deref()),
                    c.text
                ));
            }
            // Text boxes and WordArt drawn on the sheet reached `to_ir()`
            // as text boxes and this renderer not at all — a sheet whose
            // content is a drawn note came back as its name alone.
            for ts in &ws.text_shapes {
                if !ts.text.trim().is_empty() {
                    sheet.push('\n');
                    sheet.push_str(ts.text.trim());
                }
            }
            parts.push(sheet);
        }
        // `to_markdown()` already surfaces chart text (axis titles, series
        // names); `plain_text()` silently dropped it entirely.
        for text in &self.chart_text {
            if !text.trim().is_empty() {
                parts.push(text.trim().to_string());
            }
        }
        for (name, err) in &self.unreadable_sheets {
            parts.push(unreadable_notice(name, err));
        }
        parts.join("\n\n")
    }

    /// Extract a single sheet as plain text.
    pub fn sheet_plain_text(&self, sheet_index: usize) -> Option<String> {
        let mut budget = TextBudget::new();
        let mut text = self.sheet_plain_text_within(sheet_index, &mut budget)?;
        if budget.exhausted() {
            text.push('\n');
            text.push_str(&budget.notice());
        }
        Some(text)
    }

    fn sheet_plain_text_within(
        &self,
        sheet_index: usize,
        budget: &mut TextBudget,
    ) -> Option<String> {
        let ws = self.worksheets.get(sheet_index)?;
        let mut fmt = CellFormatter::new(self);
        let mut buf = String::with_capacity(ws.rows.len() * 64);
        for (row_idx, row) in ws.rows.iter().enumerate() {
            if row_idx > 0 {
                buf.push('\n');
            }
            // A cell sits in its own column: XLSX stores only non-empty
            // cells (§18.3.1.4), so the columns a row skips are tabs, not
            // nothing. A cell out of order or repeating a column follows the
            // previous one rather than being dropped.
            let mut next_col = 0u32;
            for (i, cell) in row.cells.iter().enumerate() {
                let col = cell.reference.col;
                let gap = col.saturating_sub(next_col) as usize;
                let tabs = if i == 0 { gap } else { gap + 1 };
                let before = buf.len();
                buf.extend(std::iter::repeat_n('\t', tabs));
                fmt.write(cell, &mut buf);
                if !budget.charge(buf.len() - before) {
                    buf.truncate(before);
                    return Some(buf);
                }
                next_col = next_col.max(col.saturating_add(1));
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
        let mut budget = TextBudget::new();
        let mut fmt = CellFormatter::new(self);

        // Rows the sheet does not store are empty lines, so every value
        // keeps its row number (§18.3.1.73 `r`); the padding is charged to
        // the budget like text, which bounds a sheet whose two cells sit
        // at A1 and XFD1048576.
        let empty_line = ",".repeat(col_count.saturating_sub(1));
        let mut next_row = 1u32;
        'rows: for row in &ws.rows {
            while next_row < row.index {
                if !budget.charge(empty_line.len() + 2) {
                    lines.push(budget.notice());
                    break 'rows;
                }
                lines.push(empty_line.clone());
                next_row += 1;
            }
            next_row = next_row.max(row.index.saturating_add(1));
            let mut fields: Vec<String> = vec![String::new(); col_count];
            if !budget.charge(col_count) {
                lines.push(budget.notice());
                break 'rows;
            }
            for cell in &row.cells {
                let field = csv_escape(&fmt.text(cell));
                if !budget.charge(field.len()) {
                    lines.push(budget.notice());
                    break 'rows;
                }
                place(&mut fields, cell.reference.col, field);
            }
            lines.push(fields.join(","));
        }

        Some(lines.join("\r\n"))
    }

    /// Convert to markdown (pipe-delimited tables).
    pub fn to_markdown(&self) -> String {
        let mut parts = Vec::new();
        let mut budget = TextBudget::new();
        for (i, ws) in self.worksheets.iter().enumerate() {
            let [fh, oh, eh, ff, of, ef] = ws.header_footer.active(&ws.name);
            if let Some(md) = self.sheet_to_markdown_within(i, &mut budget) {
                if !md.is_empty() {
                    parts.push(md);
                } else if !ws.comments.is_empty()
                    || ws.text_shapes.iter().any(|t| !t.text.trim().is_empty())
                    || !ws.header_footer.is_empty()
                {
                    // No cells, but comments, drawn text or page
                    // headers: they still belong under the sheet's heading.
                    parts.push(format!("## {}", ws.name));
                }
            }
            for hf in [fh, oh, eh, ff, of, ef].into_iter().flatten() {
                parts.push(crate::core::markdown::escape_text(&hf));
            }
            if budget.exhausted() {
                parts.push(budget.notice());
                break;
            }
            for c in &ws.comments {
                parts.push(format!(
                    "> **{}:** {}",
                    comment_marker(&c.cell_ref, c.author.as_deref()),
                    c.text.trim()
                ));
            }
            for ts in &ws.text_shapes {
                if !ts.text.trim().is_empty() {
                    parts.push(ts.text.trim().to_string());
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
        for (name, err) in &self.unreadable_sheets {
            parts.push(format!("## {name}\n\n{}", unreadable_notice(name, err)));
        }
        parts.join("\n\n")
    }

    /// Convert specific sheet to markdown.
    pub fn sheet_to_markdown(&self, sheet_index: usize) -> Option<String> {
        let mut budget = TextBudget::new();
        let mut md = self.sheet_to_markdown_within(sheet_index, &mut budget)?;
        if budget.exhausted() {
            md.push_str("\n\n");
            md.push_str(&budget.notice());
        }
        Some(md)
    }

    fn sheet_to_markdown_within(
        &self,
        sheet_index: usize,
        budget: &mut TextBudget,
    ) -> Option<String> {
        let ws = self.worksheets.get(sheet_index)?;
        if ws.rows.is_empty() {
            return Some(String::new());
        }
        // Every cell text passes through here; a spent budget ends the
        // sheet at the cell that spent it.
        let mut fmt = CellFormatter::new(self);

        let col_count = compute_column_count(&ws.rows);
        if col_count == 0 {
            return Some(String::new());
        }

        // If the sheet is effectively single-column with prose-length cells
        // (notes, single-column reports), emit each cell as its own paragraph
        // instead of wrapping every line in a 1-column GFM table. The table
        // form looks awful when rendered (tall, narrow, hard to read) and
        // round-trips badly through markdown→IR→office.
        let prose = col_count == 1
            && ws.rows.iter().any(|r| {
                r.cells
                    .first()
                    .is_some_and(|c| fmt.text(c).chars().count() > 20)
            });
        let mut cell_text = |cell: &Cell| -> Option<String> {
            let text = fmt.text(cell);
            budget
                .charge(text.len())
                .then(|| crate::core::markdown::escape_cell(&text))
        };
        if prose {
            let mut out = String::new();
            out.push_str(&format!("## {}\n\n", ws.name));
            for row in &ws.rows {
                if let Some(cell) = row.cells.first() {
                    let Some(text) = cell_text(cell) else { break };
                    if !text.trim().is_empty() {
                        // `cell_text` escaped for a table cell; prose keeps
                        // its line breaks as paragraphs.
                        out.push_str(&text.replace("<br>", "\n"));
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
        // A row keeps the cells that fit the budget (the rest empty) and
        // reports that the budget is spent, so the table ends after it.
        // Each cell goes under its own column's header (§18.3.1.4 `r`).
        let mut row_line = |row: &Row| -> (String, bool) {
            let mut cells: Vec<String> = vec![String::new(); col_count];
            let mut spent = false;
            for c in &row.cells {
                match cell_text(c) {
                    Some(t) => place(&mut cells, c.reference.col, t),
                    None => {
                        spent = true;
                        break;
                    },
                }
            }
            (format!("| {} |", cells.join(" | ")), spent)
        };
        let (header, spent) = row_line(&ws.rows[0]);
        lines.push(header);

        // Separator row
        let sep: Vec<&str> = vec!["---"; col_count];
        lines.push(format!("| {} |", sep.join(" | ")));

        // Data rows
        if !spent {
            for row in ws.rows.iter().skip(1) {
                let (line, spent) = row_line(row);
                lines.push(line);
                if spent {
                    break;
                }
            }
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
    ///
    /// Re-checks the cell's format for date tokens and re-parses a custom
    /// format code on every call; a caller rendering many cells should go
    /// through a per-render formatter instead (as every renderer here does).
    pub fn write_cell_value(&self, cell: &Cell, buf: &mut String) {
        let styles = self.styles.as_ref();
        self.write_value_with(
            cell,
            buf,
            |idx| date::is_date_cell(Some(idx), styles),
            |n, idx, fmt_id| {
                // The explicit declaration only: apply_format's fmt_str branch
                // is for custom codes, and feeding it a resolved built-in makes
                // the custom-format engine mangle it (id 47 "mm:ss.0" rendered as
                // "mm:ss0.6").
                let fmt_str = styles.and_then(|s| s.number_format_override_for(idx));
                numfmt::apply_format(n, fmt_id, fmt_str)
            },
        );
    }

    /// The one cell-to-text rendering every entry point shares. `is_date`
    /// answers for a style index; `format_number` renders `(value, style
    /// index, non-General format id)`.
    fn write_value_with(
        &self,
        cell: &Cell,
        buf: &mut String,
        is_date: impl FnOnce(u32) -> bool,
        format_number: impl FnOnce(f64, u32, u32) -> String,
    ) {
        match &cell.value {
            // A formula cell with no cached `<v>` (the default output shape
            // of closedxml and similar writers) rendered as a blank cell
            // indistinguishable from a genuinely empty one, and the formula
            // text never reached any consumer at all.
            CellValue::Empty => {
                if let Some(f) = &cell.formula {
                    buf.push('=');
                    buf.push_str(f);
                }
            },
            CellValue::Number(n) => {
                if cell.style_index.is_some_and(is_date) {
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
                                buf.push_str(&format_number(*n, idx, fmt_id));
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
                // 1900 date.
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
        let styles = self.styles.as_ref();
        self.write_value_with(
            cell,
            buf,
            |idx| date_indices.contains(&idx),
            |n, idx, fmt_id| {
                let fmt_str = styles.and_then(|s| s.number_format_override_for(idx));
                numfmt::apply_format(n, fmt_id, fmt_str)
            },
        );
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

/// The sheet's width: one past the rightmost column any cell references
/// (not the longest row's cell count — a sparse row is short but wide).
fn compute_column_count(rows: &[Row]) -> usize {
    rows.iter()
        .flat_map(|r| &r.cells)
        .map(|c| c.reference.col as usize + 1)
        .max()
        .unwrap_or(0)
}

/// Put `value` in column `col` of a row laid out `fields.len()` wide. A
/// column already filled (two cells claiming one reference) keeps its first
/// value, as the IR path does.
fn place(fields: &mut [String], col: u32, value: String) {
    if let Some(slot) = fields.get_mut(col as usize) {
        if slot.is_empty() {
            *slot = value;
        }
    }
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

/// The line every renderer emits for a sheet that could not be read, so
/// a workbook missing a sheet never passes for a complete one.
pub(crate) fn unreadable_notice(name: &str, err: &str) -> String {
    format!("[unreadable sheet {name:?}: {err}]")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_csv_escape_plain() {
        assert_eq!(csv_escape("hello"), "hello");
    }

    #[test]
    fn test_csv_escape_with_comma() {
        assert_eq!(csv_escape("a,b"), "\"a,b\"");
    }

    #[test]
    fn test_csv_escape_with_quotes() {
        assert_eq!(csv_escape("say \"hi\""), "\"say \"\"hi\"\"\"");
    }

    #[test]
    fn test_csv_escape_with_newline() {
        assert_eq!(csv_escape("line1\nline2"), "\"line1\nline2\"");
    }

    fn fmt_num(n: f64) -> String {
        let mut buf = String::new();
        write_number(n, &mut buf);
        buf
    }

    #[test]
    fn test_format_number_integer() {
        assert_eq!(fmt_num(42.0), "42");
        assert_eq!(fmt_num(0.0), "0");
        assert_eq!(fmt_num(-10.0), "-10");
    }

    #[test]
    fn test_format_number_float() {
        assert_eq!(fmt_num(3.15), "3.15");
        assert_eq!(fmt_num(0.5), "0.5");
    }

    /// `date_style_indices` backs `to_ir()`'s cell renderer
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
            package_properties: Default::default(),
            has_macros: false,
            unreadable_sheets: Vec::new(),
            styles_data: None,
            theme_data: None,
        };
        let idx = doc.date_style_indices();
        assert!(!idx.contains(&0), "id 50 overridden to a numeric code is not a date");
        assert!(idx.contains(&1), "a custom yyyy-mm-dd code is a date");
        assert!(idx.contains(&2), "an un-overridden built-in date id is a date");
    }

    /// A text box or WordArt drawn on a sheet reached `to_ir()` as a text
    /// box and the direct renderers not at all — a sheet whose only
    /// content was a drawn note came back as its name alone, and the
    /// markdown had no heading for it.
    #[test]
    fn test_drawn_text_shapes_reach_plain_text_and_markdown() {
        let ws = super::super::worksheet::Worksheet {
            name: "Notes".to_string(),
            dimension: None,
            rows: Vec::new(),
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            images: Vec::new(),
            comments: Vec::new(),
            text_shapes: vec![super::super::worksheet::WorksheetTextShape {
                text: "Lorem ipsum drawn in a text box".to_string(),
                font_name: None,
                font_size_pt: None,
                bold: false,
                italic: false,
                color_hex: None,
                x_emu: 0,
                y_emu: 0,
                cx_emu: 100,
                cy_emu: 100,
            }],
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            state: super::super::SheetState::Visible,
            header_footer: Default::default(),
            tables: Vec::new(),
            pivot_tables: Vec::new(),
        };
        let doc = XlsxDocument {
            workbook: super::super::WorkbookInfo {
                sheets: Vec::new(),
                defined_names: Vec::new(),
                date1904: false,
            },
            worksheets: vec![ws],
            shared_strings: super::super::SharedStringTable::empty(),
            styles: None,
            theme: None,
            chart_text: Vec::new(),
            embedded_fonts: Vec::new(),
            core_properties: None,
            app_properties: None,
            package_properties: Default::default(),
            has_macros: false,
            unreadable_sheets: Vec::new(),
            styles_data: None,
            theme_data: None,
        };
        let text = doc.plain_text();
        assert!(text.starts_with("Notes\nLorem ipsum drawn"), "{text:?}");
        let md = doc.to_markdown();
        assert!(md.starts_with("## Notes\n\nLorem ipsum drawn"), "{md:?}");
        assert!(
            crate::convert_xlsx::xlsx_to_ir(&doc)
                .plain_text()
                .contains("Lorem ipsum drawn")
        );
    }

    /// `plain_text()` starts each sheet with its name, as `.xls`,
    /// `to_markdown()` and every other spreadsheet reader do — the same
    /// workbook used to name its sheets in one format and not the other.
    #[test]
    fn test_plain_text_names_each_sheet() {
        let sheet = |name: &str, text: &str| super::super::worksheet::Worksheet {
            name: name.to_string(),
            dimension: None,
            rows: vec![Row {
                index: 1,
                cells: vec![Cell {
                    reference: super::super::CellRef { col: 0, row: 0 },
                    value: CellValue::String(text.to_string()),
                    style_index: None,
                    formula: None,
                    rich_runs: None,
                    vm: None,
                }],
            }],
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            page_setup: None,
            images: Vec::new(),
            comments: Vec::new(),
            text_shapes: Vec::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            state: super::super::SheetState::Visible,
            header_footer: Default::default(),
            tables: Vec::new(),
            pivot_tables: Vec::new(),
        };
        let doc = XlsxDocument {
            workbook: super::super::WorkbookInfo {
                sheets: Vec::new(),
                defined_names: Vec::new(),
                date1904: false,
            },
            worksheets: vec![sheet("Revenue", "north"), sheet("Costs", "south")],
            shared_strings: super::super::SharedStringTable::empty(),
            styles: None,
            theme: None,
            chart_text: Vec::new(),
            embedded_fonts: Vec::new(),
            core_properties: None,
            app_properties: None,
            package_properties: Default::default(),
            has_macros: false,
            unreadable_sheets: Vec::new(),
            styles_data: None,
            theme_data: None,
        };
        let text = doc.plain_text();
        assert!(text.starts_with("Revenue\nnorth"), "{text:?}");
        assert!(text.contains("\n\nCosts\nsouth"), "{text:?}");
        // The two renderers agree on the names.
        let md = doc.to_markdown();
        assert!(md.contains("## Revenue") && md.contains("## Costs"), "{md:?}");
    }

    /// `to_markdown()` already surfaced chart text; `plain_text()`
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
            package_properties: Default::default(),
            has_macros: false,
            unreadable_sheets: Vec::new(),
            styles_data: None,
            theme_data: None,
        };
        assert!(
            doc.plain_text().contains("Rotated Title"),
            "plain_text() must include chart text, same as to_markdown(): {:?}",
            doc.plain_text()
        );
    }

    /// `plain_text()`/`to_csv()`/`to_markdown()` checked every numeric
    /// cell for a date format by re-scanning its format code; the per-style
    /// answer is now computed once per render, as `to_ir()` already did.
    #[test]
    fn test_direct_renderers_scan_each_date_format_once_not_per_cell() {
        use crate::xlsx::test_support::{open_bytes, single_sheet_xlsx};
        let styles =
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs>
</styleSheet>"#;
        let cells: String = (1..=100)
            .map(|r| format!(r#"<row r="{r}"><c r="A{r}" s="1"><v>{}</v></c></row>"#, 38000 + r))
            .collect();
        let sheet = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{cells}</sheetData></worksheet>"#
        );
        let doc = open_bytes(single_sheet_xlsx(&sheet, &[("xl/styles.xml", styles)]));
        let scans = |f: &dyn Fn() -> String| {
            date::DATE_FORMAT_SCANS.with(|n| n.set(0));
            let out = f();
            assert!(out.contains("2004-01-15"), "{out}");
            date::DATE_FORMAT_SCANS.with(|n| n.get())
        };
        for (name, n) in [
            ("plain_text", scans(&|| doc.plain_text())),
            ("to_csv", scans(&|| doc.to_csv())),
            ("to_markdown", scans(&|| doc.to_markdown())),
        ] {
            assert!(n <= 2, "{name} scanned date formats {n} times for 100 cells");
        }
    }
}
