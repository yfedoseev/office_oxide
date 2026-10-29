use super::XlsxDocument;
use super::cell::{Cell, CellValue};
use super::date;
use super::numfmt;
use super::worksheet::Row;
use crate::ir::SheetKind;
use crate::limits::TextBudget;
use crate::stream::Prefixed;
use std::fmt::{self, Write};

/// The heading every renderer gives a pivot's cached source data.
pub(crate) fn pivot_cache_title(pivot_name: &str) -> String {
    format!("{pivot_name} (cached pivot source data)")
}

/// The line every renderer appends when a pivot cache stopped
/// materialising at its record, field or text bound.
pub(crate) const PIVOT_CACHE_TRUNCATED_NOTICE: &str = "[pivot cache truncated: a record, field or text limit was reached — later records are missing]";

impl XlsxDocument {
    /// Every pivot whose cached source data was materialised (its source
    /// is not in the workbook), as `(pivot name, data)`, in sheet order.
    pub(crate) fn pivot_caches(
        &self,
    ) -> impl Iterator<Item = (&str, &super::worksheet::PivotCacheData)> {
        self.worksheets
            .iter()
            .flat_map(|ws| &ws.pivot_tables)
            .filter_map(|p| Some((p.name.as_str(), p.cache_data.as_deref()?)))
    }
}

/// What an unused column of a Markdown table row costs: `" | "`.
pub(crate) const MARKDOWN_PAD_CHARS: usize = 3;

/// The line every renderer gives a sheet whose threaded comments stopped at
/// the document's text budget.
pub(crate) const COMMENTS_TRUNCATED_NOTICE: &str = "[comments truncated: the threaded comments exceed the character budget (office_oxide::limits::set_max_text_chars) — later threads are missing]";

/// What a comment adds to a rendering: its cell reference, author and text,
/// and the note that holds them in the IR. The author is stored once and
/// named by every comment, and a comments part may be named by every sheet,
/// so this is what the renderers charge to the budget for each: a comment of
/// one character is still a note.
pub(crate) fn comment_chars(c: &super::worksheet::SheetComment) -> usize {
    std::mem::size_of::<crate::ir::Note>()
        .saturating_add(c.cell_ref.len())
        .saturating_add(c.author.as_ref().map_or(0, |a| a.len()))
        .saturating_add(c.text.len())
}

/// `B2 (Author)` — the same marker `to_ir()` puts on a comment's endnote.
pub(crate) fn comment_marker(cell_ref: &str, author: Option<&str>) -> String {
    match author {
        Some(a) => format!("{cell_ref} ({a})"),
        None => cell_ref.to_string(),
    }
}

/// The workbook-level tables every cell render needs: which cell styles are
/// dates, and the custom `<numFmt>` codes behind the others, each compiled the
/// first time a cell needs it. They depend on the stylesheet alone, so they
/// are built once per document and shared by every sheet's render (building
/// them per sheet made a workbook of many sheets and many cell formats cost
/// their product).
#[derive(Debug, Clone)]
pub(crate) struct FormatTables {
    /// The stylesheet shape these were built from, to notice a stylesheet
    /// changed through the public field since.
    shape: (usize, usize),
    date_indices: std::collections::HashSet<u32>,
    /// Per number-format id, the first style that uses it (its code is looked
    /// up through that style) and the code compiled on first use.
    compiled:
        std::collections::HashMap<u32, (u32, std::sync::OnceLock<Option<numfmt::CompiledFormat>>)>,
}

thread_local! {
    /// How many times [`FormatTables`] were built on this thread.
    #[cfg(test)]
    pub(crate) static FORMAT_TABLE_BUILDS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

impl FormatTables {
    fn shape_of(doc: &XlsxDocument) -> (usize, usize) {
        doc.styles
            .as_ref()
            .map_or((0, 0), |s| (s.cell_formats.len(), s.number_formats.len()))
    }

    fn build(doc: &XlsxDocument) -> Self {
        #[cfg(test)]
        FORMAT_TABLE_BUILDS.with(|n| n.set(n.get() + 1));
        let mut compiled = std::collections::HashMap::new();
        if let Some(styles) = doc.styles.as_ref() {
            for idx in 0..styles.cell_formats.len() as u32 {
                if let (Some(fmt_id), Some(_)) =
                    (styles.number_format_id_for(idx), styles.number_format_override_for(idx))
                {
                    compiled
                        .entry(fmt_id)
                        .or_insert_with(|| (idx, std::sync::OnceLock::new()));
                }
            }
        }
        Self {
            shape: Self::shape_of(doc),
            date_indices: doc.date_style_indices(),
            compiled,
        }
    }
}

/// What a document keeps between renders: the [`FormatTables`] and the
/// theme parsed from its raw part.
#[derive(Debug, Default)]
pub(crate) struct RenderCache {
    tables: std::sync::OnceLock<FormatTables>,
    theme: std::sync::OnceLock<Option<crate::core::theme::Theme>>,
}

/// A copy of a document starts with an empty cache: its tables are derived
/// again from the copy's own `styles` and theme, so editing a copy cannot
/// leave it with the tables of the original.
impl Clone for RenderCache {
    fn clone(&self) -> Self {
        Self::default()
    }
}

impl RenderCache {
    pub(crate) fn theme(&self, raw: Option<&[u8]>) -> Option<&crate::core::theme::Theme> {
        self.theme
            .get_or_init(|| raw.and_then(|d| crate::core::theme::Theme::parse(d).ok()))
            .as_ref()
    }
}

/// Per-render cell formatting state over the document's [`FormatTables`]
/// rather than a date re-check of the format code and a re-parse of it for
/// every cell. Every direct renderer and `to_ir()` format through this.
pub(crate) struct CellFormatter<'a> {
    doc: &'a XlsxDocument,
    tables: std::borrow::Cow<'a, FormatTables>,
}

impl<'a> CellFormatter<'a> {
    pub(crate) fn new(doc: &'a XlsxDocument) -> Self {
        let cached = doc
            .render_cache
            .tables
            .get_or_init(|| FormatTables::build(doc));
        // `styles` is a public field: tables built for a stylesheet that has
        // since changed shape are not used.
        let tables = if cached.shape == FormatTables::shape_of(doc) {
            std::borrow::Cow::Borrowed(cached)
        } else {
            std::borrow::Cow::Owned(FormatTables::build(doc))
        };
        Self { doc, tables }
    }

    pub(crate) fn date_indices(&self) -> &std::collections::HashSet<u32> {
        &self.tables.date_indices
    }

    /// Append `cell`'s display text to `buf`.
    pub(crate) fn write(&self, cell: &Cell, buf: &mut String) {
        let doc = self.doc;
        let tables = &*self.tables;
        doc.write_value_with(
            cell,
            buf,
            |idx| tables.date_indices.contains(&idx),
            |n, _idx, fmt_id| {
                let code = tables.compiled.get(&fmt_id).and_then(|(first, once)| {
                    once.get_or_init(|| {
                        doc.styles
                            .as_ref()
                            .and_then(|s| s.number_format_override_for(*first))
                            .and_then(numfmt::compile_custom)
                    })
                    .as_ref()
                });
                numfmt::apply_format_compiled(n, fmt_id, code)
            },
        );
    }

    /// `cell`'s display text.
    pub(crate) fn text(&self, cell: &Cell) -> String {
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
    ///
    /// The text is produced sheet by sheet straight into the result, with
    /// nothing held but the row being laid out; [`Self::write_plain_text`]
    /// writes it to a stream instead.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        // Writing to a `String` cannot fail.
        let _ = self.render_plain_text(&mut out);
        out
    }

    /// Write [`Self::plain_text`] to `out` as it is produced, without
    /// building it in memory.
    pub fn write_plain_text<W: std::io::Write + ?Sized>(&self, out: &mut W) -> std::io::Result<()> {
        crate::stream::render_to_io(out, |sink| self.render_plain_text(sink))
    }

    pub(crate) fn render_plain_text<W: Write + ?Sized>(&self, out: &mut W) -> fmt::Result {
        // One text budget for the whole document: a shared string
        // referenced from every cell is rendered once per cell, and
        // nothing else bounds that product (see `crate::limits`).
        let _scale = crate::limits::InputSizeScope::new(self.input_bytes);
        let mut budget = TextBudget::new();
        // The parts of the text are separated by a blank line; the
        // separator goes in front of each part but the first.
        let mut first = true;
        let mut begin = |out: &mut W| -> fmt::Result {
            if !std::mem::take(&mut first) {
                out.write_str("\n\n")?;
            }
            Ok(())
        };
        for (i, ws) in self.worksheets.iter().enumerate() {
            begin(out)?;
            out.write_str(&ws.name)?;
            // Page headers above the cells and footers below, as the IR
            // renderers place a section's headers and footers.
            let [fh, oh, eh, ff, of, ef] = ws.header_footer.active(&ws.name);
            for h in [fh, oh, eh].into_iter().flatten() {
                out.write_char('\n')?;
                out.write_str(&h)?;
            }
            // A chart sheet's content is its chart's text.
            for text in &ws.chart_text {
                if !budget.charge(text.trim().len()) {
                    break;
                }
                out.write_char('\n')?;
                out.write_str(text.trim())?;
            }
            // The cells follow on a line of their own, when there are any.
            self.render_sheet_plain_within(i, &mut budget, &mut Prefixed::new(out, "\n"))?;
            for f in [ff, of, ef].into_iter().flatten() {
                out.write_char('\n')?;
                out.write_str(&f)?;
            }
            if budget.exhausted() {
                out.write_char('\n')?;
                out.write_str(&budget.notice())?;
                break;
            }
            // Cell comments are document content; `to_ir()` carries them
            // as endnotes, and this direct renderer dropped them.
            for c in ws.comments.iter() {
                // The author's name is repeated by every comment of theirs.
                if !budget.charge(comment_chars(c)) {
                    break;
                }
                out.write_char('\n')?;
                write!(out, "{}: {}", comment_marker(&c.cell_ref, c.author.as_deref()), c.text)?;
            }
            if ws.comments_truncated {
                out.write_char('\n')?;
                out.write_str(COMMENTS_TRUNCATED_NOTICE)?;
            }
            if let Some(damage) = &ws.damage {
                out.write_char('\n')?;
                out.write_str(&damaged_notice(&ws.name, damage))?;
            }
            if budget.exhausted() {
                out.write_char('\n')?;
                out.write_str(&budget.notice())?;
                break;
            }
            // Text boxes and WordArt drawn on the sheet reached `to_ir()`
            // as text boxes and this renderer not at all — a sheet whose
            // content is a drawn note came back as its name alone.
            for ts in ws.text_shapes.iter() {
                if !ts.text.trim().is_empty() {
                    // Shared by every sheet that names the drawing, written
                    // once for each: paid for like any other text.
                    if !budget.charge(ts.text.trim().len()) {
                        break;
                    }
                    out.write_char('\n')?;
                    out.write_str(ts.text.trim())?;
                }
            }
            if budget.exhausted() {
                out.write_char('\n')?;
                out.write_str(&budget.notice())?;
                break;
            }
        }
        // `to_markdown()` already surfaces chart text (axis titles, series
        // names); `plain_text()` silently dropped it entirely.
        for text in &self.chart_text {
            if !text.trim().is_empty() {
                begin(out)?;
                out.write_str(text.trim())?;
            }
        }
        // Pivots that name one cache each render it, so the output is
        // charged to the budget however many of them share the data.
        let mut pivots_spent = false;
        for (name, data) in self.pivot_caches() {
            if budget.exhausted() {
                break;
            }
            begin(out)?;
            out.write_str(&pivot_cache_title(name))?;
            for row in std::iter::once(&data.fields).chain(&data.records) {
                let line = row.join("\t");
                if !budget.charge(line.len() + 1) {
                    pivots_spent = true;
                    break;
                }
                out.write_char('\n')?;
                out.write_str(&line)?;
            }
            if data.truncated {
                out.write_char('\n')?;
                out.write_str(PIVOT_CACHE_TRUNCATED_NOTICE)?;
            }
        }
        if pivots_spent {
            begin(out)?;
            out.write_str(&budget.notice())?;
        }
        write_report_notices(out, &mut budget, &mut begin, &self.report_notices(), false)
    }

    /// Extract a single sheet as plain text.
    pub fn sheet_plain_text(&self, sheet_index: usize) -> Option<String> {
        let _scale = crate::limits::InputSizeScope::new(self.input_bytes);
        let mut budget = TextBudget::new();
        let mut text = String::new();
        // Writing to a `String` cannot fail.
        if !self
            .render_sheet_plain_within(sheet_index, &mut budget, &mut text)
            .ok()?
        {
            return None;
        }
        if budget.exhausted() {
            text.push('\n');
            text.push_str(&budget.notice());
        }
        Some(text)
    }

    /// Write the rows of sheet `sheet_index` to `out`, each cell charged to
    /// `budget` before it is written. `Ok(false)` when there is no such sheet.
    fn render_sheet_plain_within<W: Write + ?Sized>(
        &self,
        sheet_index: usize,
        budget: &mut TextBudget,
        out: &mut W,
    ) -> Result<bool, fmt::Error> {
        let Some(ws) = self.worksheets.get(sheet_index) else {
            return Ok(false);
        };
        let fmt = CellFormatter::new(self);
        // The tabs in front of a cell and the cell, the largest piece held.
        let mut piece = String::new();
        for (row_idx, row) in ws.rows.iter().enumerate() {
            if row_idx > 0 {
                out.write_char('\n')?;
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
                piece.clear();
                piece.extend(std::iter::repeat_n('\t', tabs));
                fmt.write(cell, &mut piece);
                if !budget.charge(piece.len()) {
                    return Ok(true);
                }
                out.write_str(&piece)?;
                next_col = next_col.max(col.saturating_add(1));
            }
        }
        Ok(true)
    }

    /// Convert the workbook to IR one section at a time, in the order of
    /// [`Document::to_ir`](crate::Document::to_ir)`().sections`: each worksheet,
    /// then the chart text, the pivot caches and the sheets that could not
    /// be read.
    ///
    /// Nothing is converted until the iterator is advanced, and nothing it
    /// has handed out is kept, so a caller that uses and drops each
    /// [`SheetIr`](super::SheetIr) before asking for the next holds about
    /// one sheet's IR whatever the workbook's size. This is how a browser
    /// holds a page, and how a streaming reader holds a worksheet.
    ///
    /// The allowances that bound the IR's memory are therefore a sheet's,
    /// fresh for each: the empty cells that pad its rows
    /// (`max_sheet_ir_padding_bytes`)
    /// and the picture bytes it copies
    /// (`max_image_bytes`). What
    /// bounds the run as a whole stays document-wide: the text budget
    /// ([`limits::max_text_chars`](crate::limits::max_text_chars)), which the
    /// padding is charged to as well, and the total of picture bytes copied
    /// (`max_total_image_bytes`).
    /// A sheet that does not fit its own allowance is cut, with a notice in
    /// the section, a [`SheetIr::warnings`](super::SheetIr::warnings) entry
    /// and [`SheetIr::text_truncated`](super::SheetIr::text_truncated), and
    /// the next sheet is converted whole; a document-wide limit ends the
    /// iteration the same way. [`Document::to_ir`](crate::Document::to_ir)
    /// holds every sheet at once, so it keeps document-wide allowances.
    pub fn sheets_ir(&self) -> super::SheetIrs<'_> {
        super::SheetIrs::new(self)
    }

    /// Convert worksheet `sheet_index` alone, with a sheet's own allowances
    /// (see [`Self::sheets_ir`]); `None` when there is no such worksheet.
    pub fn sheet_ir(&self, sheet_index: usize) -> Option<super::SheetIr> {
        crate::convert_xlsx::sheet_to_ir(self, sheet_index)
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
        let _scale = crate::limits::InputSizeScope::new(self.input_bytes);
        let mut budget = TextBudget::new();
        let fmt = CellFormatter::new(self);

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
    ///
    /// The text is produced sheet by sheet straight into the result, with
    /// nothing held but the row being laid out; [`Self::write_markdown`]
    /// writes it to a stream instead.
    pub fn to_markdown(&self) -> String {
        let mut out = String::new();
        // Writing to a `String` cannot fail.
        let _ = self.render_markdown(&mut out);
        out
    }

    /// Write [`Self::to_markdown`] to `out` as it is produced, without
    /// building it in memory.
    pub fn write_markdown<W: std::io::Write + ?Sized>(&self, out: &mut W) -> std::io::Result<()> {
        crate::stream::render_to_io(out, |sink| self.render_markdown(sink))
    }

    pub(crate) fn render_markdown<W: Write + ?Sized>(&self, out: &mut W) -> fmt::Result {
        let _scale = crate::limits::InputSizeScope::new(self.input_bytes);
        let mut budget = TextBudget::new();
        // The parts of the Markdown are separated by a blank line; the
        // separator goes in front of each part but the first.
        let mut first = true;
        let mut begin = |out: &mut W| -> fmt::Result {
            if !std::mem::take(&mut first) {
                out.write_str("\n\n")?;
            }
            Ok(())
        };
        for (i, ws) in self.worksheets.iter().enumerate() {
            let [fh, oh, eh, ff, of, ef] = ws.header_footer.active(&ws.name);
            if markdown_column_count(ws) > 0 {
                begin(out)?;
                self.render_sheet_markdown_within(i, &mut budget, out)?;
            } else if !ws.comments.is_empty()
                || ws.text_shapes.iter().any(|t| !t.text.trim().is_empty())
                || !ws.header_footer.is_empty()
                || !matches!(ws.kind, SheetKind::Worksheet | SheetKind::Macro)
            {
                // No cells, but comments, drawn text or page
                // headers: they still belong under the sheet's heading.
                // A chart sheet or dialog sheet has no cells at all;
                // its tab is still a sheet of the workbook.
                begin(out)?;
                write!(out, "## {}", ws.name)?;
            }
            for text in &ws.chart_text {
                if !budget.charge(text.trim().len()) {
                    break;
                }
                begin(out)?;
                out.write_str(text.trim())?;
            }
            for hf in [fh, oh, eh, ff, of, ef].into_iter().flatten() {
                begin(out)?;
                out.write_str(&crate::core::markdown::escape_text(&hf))?;
            }
            if budget.exhausted() {
                begin(out)?;
                out.write_str(&budget.notice())?;
                break;
            }
            for c in ws.comments.iter() {
                if !budget.charge(comment_chars(c)) {
                    break;
                }
                begin(out)?;
                write!(
                    out,
                    "> **{}:** {}",
                    comment_marker(&c.cell_ref, c.author.as_deref()),
                    c.text.trim()
                )?;
            }
            if ws.comments_truncated {
                begin(out)?;
                out.write_str(COMMENTS_TRUNCATED_NOTICE)?;
            }
            if let Some(damage) = &ws.damage {
                begin(out)?;
                out.write_str(&damaged_notice(&ws.name, damage))?;
            }
            if budget.exhausted() {
                begin(out)?;
                out.write_str(&budget.notice())?;
                break;
            }
            for ts in ws.text_shapes.iter() {
                if !ts.text.trim().is_empty() {
                    if !budget.charge(ts.text.trim().len()) {
                        break;
                    }
                    begin(out)?;
                    out.write_str(ts.text.trim())?;
                }
            }
            if budget.exhausted() {
                begin(out)?;
                out.write_str(&budget.notice())?;
                break;
            }
        }
        // Charts: emit each chart's extracted text under a "## Chart N" heading
        // so its words appear in markdown / search / PDF without needing a
        // graphical chart renderer.
        for (i, text) in self.chart_text.iter().enumerate() {
            if !text.trim().is_empty() {
                begin(out)?;
                write!(out, "## Chart {}\n\n{}", i + 1, text)?;
            }
        }
        // Pivots that name one cache each render it, so the output is
        // charged to the budget however many of them share the data.
        let mut pivots_spent = false;
        for (name, data) in self.pivot_caches() {
            if budget.exhausted() {
                break;
            }
            let cols = data.fields.len().max(1);
            let md_row = |row: &[String]| {
                let mut line = String::from("|");
                for i in 0..cols {
                    let v = row.get(i).map(String::as_str).unwrap_or("");
                    line.push(' ');
                    line.push_str(&crate::core::markdown::escape_cell(v));
                    line.push_str(" |");
                }
                line
            };
            // The heading, the header row and the separator are paid for
            // before they are written; a pivot whose header does not fit is
            // not written, and no later one is.
            let heading = format!("## {}\n\n", pivot_cache_title(name));
            let header = md_row(&data.fields);
            let separator = format!("|{}", " --- |".repeat(cols));
            if !budget.charge(heading.len() + header.len() + separator.len() + 2) {
                pivots_spent = true;
                break;
            }
            begin(out)?;
            out.write_str(&heading)?;
            out.write_str(&header)?;
            out.write_char('\n')?;
            out.write_str(&separator)?;
            for rec in &data.records {
                let line = md_row(rec);
                if !budget.charge(line.len() + 1) {
                    pivots_spent = true;
                    break;
                }
                out.write_char('\n')?;
                out.write_str(&line)?;
            }
            if data.truncated {
                out.write_str("\n\n")?;
                out.write_str(PIVOT_CACHE_TRUNCATED_NOTICE)?;
            }
        }
        if pivots_spent {
            begin(out)?;
            out.write_str(&budget.notice())?;
        }
        write_report_notices(out, &mut budget, &mut begin, &self.report_notices(), true)
    }

    /// Convert specific sheet to markdown.
    pub fn sheet_to_markdown(&self, sheet_index: usize) -> Option<String> {
        let _scale = crate::limits::InputSizeScope::new(self.input_bytes);
        let mut budget = TextBudget::new();
        let mut md = String::new();
        // Writing to a `String` cannot fail.
        if !self
            .render_sheet_markdown_within(sheet_index, &mut budget, &mut md)
            .ok()?
        {
            return None;
        }
        if budget.exhausted() {
            md.push_str("\n\n");
            md.push_str(&budget.notice());
        }
        Some(md)
    }

    /// Write sheet `sheet_index` as Markdown to `out`, each cell charged to
    /// `budget` before it is written; nothing for a sheet without cells.
    /// `Ok(false)` when there is no such sheet.
    fn render_sheet_markdown_within<W: Write + ?Sized>(
        &self,
        sheet_index: usize,
        budget: &mut TextBudget,
        out: &mut W,
    ) -> Result<bool, fmt::Error> {
        let Some(ws) = self.worksheets.get(sheet_index) else {
            return Ok(false);
        };
        let col_count = markdown_column_count(ws);
        if col_count == 0 {
            return Ok(true);
        }
        // Every cell text passes through here; a spent budget ends the
        // sheet at the cell that spent it.
        let fmt = CellFormatter::new(self);

        // A "document style" sheet (see `rows_read_as_prose`) is emitted a
        // paragraph per row instead of a tall, narrow GFM table that reads
        // badly and round-trips badly through markdown→IR→office — the
        // same decision, and the same row layout, as `to_ir()`.
        let prose = rows_read_as_prose(ws.rows.iter().map(|r| {
            r.cells
                .iter()
                .map(|c| fmt.text(c).chars().count())
                .filter(|&n| n > 0)
                .collect::<Vec<_>>()
        }));
        let cell_text = |cell: &Cell, budget: &mut TextBudget| -> Option<String> {
            let text = fmt.text(cell);
            budget
                .charge(text.len())
                .then(|| crate::core::markdown::escape_cell(&text))
        };
        if prose {
            // The paragraphs, each followed by a blank line, but for the
            // last: the line written is held back until the next one comes.
            let heading = format!("## {}\n\n", ws.name);
            let mut heading_written = false;
            let mut held: Option<String> = None;
            'rows: for row in &ws.rows {
                // Every non-empty cell of the row, tab-separated.
                let mut line = String::new();
                for cell in &row.cells {
                    let Some(text) = cell_text(cell, budget) else {
                        break 'rows;
                    };
                    if text.trim().is_empty() {
                        continue;
                    }
                    if !line.is_empty() {
                        line.push('\t');
                    }
                    // `cell_text` escaped for a table cell; prose keeps
                    // its line breaks.
                    line.push_str(&text.replace("<br>", "\n"));
                }
                if !line.is_empty() {
                    if !heading_written {
                        out.write_str(&heading)?;
                        heading_written = true;
                    }
                    if let Some(previous) = held.replace(line) {
                        out.write_str(&previous)?;
                        out.write_str("\n\n")?;
                    }
                }
            }
            match held {
                Some(last) => out.write_str(last.trim_end())?,
                None => out.write_str(heading.trim_end())?,
            }
            return Ok(true);
        }

        // The lines of the table are separated by a line break.
        let mut first_line = true;
        let mut line_out = |out: &mut W, line: &str| -> fmt::Result {
            if !std::mem::take(&mut first_line) {
                out.write_char('\n')?;
            }
            out.write_str(line)
        };

        // Sheet name as heading
        line_out(out, &format!("## {}", ws.name))?;
        line_out(out, "")?;

        // First row as header
        // A row keeps the cells that fit the budget (the rest empty) and
        // reports that the budget is spent, so the table ends after it.
        // Each cell goes under its own column's header (§18.3.1.4 `r`).
        // Every row is laid out the sheet's full width, so the columns a
        // row does not use are output too: " | " for each one. They are
        // charged like text, which bounds a sheet whose cells sit at A1
        // and XFD1 however many rows it stores, as `to_csv` does.
        let row_line = |row: &Row, budget: &mut TextBudget| -> (Option<String>, bool) {
            let mut cells: Vec<String> = vec![String::new(); col_count];
            if !budget.charge(col_count.saturating_mul(MARKDOWN_PAD_CHARS)) {
                return (None, true);
            }
            let mut spent = false;
            for c in &row.cells {
                match cell_text(c, budget) {
                    Some(t) => place(&mut cells, c.reference.col, t),
                    None => {
                        spent = true;
                        break;
                    },
                }
            }
            (Some(format!("| {} |", cells.join(" | "))), spent)
        };
        let (header, spent) = row_line(&ws.rows[0], budget);
        let Some(header) = header else {
            // Not even the header fits: the sheet is its heading alone.
            return Ok(true);
        };
        line_out(out, &header)?;

        // Separator row, as wide as the table.
        if !budget.charge(col_count.saturating_mul(MARKDOWN_PAD_CHARS + 2)) {
            return Ok(true);
        }
        let sep: Vec<&str> = vec!["---"; col_count];
        line_out(out, &format!("| {} |", sep.join(" | ")))?;

        // Data rows
        if !spent {
            for row in ws.rows.iter().skip(1) {
                let (line, spent) = row_line(row, budget);
                if let Some(line) = line {
                    line_out(out, &line)?;
                }
                if spent {
                    break;
                }
            }
        }

        Ok(true)
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
                if let Some(f) = cell.formula() {
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
/// How many columns a sheet's Markdown table has; zero when the sheet
/// renders as nothing at all (no stored rows, or no cells in them).
fn markdown_column_count(ws: &super::worksheet::Worksheet) -> usize {
    if ws.rows.is_empty() {
        0
    } else {
        compute_column_count(&ws.rows)
    }
}

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

/// The line every renderer emits for a sheet whose XML was malformed part-way
/// and which was read only up to the damage.
pub(crate) fn damaged_notice(name: &str, damage: &str) -> String {
    inert(&format!("[sheet {name:?} cut short: {damage}]"))
}

/// The line every renderer emits for a sheet that could not be read, so
/// a workbook missing a sheet never passes for a complete one.
pub(crate) fn unreadable_notice(name: &str, err: &str) -> String {
    inert(&format!("[unreadable sheet {name:?}: {err}]"))
}

/// `s` as one line: every control character, and the Unicode line and
/// paragraph separators, become a space. A notice echoes names and reasons
/// the file controls, and one that could end its line could start another
/// (a heading, a list item) in any rendering.
pub(crate) fn inert(s: &str) -> String {
    crate::core::unreadable::one_line(s)
}

/// What could not be read, as every renderer reports it: one entry for each
/// unreadable sheet listed, one for the sheets left unlisted and one for the
/// parts (relationship, drawing, comments, table, pivot, chart and picture
/// parts). They are bounded (see [`XlsxDocument::unreadable_sheets`]), cut and
/// charged to the text budget by every renderer before it writes them.
pub(crate) struct ReportNotice {
    /// The heading of the entry's section: a fixed label, never text from the
    /// file.
    pub(crate) title: String,
    /// The in-band lines, each a bracketed notice.
    pub(crate) lines: Vec<String>,
    /// The `Metadata::warnings` entries for the same.
    pub(crate) warnings: Vec<String>,
}

impl ReportNotice {
    /// What the entry costs to write, in characters like the rest of the
    /// text budget: its heading and lines. (Each is cut at 256 characters
    /// and there are at most 201 of them, so counting is bounded.)
    pub(crate) fn chars(&self) -> usize {
        self.lines.iter().fold(self.title.chars().count(), |n, l| {
            n.saturating_add(l.chars().count()).saturating_add(4)
        })
    }
}

impl XlsxDocument {
    /// The entries reporting what could not be read, in the order every
    /// renderer writes them: the unreadable sheets, then the parts.
    pub(crate) fn report_notices(&self) -> Vec<ReportNotice> {
        let mut out: Vec<ReportNotice> = self
            .unreadable_sheets
            .iter()
            .map(|(name, err)| ReportNotice {
                // The name is in the line, where every rendering writes it as
                // text; a heading would write it as markup.
                title: "Unreadable sheet".to_string(),
                lines: vec![unreadable_notice(name, err)],
                warnings: vec![format!("skipped unreadable sheet {name:?}: {err}")],
            })
            .collect();
        if self.unreadable_sheets_omitted > 0 {
            let more =
                format!("and {} more unreadable sheets not listed", self.unreadable_sheets_omitted);
            out.push(ReportNotice {
                title: "Unreadable sheets".to_string(),
                lines: vec![format!("[{more}]")],
                warnings: vec![more],
            });
        }
        if !self.unreadable_parts.is_empty() || self.unreadable_parts_omitted > 0 {
            let mut warnings: Vec<String> = self
                .unreadable_parts
                .iter()
                .map(|(part, err)| format!("skipped unreadable part {part}: {err}"))
                .collect();
            if self.unreadable_parts_omitted > 0 {
                warnings.push(format!(
                    "and {} more unreadable parts not listed",
                    self.unreadable_parts_omitted
                ));
            }
            out.push(ReportNotice {
                title: "Unreadable parts".to_string(),
                lines: warnings.iter().map(|w| inert(&format!("[{w}]"))).collect(),
                warnings,
            });
        }
        out
    }
}

/// Write the entries of `notices` that the budget pays for, each by `write`,
/// and the budget's own notice (unless something already wrote it) where it
/// runs out. `begin` starts each part of the output.
fn write_report_notices<W: Write + ?Sized>(
    out: &mut W,
    budget: &mut TextBudget,
    begin: &mut dyn FnMut(&mut W) -> fmt::Result,
    notices: &[ReportNotice],
    markdown: bool,
) -> fmt::Result {
    for n in notices {
        let was_spent = budget.exhausted();
        if !budget.charge(n.chars()) {
            if !was_spent {
                begin(out)?;
                out.write_str(&budget.notice())?;
            }
            break;
        }
        if markdown {
            begin(out)?;
            let lines: Vec<String> = n
                .lines
                .iter()
                .map(|l| crate::core::markdown::escape_text(l))
                .collect();
            write!(out, "## {}\n\n{}", n.title, lines.join("\n\n"))?;
        } else {
            for line in &n.lines {
                begin(out)?;
                out.write_str(line)?;
            }
        }
    }
    Ok(())
}

/// Whether a sheet reads as prose — flowing text laid out one paragraph
/// per row — rather than as a grid. `rows` yields, per row, the character
/// count of each non-empty cell.
///
/// Prose means: at least three non-empty rows, at least 80 % of them with a
/// single non-empty cell, and at least one cell of prose length (over 20
/// characters). A single column of short values (names, numbers, codes) is
/// data and stays a table. `to_ir()` and `to_markdown()` both decide
/// through this one predicate; they used to use two different rules, so
/// one rendered a sheet as a table and the other as paragraphs.
pub(crate) fn rows_read_as_prose<R, C>(rows: R) -> bool
where
    R: IntoIterator<Item = C>,
    C: IntoIterator<Item = usize>,
{
    const PROSE_CHARS: usize = 20;
    let (mut nonempty, mut single, mut long) = (0usize, 0usize, false);
    for row in rows {
        let mut n = 0usize;
        for len in row {
            n += 1;
            long |= len > PROSE_CHARS;
        }
        if n == 0 {
            continue;
        }
        nonempty += 1;
        if n == 1 {
            single += 1;
        }
    }
    long && nonempty >= 3 && single * 100 >= nonempty * 80
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
            input_bytes: 0,
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
            unreadable_sheets_omitted: 0,
            unreadable_parts: Vec::new(),
            unreadable_parts_omitted: 0,
            cosmetic_parts: Vec::new(),
            styles_data: None,
            theme_data: None,
            render_cache: Default::default(),
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
            formulas_dropped: 0,
            damage: None,
            page_setup: None,
            images: std::sync::Arc::from([]),
            cell_images: Vec::new(),
            comments: std::sync::Arc::from([]),
            comments_truncated: false,
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
            }]
            .into(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            state: super::super::SheetState::Visible,
            header_footer: Default::default(),
            tables: Vec::new(),
            pivot_tables: Vec::new(),
            kind: crate::ir::SheetKind::Worksheet,
            chart_text: Vec::new(),
        };
        let doc = XlsxDocument {
            input_bytes: 0,
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
            unreadable_sheets_omitted: 0,
            unreadable_parts: Vec::new(),
            unreadable_parts_omitted: 0,
            cosmetic_parts: Vec::new(),
            styles_data: None,
            theme_data: None,
            render_cache: Default::default(),
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
                cells: vec![Cell::new(
                    super::super::CellRef { col: 0, row: 0 },
                    CellValue::String(text.to_string()),
                    None,
                )],
            }],
            merged_cells: Vec::new(),
            hyperlinks: Vec::new(),
            formulas_dropped: 0,
            damage: None,
            page_setup: None,
            images: std::sync::Arc::from([]),
            cell_images: Vec::new(),
            comments: std::sync::Arc::from([]),
            comments_truncated: false,
            text_shapes: std::sync::Arc::from([]),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            state: super::super::SheetState::Visible,
            header_footer: Default::default(),
            tables: Vec::new(),
            pivot_tables: Vec::new(),
            kind: crate::ir::SheetKind::Worksheet,
            chart_text: Vec::new(),
        };
        let doc = XlsxDocument {
            input_bytes: 0,
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
            unreadable_sheets_omitted: 0,
            unreadable_parts: Vec::new(),
            unreadable_parts_omitted: 0,
            cosmetic_parts: Vec::new(),
            styles_data: None,
            theme_data: None,
            render_cache: Default::default(),
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
            input_bytes: 0,
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
            unreadable_sheets_omitted: 0,
            unreadable_parts: Vec::new(),
            unreadable_parts_omitted: 0,
            cosmetic_parts: Vec::new(),
            styles_data: None,
            theme_data: None,
            render_cache: Default::default(),
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

    /// A workbook of many sheets over a stylesheet of many cell formats: the
    /// date styles and compiled number formats are workbook-level, so every
    /// rendering builds them once, not once per sheet.
    #[test]
    fn test_workbook_level_format_tables_are_built_once_for_all_sheets() {
        use crate::xlsx::test_support::{open_bytes, zip_parts};
        const SHEETS: usize = 40;
        const FORMATS: usize = 60;
        let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let mut workbook = format!(r#"<workbook xmlns="{ns}" xmlns:r="{rel}"><sheets>"#);
        let mut rels = String::from(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
        for i in 1..=SHEETS {
            workbook.push_str(&format!(r#"<sheet name="S{i}" sheetId="{i}" r:id="rId{i}"/>"#));
            rels.push_str(&format!(
                r#"<Relationship Id="rId{i}" Type="{rel}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
            ));
            parts.push((
                format!("xl/worksheets/sheet{i}.xml"),
                format!(
                    r#"<worksheet xmlns="{ns}"><sheetData><row r="1"><c r="A1" s="{}"><v>{}</v></c></row></sheetData></worksheet>"#,
                    i % FORMATS + 1,
                    38000 + i
                )
                .into_bytes(),
            ));
        }
        workbook.push_str("</sheets></workbook>");
        rels.push_str(&format!(
            r#"<Relationship Id="rIdS" Type="{rel}/styles" Target="styles.xml"/></Relationships>"#
        ));
        let mut styles = format!(r#"<styleSheet xmlns="{ns}"><numFmts count="{FORMATS}">"#);
        for i in 0..FORMATS {
            styles.push_str(&format!(
                r#"<numFmt numFmtId="{}" formatCode="yyyy-mm-dd;@;0.0{i}"/>"#,
                164 + i
            ));
        }
        styles
            .push_str(&format!(r#"</numFmts><cellXfs count="{}"><xf numFmtId="0"/>"#, FORMATS + 1));
        for i in 0..FORMATS {
            styles.push_str(&format!(r#"<xf numFmtId="{}" applyNumberFormat="1"/>"#, 164 + i));
        }
        styles.push_str("</cellXfs></styleSheet>");
        parts.push(("xl/workbook.xml".into(), workbook.into_bytes()));
        parts.push(("xl/_rels/workbook.xml.rels".into(), rels.into_bytes()));
        parts.push(("xl/styles.xml".into(), styles.into_bytes()));
        let refs: Vec<(&str, &[u8])> = parts
            .iter()
            .map(|(n, d)| (n.as_str(), d.as_slice()))
            .collect();
        let doc = open_bytes(zip_parts(&refs));

        FORMAT_TABLE_BUILDS.with(|n| n.set(0));
        let plain = doc.plain_text();
        assert!(plain.contains("2004-01-"), "{plain}");
        let _ = doc.to_markdown();
        let _ = doc.to_csv();
        let _ = crate::convert_xlsx::xlsx_to_ir(&doc);
        for i in 0..SHEETS {
            assert!(doc.sheet_ir(i).is_some());
            assert!(doc.sheet_plain_text(i).is_some());
            assert!(doc.sheet_to_markdown(i).is_some());
            assert!(doc.sheet_to_csv(i).is_some());
        }
        assert_eq!(doc.sheets_ir().count(), SHEETS);
        let builds = FORMAT_TABLE_BUILDS.with(|n| n.get());
        assert_eq!(builds, 1, "format tables built {builds} times for {SHEETS} sheets");

        // A stylesheet changed through the public field is not served from
        // the tables built for the old one.
        let mut changed = doc.clone();
        changed.styles.as_mut().unwrap().cell_formats.truncate(1);
        assert!(changed.plain_text().contains("38001"), "{}", changed.plain_text());
        assert!(!changed.plain_text().contains("2004-01-"));
    }
}
