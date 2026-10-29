use crate::format::DocumentFormat;
use crate::ir::*;
use crate::xlsx::range_sweep::{CellRange, RangeSweep};

/// Maximum worksheet rows materialised into the IR per sheet. A worksheet is
/// converted eagerly into an in-memory `Table` (or one `Paragraph` per row),
/// and downstream a single table cannot render to an unbounded PDF, so an
/// uncapped sheet (e.g. a 200k-row stress fixture) explodes memory and stalls
/// rendering. Beyond this many rows the remainder is dropped and a visible
/// truncation notice is appended, rather than hanging. Chosen to preserve
/// ordinary large spreadsheets while bounding pathological ones.
const MAX_ROWS_PER_SHEET: usize = 10_000;

/// How many sheets' names the notice for the sheets a spent budget left out
/// lists; the rest are counted.
const OMITTED_SHEETS_LISTED: usize = 20;

/// Parse a 6-char hex colour like `"FFA500"` into `[r, g, b]`.
fn parse_hex_rgb(s: &str) -> Option<[u8; 3]> {
    let s = s.trim_start_matches('#');
    if s.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some([r, g, b])
}

/// How much of a workbook the IR holds at once, and so what its resource
/// allowances are measured against.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum IrScope {
    /// The whole workbook is built into one IR and held until the caller
    /// drops it: the empty cells that pad rows and the picture bytes are
    /// allowed once for the document.
    Document,
    /// One sheet is built, used and dropped at a time: those allowances are
    /// the sheet's, given back with it. The text budget, which is what the
    /// whole run may output, and the characters charged for padding stay
    /// document-wide, as does a total on picture bytes copied, so that many
    /// sheets cannot multiply the work without bound.
    Sheet,
}

/// Converts the worksheets of one workbook to IR sections, one at a time.
///
/// The state that outlives a sheet is here: the document's text budgets, the
/// formatter that compiles each number format once, the warnings so far.
pub(crate) struct SheetConverter<'d> {
    doc: &'d crate::xlsx::XlsxDocument,
    scope: IrScope,
    formatter: crate::xlsx::text::CellFormatter<'d>,
    // Theme colours resolve cell fills; parsed once per document.
    theme: Option<std::borrow::Cow<'d, crate::core::theme::Theme>>,
    // Single String buffer reused across all cells — clear() keeps the heap
    // allocation; std::mem::take() moves it into TextSpan for non-empty cells.
    buf: String,
    // One text budget for the whole document (see `crate::limits`): a
    // shared string referenced from every cell is copied into every
    // cell's span here.
    budget: crate::limits::TextBudget,
    image_budget: crate::limits::TextBudget,
    // Picture bytes copied in all, whatever the sheets' own allowances: the
    // bound on the work of a workbook whose sheets each draw the same picture.
    total_image_budget: crate::limits::TextBudget,
    padding_budget: crate::limits::IrPaddingBudget,
    // Padding is also paid for in characters, 3 an empty cell as Markdown
    // pays for an absent column, from a budget of its own: the cells' text
    // has spent the text budget by the time the rows are laid out, and the
    // rows whose text fit must not be lost for it. Lowering the text limit
    // still lowers this one.
    padding_chars: crate::limits::TextBudget,
    comments_truncated: bool,
    format_code_omitted: bool,
    // The picture notice belongs to the sheet whose pictures spent the budget.
    image_notice_given: bool,
    total_image_notice_given: bool,
    formulas_truncated: bool,
    // How the sheet just built was cut by padding, if it was: `true` when it
    // was the memory allowance, `false` when the character budget.
    sheet_padding_cut: Option<bool>,
    // A sheet's own allowances are fresh for the next sheet, so what was
    // cut is remembered here for the document's `text_truncated`.
    padding_cut: bool,
    image_cut: bool,
    pivot_truncated: bool,
    unreadable: bool,
    // Sheets were left out of the conversion.
    omitted: bool,
    // Whether the section just built was cut for any reason of its own.
    item_truncated: bool,
    warnings: Vec<String>,
    // The reports of what could not be read, once first asked for.
    reports: Option<Vec<crate::xlsx::text::ReportNotice>>,
}

impl<'d> SheetConverter<'d> {
    pub(crate) fn new(doc: &'d crate::xlsx::XlsxDocument, scope: IrScope) -> Self {
        // Date styles found once and custom number formats compiled once, not
        // re-scanned and re-parsed per cell.
        let _scale = crate::limits::InputSizeScope::new(doc.input_bytes);
        let formatter = crate::xlsx::text::CellFormatter::new(doc);
        let total_image =
            usize::try_from(crate::limits::max_total_image_bytes()).unwrap_or(usize::MAX);
        Self {
            doc,
            scope,
            formatter,
            theme: doc.theme_for_render(),
            buf: String::new(),
            budget: crate::limits::TextBudget::new(),
            image_budget: crate::limits::TextBudget::with_limit(crate::limits::max_image_bytes()),
            total_image_budget: crate::limits::TextBudget::with_limit(total_image),
            padding_budget: Self::fresh_padding_budget(scope),
            padding_chars: crate::limits::TextBudget::new(),
            comments_truncated: false,
            format_code_omitted: false,
            image_notice_given: false,
            total_image_notice_given: false,
            formulas_truncated: false,
            sheet_padding_cut: None,
            padding_cut: false,
            image_cut: false,
            pivot_truncated: false,
            unreadable: false,
            omitted: false,
            item_truncated: false,
            warnings: Vec::new(),
            reports: None,
        }
    }

    fn fresh_padding_budget(scope: IrScope) -> crate::limits::IrPaddingBudget {
        match scope {
            IrScope::Document => crate::limits::IrPaddingBudget::new(),
            IrScope::Sheet => crate::limits::IrPaddingBudget::new_for_sheet(),
        }
    }

    /// Whether what is spent for the whole run ends it: no further sheet is
    /// converted (the one that spent it carries the notice, and the sheets
    /// left out are named by [`Self::omitted_section`]). The memory allowance
    /// for padding cells is not one of these: a later sheet that needs no
    /// padding fits whatever an earlier one spent, and one that does is cut
    /// at the rows it cannot pay for, with a notice of its own.
    pub(crate) fn should_stop(&self) -> bool {
        self.budget.exhausted() || self.padding_chars.exhausted()
    }

    /// The section naming the worksheets from `from` on, which were not
    /// converted because the run's budget was spent. `None` when there are
    /// none. At most [`OMITTED_SHEETS_LISTED`] names are listed.
    pub(crate) fn omitted_section(&mut self, from: usize) -> Option<Section> {
        let left = self.doc.worksheets.get(from..).filter(|l| !l.is_empty())?;
        let mut names: Vec<String> = left
            .iter()
            .take(OMITTED_SHEETS_LISTED)
            .map(|w| format!("{:?}", w.name.chars().take(64).collect::<String>()))
            .collect();
        if left.len() > OMITTED_SHEETS_LISTED {
            names.push(format!("and {} more", left.len() - OMITTED_SHEETS_LISTED));
        }
        let notice = format!(
            "[{} sheet(s) not converted, because the character budget for the workbook was spent \
             (office_oxide::limits::set_max_text_chars): {}]",
            left.len(),
            names.join(", ")
        );
        self.warnings.push(notice.clone());
        self.item_truncated = true;
        self.omitted = true;
        Some(Section {
            title: Some("Sheets not converted".to_string()),
            elements: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(notice))],
                ..Default::default()
            })],
            ..Default::default()
        })
    }

    /// The warnings the sections built so far have raised.
    pub(crate) fn warnings(&self) -> &[String] {
        &self.warnings
    }

    /// Whether anything built so far was cut, for `Metadata::text_truncated`.
    pub(crate) fn text_truncated(&self) -> bool {
        self.unreadable
            || self.omitted
            || self.pivot_truncated
            || self.budget.exhausted()
            || self.padding_cut
            || self.padding_chars.exhausted()
            || self.comments_truncated
            || self.formulas_truncated
            || self.image_cut
            || self.total_image_budget.exhausted()
    }

    /// Whether the section built last was itself cut, or the run has been.
    pub(crate) fn take_item_truncated(&mut self) -> bool {
        let cut = self.item_truncated
            || self.budget.exhausted()
            || self.padding_chars.exhausted()
            || self.total_image_budget.exhausted();
        self.item_truncated = false;
        cut
    }

    /// The warnings that describe the state of the run's budgets rather than
    /// one event: the same text however many sheets noticed it.
    pub(crate) fn end_warnings(&self) -> Vec<String> {
        let mut warnings = Vec::new();
        if self.budget.exhausted() {
            warnings.push(self.budget.notice());
        }
        match self.scope {
            IrScope::Document => {
                if self.padding_chars.exhausted() {
                    warnings.push(self.padding_chars.notice());
                }
                if self.image_budget.exhausted() {
                    warnings.push(format!(
                        "some pictures were left out: the document's pictures exceed the {} byte budget \
                         (office_oxide::limits::set_max_memory_bytes)",
                        self.image_budget.limit()
                    ));
                }
            },
            IrScope::Sheet => {
                if self.padding_chars.exhausted() {
                    warnings.push(self.padding_chars.notice());
                }
                if self.total_image_budget.exhausted() {
                    warnings.push(total_image_notice(&self.total_image_budget));
                }
            },
        }
        if self.format_code_omitted {
            warnings.push(format!(
                "some cells' number format code is longer than {} characters and is not listed on \
                 them (they keep their format id and values)",
                crate::limits::MAX_NUMBER_FORMAT_CHARS
            ));
        }
        warnings
    }

    /// Convert worksheet `ws_idx` to its section.
    pub(crate) fn convert_sheet(&mut self, ws_idx: usize) -> Section {
        let doc = self.doc;
        let ws = &doc.worksheets[ws_idx];
        if self.scope == IrScope::Sheet {
            // The previous sheet has been released, and with it what it held.
            self.padding_budget = Self::fresh_padding_budget(self.scope);
            self.image_budget =
                crate::limits::TextBudget::with_limit(crate::limits::max_image_bytes());
            self.image_notice_given = false;
        }
        self.sheet_padding_cut = None;
        let section = self.convert_sheet_inner(ws_idx, ws);
        let name = &ws.name;
        if let Some(by_memory) = self.sheet_padding_cut {
            let notice = if by_memory {
                self.padding_budget.notice()
            } else {
                self.padding_chars.notice()
            };
            self.warnings.push(format!("sheet {name:?}: {notice}"));
        }
        if self.scope == IrScope::Sheet {
            if self.image_budget.exhausted() {
                self.warnings.push(format!(
                    "sheet {name:?}: some pictures were left out: the sheet's pictures exceed the {} \
                     byte budget (office_oxide::limits::set_max_memory_bytes)",
                    self.image_budget.limit()
                ));
            }
        }
        self.padding_cut |= self.sheet_padding_cut.is_some();
        self.image_cut |= self.image_budget.exhausted();
        self.item_truncated |= self.sheet_padding_cut.is_some()
            || self.image_budget.exhausted()
            || ws.comments_truncated
            || ws.damage.is_some()
            || ws.formulas_dropped > 0;
        section
    }

    fn convert_sheet_inner(&mut self, ws_idx: usize, ws: &crate::xlsx::Worksheet) -> Section {
        let doc = self.doc;
        let Self {
            formatter,
            theme,
            buf,
            budget,
            image_budget,
            total_image_budget,
            padding_budget,
            padding_chars,
            comments_truncated,
            format_code_omitted,
            image_notice_given,
            total_image_notice_given,
            formulas_truncated,
            sheet_padding_cut,
            warnings,
            scope,
            ..
        } = self;
        // First pass: parse all rows into `CellData` — the rendered display
        // string plus the structured facts (semantic type, raw value, number
        // format) that the grid path threads into the IR so `to_ir()`
        // consumers can tell numbers/dates from text.
        // Cap eager row materialisation: an unbounded sheet would build
        // millions of IR cell allocations and then stall the single-table
        // render downstream. Excess rows are dropped and flagged below.
        // `Worksheet::hyperlinks` was parsed and then never rendered, so
        // every clickable cell in a workbook lost its target — DOCX and
        // PPTX both emit links. Index by (row, col) for O(1) lookup.
        // The value is the target and its hover text (`@tooltip`).
        // Targets and hover text are shared from here to the cells; the IR
        // owns a copy per span, which is charged to the budget below.
        type Link = (std::sync::Arc<str>, Option<std::sync::Arc<str>>);
        let links: std::collections::HashMap<(u32, u32), Link> = ws
            .hyperlinks
            .iter()
            .filter_map(|h| {
                let r = crate::xlsx::CellRef::parse(&h.cell_ref)?;
                let target = match &h.target {
                    crate::xlsx::HyperlinkTarget::External(u) => u.clone(),
                    crate::xlsx::HyperlinkTarget::Internal(loc) if !loc.is_empty() => {
                        std::sync::Arc::from(format!("#{loc}"))
                    },
                    crate::xlsx::HyperlinkTarget::Internal(_) => return None,
                };
                Some(((r.row, r.col), (target, h.tooltip.as_deref().map(std::sync::Arc::from))))
            })
            .collect();

        // `merged_cells` ("A1:C1"): the anchor carries the real span and
        // covered positions are excluded from the row entirely — the same
        // sparse, span-driven model every other format's TableRow uses
        // (matches ir_render.rs's table_grid, which resolves spans by
        // walking row.cells and skipping ahead over covered columns).
        // Each range is kept whole and consulted per stored row: expanding
        // it into the set of positions it covers cost its declared area,
        // and one legal `A1:XFD1048576` was ~17 billion inserts.
        let mut merges = RangeSweep::new(
            ws.merged_cells
                .iter()
                .filter_map(|range| {
                    let (start, end) = range.split_once(':')?;
                    let s = crate::xlsx::CellRef::parse(start)?;
                    let e = crate::xlsx::CellRef::parse(end)?;
                    let r = CellRange::from_corners(s.row, s.col, e.row, e.col);
                    (r.row_span() > 1 || r.col_span() > 1).then_some((r, ()))
                })
                .collect(),
        );

        let total_rows = ws.rows.len();
        let mut parsed_rows: Vec<Vec<CellData>> =
            Vec::with_capacity(total_rows.min(MAX_ROWS_PER_SHEET));
        // Absolute 0-based sheet row number per `parsed_rows` entry, kept
        // alongside rather than folded into `CellData` (only the merge
        // lookup below needs it). `merged_cells` ranges like "A5:C5" are
        // sheet-absolute, and `parsed_rows`'s own index is only sheet-row-
        // aligned when there's no gap of fully-empty rows — the same
        // assumption `grid_width`/`is_header` already make elsewhere in
        // this function, not a new limitation introduced here.
        let mut row_numbers: Vec<u32> = Vec::with_capacity(total_rows.min(MAX_ROWS_PER_SHEET));
        for row in ws.rows.iter().take(MAX_ROWS_PER_SHEET) {
            if budget.exhausted() {
                break;
            }
            row_numbers.push(row.index.saturating_sub(1));
            let mut cells: Vec<CellData> = Vec::with_capacity(row.cells.len());
            for cell in &row.cells {
                buf.clear();
                formatter.write(cell, buf);
                let link = links.get(&(cell.reference.row, cell.reference.col));
                // The cell's text, and the hyperlink target and hover text
                // each span of it will carry, are paid for before any is
                // made. A row keeps the cells that fit the budget; the sheet
                // ends after it.
                // A rich-text cell has one span per run, each with its own
                // copy of the link and its own formatting.
                let rich_runs = rich_runs_of(doc, cell);
                let runs = rich_runs.as_deref().filter(|r| !r.is_empty());
                let link_chars = link
                    .map_or(0, |(target, tip)| target.len() + tip.as_ref().map_or(0, |t| t.len()))
                    .saturating_mul(runs.map_or(1, <[_]>::len));
                let run_chars = runs.map_or(0, rich_span_chars);
                if !budget.charge(
                    buf.len()
                        .saturating_add(link_chars)
                        .saturating_add(run_chars),
                ) {
                    break;
                }
                let text = if buf.is_empty() {
                    String::new()
                } else {
                    std::mem::take(buf)
                };
                let (data_type, raw_number, number_format, number_format_id) =
                    cell_semantics(doc, cell, formatter.date_indices(), format_code_omitted);
                cells.push(CellData {
                    col: cell.reference.col,
                    hyperlink: link.map(|(target, _)| target.clone()),
                    hyperlink_tooltip: link.and_then(|(_, tip)| tip.clone()),
                    text,
                    style_index: cell.style_index,
                    data_type,
                    raw_number,
                    number_format,
                    number_format_id,
                    formula: cell.formula.clone(),
                    rich_runs,
                });
            }
            // Drop trailing empty cells.
            while cells
                .last()
                .is_some_and(|cd| cd.text.is_empty() && cd.formula.is_none())
            {
                cells.pop();
            }
            parsed_rows.push(cells);
        }

        // Widest column actually used, so every emitted row is the same
        // length and each value sits under its own header.
        // Widest index seen anywhere in the row, not the index of its last
        // cell: cells are not guaranteed to arrive in ascending column order,
        // and a short grid truncates every row laid out against it.
        let grid_width = parsed_rows
            .iter()
            .filter_map(|cells| cells.iter().map(|cd| cd.col as usize + 1).max())
            .max()
            .unwrap_or(0)
            // The readers already refuse a column past XFD; this is the
            // bound the allocation below relies on, kept next to it.
            .min(crate::xlsx::cell::MAX_COL as usize + 1);

        // Decide row layout: a "document style" sheet — flowing text laid
        // out one paragraph per row — becomes Paragraphs, not a 1-column
        // Table, so the downstream PDF renderer flows it like body text and
        // honours per-paragraph font sizes. The direct markdown renderer
        // makes the same decision through the same predicate.
        let prose_mode = crate::xlsx::text::rows_read_as_prose(parsed_rows.iter().map(|cells| {
            cells
                .iter()
                .filter(|cd| !cd.text.is_empty() || cd.formula.is_some())
                .map(|cd| cd.text.chars().count())
        }));

        // Materialise any pictures or text shapes anchored on the
        // worksheet as positional IR elements so they survive the
        // round-trip back to PDF. Pictures wrap an `Element::Image`
        // in an `Element::TextBox`; text shapes wrap a styled
        // paragraph the same way. The flow renderer then paints
        // both at their absolute EMU rectangle (see
        // `render_text_box`).
        let mut image_elements: Vec<Element> = Vec::new();
        for ts in ws.text_shapes.iter() {
            // A drawing is shared by every sheet that names it, but each
            // sheet's IR owns a copy of its text: paid for before it is made.
            if !budget.charge(text_shape_chars(ts)) {
                break;
            }
            let mut span = TextSpan::plain(ts.text.clone());
            if let Some(sz) = ts.font_size_pt {
                span.font_size_half_pt =
                    Some(crate::core::units::HalfPoint::from_points_rounded(sz as f64).0);
            }
            if ts.bold {
                span.bold = true;
            }
            if ts.italic {
                span.italic = true;
            }
            if let Some(ref hex) = ts.color_hex {
                if let Some(rgb) = parse_hex_rgb(hex) {
                    span.color = Some(rgb);
                }
            }
            if let Some(ref f) = ts.font_name {
                span.font_name = Some(f.clone());
            }
            let para = Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(span)],
                ..Default::default()
            });
            image_elements.push(Element::TextBox(TextBox {
                content: vec![para],
                x_emu: Some(ts.x_emu),
                y_emu: Some(ts.y_emu),
                width_emu: Some(ts.cx_emu.max(0) as u64),
                height_emu: Some(ts.cy_emu.max(0) as u64),
                ..Default::default()
            }));
        }
        for pic in ws.images.iter().chain(&ws.cell_images) {
            // A picture drawn from many anchors shares its bytes in the
            // workbook, but each IR image owns a copy, so every one is paid
            // for: the pictures past the budget are left out, and said so.
            // Its alternative text and the elements around it are paid for
            // as text.
            let total = (*scope == IrScope::Sheet).then_some(&mut *total_image_budget);
            if !charge_picture(budget, image_budget, total, picture_chars(pic), pic.data.len()) {
                break;
            }
            let format = image_format_from_ext(&pic.format);
            let img = Image {
                alt_text: pic.alt_text.clone(),
                data: Some(pic.data.to_vec()),
                format,
                display_width_emu: Some(pic.cx_emu.max(0) as u64),
                display_height_emu: Some(pic.cy_emu.max(0) as u64),
                ..Default::default()
            };
            // Wrap in TextBox so downstream renderers can paint at the
            // exact (x_emu, y_emu) anchor instead of inline-after-text.
            // When the source drawing was a cell-anchor and we
            // couldn't resolve EMU coords (cx == 0), drop the wrap so
            // the image flows inline at the section start.
            if pic.cx_emu > 0 && pic.cy_emu > 0 {
                image_elements.push(Element::TextBox(TextBox {
                    content: vec![Element::Image(img)],
                    x_emu: Some(pic.x_emu),
                    y_emu: Some(pic.y_emu),
                    width_emu: Some(pic.cx_emu.max(0) as u64),
                    height_emu: Some(pic.cy_emu.max(0) as u64),
                    ..Default::default()
                }));
            } else {
                image_elements.push(Element::Image(img));
            }
        }

        let elements = if prose_mode {
            // Each row → one Paragraph. Pull font size from cell style if the
            // worksheet's stylesheet is loaded. Skip empty rows entirely
            // (they were just visual separators).
            let mut out: Vec<Element> = Vec::new();
            for cells in &parsed_rows {
                // Every non-empty cell of the row (a formula with no cached
                // value counts too — it has real content, just no cached
                // display text), tab-separated as `plain_text()` lays the
                // row out. Prose mode allows one row in five to hold more
                // than one cell, and taking only the first cell of such a
                // row dropped the rest.
                let mut content: Vec<InlineContent> = Vec::new();
                for cd in cells.iter().filter(|cd| !cd.text.is_empty()) {
                    if !content.is_empty() {
                        content.push(InlineContent::Text(TextSpan::plain("\t")));
                    }
                    content.extend(cell_spans(doc, cd));
                }
                if content.is_empty() {
                    continue;
                }
                out.push(Element::Paragraph(Paragraph {
                    content,
                    ..Default::default()
                }));
            }
            out
        } else {
            // Genuine grid → emit a Table.
            let mut rows: Vec<TableRow> = Vec::with_capacity(parsed_rows.len());
            for (row_idx, cells) in parsed_rows.iter().enumerate() {
                // Lay the row out on the grid: a cell goes at its own column
                // index and skipped columns become empty cells, so `B2`
                // stays under the `B` header even when `A2` is absent.
                // Every row is as wide as the sheet, so the columns it does
                // not use become empty cells, each an allocation. They
                // are charged to the budget, which ends the table at the row
                // that spends it: a sheet with cells at A1 and XFD1 and a
                // thousand rows between would otherwise be 16 million cells.
                // The empty cells made are those of the grid's columns no cell
                // of the row occupies (cells may share a column or arrive out
                // of order). They are paid for as memory, and as characters
                // too, as Markdown pays for an absent column.
                let padding = grid_width.saturating_sub(distinct_columns(cells));
                let by_memory = !padding_budget.charge_cells(padding);
                if by_memory
                    || !padding_chars
                        .charge(padding.saturating_mul(crate::xlsx::text::MARKDOWN_PAD_CHARS))
                {
                    *sheet_padding_cut = Some(by_memory);
                    break;
                }
                let mut tcells: Vec<TableCell> = Vec::with_capacity(grid_width);
                for cd in cells {
                    let content = if cd.text.is_empty() {
                        Vec::new()
                    } else {
                        // Cell font formatting reached the IR in prose mode
                        // but not here, so the same cell rendered differently
                        // depending on the shape of the sheet around it.
                        cell_spans(doc, cd)
                    };
                    let cell = TableCell {
                        content: vec![Element::Paragraph(Paragraph {
                            content,
                            ..Default::default()
                        })],
                        col_span: 1,
                        row_span: 1,
                        data_type: cd.data_type,
                        raw_number: cd.raw_number,
                        number_format: cd.number_format.clone(),
                        number_format_id: cd.number_format_id,
                        formula: cd.formula.clone(),
                        background_color: cell_background(doc, cd, theme.as_deref()),
                        border: cell_border(doc, cd, theme.as_deref()),
                        text_align: cell_h_align(doc, cd),
                        vertical_align: cell_v_align(doc, cd),
                        wrap_text: cell_alignment(doc, cd).is_some_and(|a| a.wrap_text),
                        ..Default::default()
                    };
                    while tcells.len() < cd.col as usize {
                        tcells.push(empty_cell());
                    }
                    match tcells.get_mut(cd.col as usize) {
                        // The column is already occupied: cells arrived out of
                        // order, or two of them share a reference. Fill the
                        // slot if it is still blank, otherwise keep the first
                        // value — either way, never drop the rest of the row.
                        Some(slot) => {
                            if cell_is_blank(slot) {
                                *slot = cell;
                            }
                        },
                        None => tcells.push(cell),
                    }
                }
                while tcells.len() < grid_width {
                    tcells.push(empty_cell());
                }
                // Apply merges: set the anchor's real span, then drop
                // every position the merge covers from the row entirely
                // (dense-grid -> sparse, span-driven row) rather than
                // rebuilding the placement loop above around them.
                let true_row = row_numbers.get(row_idx).copied().unwrap_or(row_idx as u32);
                let tcells = if merges.is_empty() {
                    tcells
                } else {
                    apply_merges(&mut merges, tcells, true_row, &row_numbers[row_idx..])
                };
                rows.push(TableRow {
                    cells: tcells,
                    is_header: row_idx == 0,
                    ..Default::default()
                });
            }
            if rows.is_empty() {
                Vec::new()
            } else {
                vec![Element::Table(Table {
                    rows,
                    ..Default::default()
                })]
            }
        };

        // Per-sheet page geometry (parsed from <pageMargins>/<pageSetup>).
        // Default the margins back to 0.5"/0.5" (720 twips) when the source
        // had no <pageMargins> — Excel's default 0.7"/0.75" is wider than
        // we want for a tight PDF round-trip and would shrink the usable
        // text area.
        let page_setup = ws.page_setup.map(|wsp| {
            // When <pageMargins> was present but <pageSetup> was not,
            // wsp's width/height come through as 0. Fall back to the
            // IR PageSetup default geometry rather than dropping the
            // parsed margins on the floor.
            let default = PageSetup::default();
            PageSetup {
                width_twips: if wsp.width_twips == 0 {
                    default.width_twips
                } else {
                    wsp.width_twips
                },
                height_twips: if wsp.height_twips == 0 {
                    default.height_twips
                } else {
                    wsp.height_twips
                },
                margin_top_twips: wsp.margin_top_twips,
                margin_bottom_twips: wsp.margin_bottom_twips,
                margin_left_twips: wsp.margin_left_twips,
                margin_right_twips: wsp.margin_right_twips,
                header_distance_twips: wsp.header_distance_twips,
                footer_distance_twips: wsp.footer_distance_twips,
                landscape: wsp.landscape,
                ..default
            }
        });

        // Each XLSX worksheet renders to its own PDF page sequence, so
        // mark every section after the first as a hard page break (same
        // pattern as PPTX in convert_pptx.rs). Without this the second
        // worksheet's content flows into the first sheet's last page.
        let break_type = if ws_idx == 0 {
            SectionBreakType::Continuous
        } else {
            SectionBreakType::NextPage
        };

        // Stitch worksheet pictures in front of cell-derived content
        // so they paint underneath the text (positional TextBoxes are
        // absolute regardless of order, but inline images render
        // first). Empty `image_elements` means no drawing on this sheet.
        // A chart sheet's content is its chart's text; it has no cells.
        // The chart is shared by every chartsheet that displays it, and each
        // sheet's IR owns a copy of its text: paid for before it is made.
        let mut combined: Vec<Element> = Vec::new();
        for t in &ws.chart_text {
            if !budget.charge(
                std::mem::size_of::<Paragraph>()
                    .saturating_add(std::mem::size_of::<TextSpan>())
                    .saturating_add(t.trim().len()),
            ) {
                break;
            }
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(t.trim()))],
                ..Default::default()
            }));
        }
        combined.extend(image_elements);
        combined.extend(elements);

        // Cell comments are document content — review notes, provenance,
        // caveats on a figure — and were never surfaced anywhere. Append
        // them as endnotes so every renderer sees them, with the author and
        // cell in the marker.
        for (i, c) in ws.comments.iter().enumerate() {
            // The author is named in the marker and again as the note's
            // author, for every comment: charged like any other text.
            if !budget.charge(
                crate::xlsx::text::comment_chars(c)
                    .saturating_add(c.author.as_ref().map_or(0, |a| a.len())),
            ) {
                break;
            }
            let marker = match c.author.as_deref() {
                Some(a) => format!("{} ({a})", c.cell_ref),
                None => c.cell_ref.clone(),
            };
            combined.push(Element::Endnote(Note {
                id: i as u32,
                marker: Some(marker),
                author: c.author.as_deref().map(str::to_string),
                content: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(c.text.clone()))],
                    ..Default::default()
                })],
            }));
        }

        if ws.comments_truncated {
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(
                    crate::xlsx::text::COMMENTS_TRUNCATED_NOTICE,
                ))],
                ..Default::default()
            }));
            *comments_truncated = true;
        }
        // Graceful truncation notice: when the sheet exceeded the row cap,
        // the dropped rows are not silently lost — a visible paragraph records
        // how many rows were omitted so downstream text/markdown/PDF makes the
        // truncation obvious rather than appearing complete.
        if total_rows > MAX_ROWS_PER_SHEET {
            let omitted = total_rows - MAX_ROWS_PER_SHEET;
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "[{omitted} of {total_rows} rows not shown — worksheet truncated at {MAX_ROWS_PER_SHEET} rows]"
                )))],
                ..Default::default()
            }));
        }
        // Only a file declaring vast numbers of overlapping merge ranges
        // spends the sweep's work budget; the ranges past it are not
        // applied, and that is stated.
        if merges.exhausted() {
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(
                    "[some merged cell ranges not applied — too many overlapping ranges]",
                ))],
                ..Default::default()
            }));
        }
        if image_budget.exhausted() && !*image_notice_given {
            *image_notice_given = true;
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "[some pictures not shown: together they exceed the {} byte picture budget \
                     (office_oxide::limits::set_max_memory_bytes)]",
                    image_budget.limit()
                )))],
                ..Default::default()
            }));
        }
        if total_image_budget.exhausted() && !*total_image_notice_given {
            *total_image_notice_given = true;
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(total_image_notice(
                    total_image_budget,
                )))],
                ..Default::default()
            }));
        }
        if let Some(damage) = &ws.damage {
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(
                    crate::xlsx::text::damaged_notice(&ws.name, damage),
                ))],
                ..Default::default()
            }));
            warnings.push(format!("sheet {:?} cut short: {damage}", ws.name));
            // The sheet lost rows: the document's text is truncated.
            *formulas_truncated = true;
        }
        if ws.formulas_dropped > 0 {
            *formulas_truncated = true;
            warnings.push(format!(
                "sheet {:?}: {} shared formula(s) left out: a master longer than {} characters, or \
                 copies past the character budget (office_oxide::limits::set_max_text_chars); \
                 the cells keep their values",
                ws.name,
                ws.formulas_dropped,
                crate::limits::MAX_FORMULA_CHARS
            ));
        }
        if budget.exhausted() {
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(budget.notice()))],
                ..Default::default()
            }));
        }
        if let Some(by_memory) = *sheet_padding_cut {
            combined.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(if by_memory {
                    padding_budget.notice()
                } else {
                    padding_chars.notice()
                }))],
                ..Default::default()
            }));
        }

        // Page headers/footers in use ([ECMA-376] §18.3.1.46).
        let hf = |text: Option<String>| {
            text.map(|t| HeaderFooter {
                content: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(t))],
                    ..Default::default()
                })],
            })
        };
        let [fh, oh, eh, ff, of, ef] = ws.header_footer.active(&ws.name);
        Section {
            title: Some(ws.name.clone()),
            elements: combined,
            page_setup,
            break_type,
            header: hf(oh),
            footer: hf(of),
            first_page_header: hf(fh),
            first_page_footer: hf(ff),
            even_page_header: hf(eh),
            even_page_footer: hf(ef),
            hidden: ws.state != crate::xlsx::SheetState::Visible,
            sheet_kind: Some(ws.kind),
            conditional_formats: ws.conditional_formats.clone(),
            data_validations: ws.data_validations.clone(),
            ..Default::default()
        }
    }

    /// The section for the text of the workbook's charts, when there is any.
    /// We don't render charts as graphics; capturing their text (titles, axis
    /// labels, series names, cached values) ensures that all
    /// human-meaningful words in the workbook appear in the IR and downstream
    /// conversions, even when the chart itself isn't drawn.
    pub(crate) fn chart_section(&self) -> Option<Section> {
        let doc = self.doc;
        if doc.chart_text.is_empty() {
            return None;
        }
        let mut chart_elements: Vec<Element> = Vec::new();
        for (i, text) in doc.chart_text.iter().enumerate() {
            chart_elements.push(Element::Heading(Heading {
                level: 3,
                content: vec![InlineContent::Text(TextSpan::plain(format!(
                    "Chart {}",
                    i + 1
                )))],
                ..Default::default()
            }));
            chart_elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(text.clone()))],
                ..Default::default()
            }));
        }
        Some(Section {
            title: Some("Charts".to_string()),
            elements: chart_elements,
            ..Default::default()
        })
    }

    /// The section for pivot caches whose source is not in the workbook: the
    /// cached records are the only copy of the data behind the pivot. Like
    /// chart text, one heading and table per pivot.
    pub(crate) fn pivot_section(&mut self) -> Option<Section> {
        let doc = self.doc;
        let budget = &mut self.budget;
        let mut pivot_elements: Vec<Element> = Vec::new();
        let mut pivot_truncated = false;
        for (name, data) in doc.pivot_caches() {
            // Pivots that name one cache each render it, so the table is
            // charged to the budget however many of them share the data.
            if budget.exhausted() {
                break;
            }
            // The heading and the header row are paid for before they are built
            // (an empty field name still costs its cell); a pivot whose header
            // does not fit is left out, and so is every later one.
            let title = crate::xlsx::text::pivot_cache_title(name);
            let header_cost = title.len()
                + data
                    .fields
                    .iter()
                    .map(|f| f.len().max(1))
                    .sum::<usize>()
                    .max(1);
            if !budget.charge(header_cost) {
                break;
            }
            pivot_elements.push(Element::Heading(Heading {
                level: 3,
                content: vec![InlineContent::Text(TextSpan::plain(title))],
                ..Default::default()
            }));
            let text_cell = |t: &str| TableCell {
                content: vec![Element::Paragraph(Paragraph {
                    content: if t.is_empty() {
                        Vec::new()
                    } else {
                        vec![InlineContent::Text(TextSpan::plain(t))]
                    },
                    ..Default::default()
                })],
                col_span: 1,
                row_span: 1,
                ..Default::default()
            };
            let cols = data.fields.len();
            let mut rows = vec![TableRow {
                cells: data.fields.iter().map(|f| text_cell(f)).collect(),
                is_header: true,
                ..Default::default()
            }];
            for rec in &data.records {
                if !budget.charge(rec.iter().map(String::len).sum::<usize>() + cols.max(rec.len()))
                {
                    break;
                }
                rows.push(TableRow {
                    cells: (0..cols.max(rec.len()))
                        .map(|i| text_cell(rec.get(i).map(String::as_str).unwrap_or("")))
                        .collect(),
                    ..Default::default()
                });
            }
            pivot_elements.push(Element::Table(Table {
                rows,
                ..Default::default()
            }));
            if data.truncated {
                pivot_truncated = true;
                pivot_elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(
                        crate::xlsx::text::PIVOT_CACHE_TRUNCATED_NOTICE,
                    ))],
                    ..Default::default()
                }));
            }
        }
        if budget.exhausted() && !pivot_elements.is_empty() {
            pivot_elements.push(Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan::plain(budget.notice()))],
                ..Default::default()
            }));
        }
        if pivot_truncated {
            self.pivot_truncated = true;
            self.item_truncated = true;
        }
        (!pivot_elements.is_empty()).then(|| Section {
            title: Some("Pivot Caches".to_string()),
            elements: pivot_elements,
            ..Default::default()
        })
    }

    /// How many sections report what could not be read (see
    /// [`Self::unreadable_section`]).
    pub(crate) fn unreadable_count(&mut self) -> usize {
        self.reports
            .get_or_insert_with(|| self.doc.report_notices())
            .len()
    }

    /// The section for the report `index` of what could not be read (an
    /// unreadable sheet, the sheets and parts left unlisted, the unreadable
    /// parts), and the warnings it carries: the notice, so the loss is
    /// visible in every projection of the IR. The notice is charged to the
    /// text budget before it is copied; one the budget cannot pay for is
    /// left out, with the budget's notice in its place, and the metadata
    /// still lists it.
    pub(crate) fn unreadable_section(
        &mut self,
        index: usize,
    ) -> Option<(Option<Section>, Vec<String>)> {
        let notices = self
            .reports
            .get_or_insert_with(|| self.doc.report_notices());
        let notice = notices.get(index)?;
        self.unreadable = true;
        self.item_truncated = true;
        let was_spent = self.budget.exhausted();
        if !self.budget.charge(notice.chars()) {
            // The budget's notice is said once, in the content, unless
            // something already said it. The warnings (bounded: see
            // `report_notices`) are kept either way, so that a caller who
            // drops each sheet still has what could not be read.
            let section = (!was_spent).then(|| Section {
                title: Some("Output truncated".to_string()),
                elements: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(self.budget.notice()))],
                    ..Default::default()
                })],
                ..Default::default()
            });
            return Some((section, notice.warnings.clone()));
        }
        let section = Section {
            title: Some(notice.title.clone()),
            elements: notice
                .lines
                .iter()
                .map(|line| {
                    Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan::plain(line.clone()))],
                        ..Default::default()
                    })
                })
                .collect(),
            ..Default::default()
        };
        Some((Some(section), notice.warnings.clone()))
    }

    /// The document's metadata as of the sections built so far: the
    /// properties, the defined names and every warning and truncation flag
    /// the conversion has raised. `first_title` is the title of the first
    /// section, the workbook's title when its core properties carry none.
    pub(crate) fn metadata(&self, first_title: Option<String>) -> Metadata {
        let doc = self.doc;
        // The first sheet's *name* is not the workbook's title. Read the real
        // one from `docProps/core.xml` and fall back to the sheet name only
        // when the package carries no core properties.
        let cp = doc.core_properties.as_ref();
        let title = cp
            .and_then(|c| c.title.clone())
            .filter(|t| !t.is_empty())
            .or(first_title);
        let mut metadata = Metadata {
            format: DocumentFormat::Xlsx,
            title,
            author: cp.and_then(|c| c.creator.clone()),
            subject: cp.and_then(|c| c.subject.clone()),
            keywords: cp
                .and_then(|c| c.keywords.as_deref())
                .map(crate::convert_docx::split_keywords)
                .unwrap_or_default(),
            created: cp.and_then(|c| c.created.clone()),
            modified: cp.and_then(|c| c.modified.clone()),
            description: cp.and_then(|c| c.description.clone()),
            has_macros: doc.has_macros,
            text_truncated: self.text_truncated(),
            defined_names: doc
                .workbook
                .defined_names
                .iter()
                .map(|dn| DefinedName {
                    name: dn.name.clone(),
                    value: dn.value.clone(),
                    local_sheet_id: dn.local_sheet_id,
                    hidden: dn.hidden,
                })
                .collect(),
            ..crate::core::core_properties::ooxml_metadata_extras(
                cp,
                doc.app_properties.as_ref(),
                Some(&doc.package_properties),
            )
        };
        // Sheets and parts that could not be read: the sheet, drawing or
        // pivot they belong to lacks what they name.
        let notices = doc.report_notices();
        metadata.text_truncated |= !notices.is_empty();
        metadata
            .warnings
            .extend(notices.into_iter().flat_map(|n| n.warnings));
        metadata.warnings.extend(
            self.doc
                .cosmetic_parts
                .iter()
                .map(|(part, err)| format!("skipped unreadable part {part}: {err}")),
        );
        metadata.warnings.extend(self.warnings.iter().cloned());
        metadata.warnings.extend(self.end_warnings());
        metadata
    }
}

/// What the IR's text box for a drawn text shape costs: the shape's text and
/// font name, which the IR copies, and the elements that hold them.
fn text_shape_chars(ts: &crate::xlsx::worksheet::WorksheetTextShape) -> usize {
    std::mem::size_of::<TextBox>()
        .saturating_add(std::mem::size_of::<Paragraph>())
        .saturating_add(std::mem::size_of::<TextSpan>())
        .saturating_add(ts.text.len())
        .saturating_add(ts.font_name.as_ref().map_or(0, String::len))
}

/// What the IR's image for a picture costs besides its bytes (which the
/// picture budgets count): its alternative text and the elements that hold it.
fn picture_chars(pic: &crate::xlsx::worksheet::WorksheetPicture) -> usize {
    std::mem::size_of::<TextBox>()
        .saturating_add(std::mem::size_of::<Image>())
        .saturating_add(pic.alt_text.as_ref().map_or(0, String::len))
}

/// Pay for a picture from every budget that bounds it, or from none: a
/// picture that is not produced costs nothing, whichever budget refuses it.
/// The budget (or budgets) that cannot pay is charged all the same, which is
/// what marks it spent for the notice that says so.
fn charge_picture(
    text: &mut crate::limits::TextBudget,
    image: &mut crate::limits::TextBudget,
    total: Option<&mut crate::limits::TextBudget>,
    chars: usize,
    bytes: usize,
) -> bool {
    let fits =
        text.fits(chars) && image.fits(bytes) && total.as_deref().is_none_or(|t| t.fits(bytes));
    if fits {
        // Each fits, so each charge succeeds.
        let _ = text.charge(chars);
        let _ = image.charge(bytes);
        if let Some(t) = total {
            let _ = t.charge(bytes);
        }
        return true;
    }
    // Only the budgets that cannot pay are touched.
    if !text.fits(chars) {
        let _ = text.charge(chars);
    }
    if !image.fits(bytes) {
        let _ = image.charge(bytes);
    }
    if let Some(t) = total {
        if !t.fits(bytes) {
            let _ = t.charge(bytes);
        }
    }
    false
}

/// The notice for picture bytes copied past the total the whole run may copy.
fn total_image_notice(budget: &crate::limits::TextBudget) -> String {
    format!(
        "[some pictures not shown: the workbook's pictures together exceed the {} byte total \
         (office_oxide::limits::set_max_memory_bytes)]",
        budget.limit()
    )
}

/// One section of a workbook's IR, converted on its own: a worksheet, or one
/// of the sections that follow the worksheets (chart text, pivot caches, a
/// sheet that could not be read).
///
/// Everything a whole-document [`DocumentIR`] reports in its metadata about
/// this section is here, so a caller that drops each section after use still
/// learns what was cut.
#[derive(Debug, Clone)]
pub struct SheetIr {
    /// The converted section.
    pub section: Section,
    /// What this section's conversion left out or cut, and why: the
    /// entries `Metadata::warnings` would carry for it.
    pub warnings: Vec<String>,
    /// Whether this section was cut (its own rows, pictures, comments or
    /// formulas, or the text budget of the whole run).
    pub text_truncated: bool,
}

enum Stage {
    Sheets(usize),
    Omitted(usize),
    Charts,
    Pivots,
    Unreadable(usize),
    Done,
}

/// The sections of a workbook's IR, converted one at a time as the iterator
/// is advanced (see [`XlsxDocument::sheets_ir`](crate::xlsx::XlsxDocument::sheets_ir)).
///
/// Only the section being handed out exists: the conversion holds no earlier
/// one, so a caller that drops each before asking for the next holds about
/// one sheet's IR however many sheets the workbook has.
pub struct SheetIrs<'d> {
    converter: SheetConverter<'d>,
    stage: Stage,
    first_title: Option<String>,
    warnings_reported: usize,
}

impl<'d> SheetIrs<'d> {
    pub(crate) fn new(doc: &'d crate::xlsx::XlsxDocument) -> Self {
        Self {
            converter: SheetConverter::new(doc, IrScope::Sheet),
            stage: Stage::Sheets(0),
            first_title: None,
            warnings_reported: 0,
        }
    }

    /// The workbook's metadata as of the sections handed out so far: its
    /// properties and defined names, every warning raised and whether
    /// anything was cut, as `to_ir().metadata` reports them once the whole
    /// workbook has been converted.
    pub fn metadata(&self) -> Metadata {
        self.converter.metadata(self.first_title.clone())
    }

    fn item(&mut self, section: Section, end_before: Vec<String>) -> SheetIr {
        self.item_with(section, end_before, Vec::new())
    }

    fn item_with(
        &mut self,
        section: Section,
        end_before: Vec<String>,
        extra: Vec<String>,
    ) -> SheetIr {
        if self.first_title.is_none() {
            self.first_title = section.title.clone();
        }
        let mut warnings = self.converter.warnings()[self.warnings_reported..].to_vec();
        warnings.extend(extra);
        self.warnings_reported = self.converter.warnings().len();
        for w in self.converter.end_warnings() {
            if !end_before.contains(&w) {
                warnings.push(w);
            }
        }
        SheetIr {
            section,
            warnings,
            text_truncated: self.converter.take_item_truncated(),
        }
    }
}

impl Iterator for SheetIrs<'_> {
    type Item = SheetIr;

    fn next(&mut self) -> Option<SheetIr> {
        loop {
            let end_before = self.converter.end_warnings();
            match self.stage {
                Stage::Sheets(i) => {
                    if i >= self.converter.doc.worksheets.len() {
                        self.stage = Stage::Charts;
                        continue;
                    }
                    if self.converter.should_stop() {
                        self.stage = Stage::Omitted(i);
                        continue;
                    }
                    self.stage = Stage::Sheets(i + 1);
                    let section = self.converter.convert_sheet(i);
                    return Some(self.item(section, end_before));
                },
                Stage::Omitted(from) => {
                    self.stage = Stage::Charts;
                    if let Some(section) = self.converter.omitted_section(from) {
                        return Some(self.item(section, end_before));
                    }
                },
                Stage::Charts => {
                    self.stage = Stage::Pivots;
                    if let Some(section) = self.converter.chart_section() {
                        return Some(self.item(section, end_before));
                    }
                },
                Stage::Pivots => {
                    self.stage = Stage::Unreadable(0);
                    if let Some(section) = self.converter.pivot_section() {
                        return Some(self.item(section, end_before));
                    }
                },
                Stage::Unreadable(i) => {
                    self.stage = Stage::Unreadable(i + 1);
                    if i >= self.converter.unreadable_count() {
                        self.stage = Stage::Done;
                        continue;
                    }
                    if let Some((section, extra)) = self.converter.unreadable_section(i) {
                        // A notice the budget cannot pay and has already said
                        // so leaves no content, only its warnings.
                        return Some(self.item_with(
                            section.unwrap_or_default(),
                            end_before,
                            extra,
                        ));
                    }
                },
                Stage::Done => return None,
            }
        }
    }
}

/// Convert worksheet `index` alone, with the allowances of one sheet.
pub(crate) fn sheet_to_ir(doc: &crate::xlsx::XlsxDocument, index: usize) -> Option<SheetIr> {
    if index >= doc.worksheets.len() {
        return None;
    }
    let mut converter = SheetConverter::new(doc, IrScope::Sheet);
    let end_before = converter.end_warnings();
    let section = converter.convert_sheet(index);
    let mut warnings = converter.warnings().to_vec();
    warnings.extend(
        converter
            .end_warnings()
            .into_iter()
            .filter(|w| !end_before.contains(w)),
    );
    Some(SheetIr {
        section,
        warnings,
        text_truncated: converter.take_item_truncated(),
    })
}

/// Convert every worksheet, the chart text, the pivot caches and the sheets
/// that could not be read to one IR, held whole.
pub(crate) fn xlsx_to_ir(doc: &crate::xlsx::XlsxDocument) -> DocumentIR {
    let mut converter = SheetConverter::new(doc, IrScope::Document);
    let mut sections = Vec::new();
    for ws_idx in 0..doc.worksheets.len() {
        if converter.should_stop() {
            sections.extend(converter.omitted_section(ws_idx));
            break;
        }
        sections.push(converter.convert_sheet(ws_idx));
    }
    sections.extend(converter.chart_section());
    sections.extend(converter.pivot_section());
    for index in 0..converter.unreadable_count() {
        sections.extend(
            converter
                .unreadable_section(index)
                .and_then(|(section, _)| section),
        );
    }
    let metadata = converter.metadata(sections.first().and_then(|s| s.title.clone()));
    DocumentIR { metadata, sections }
}

/// Give each merge anchor in `tcells` (one per grid column of sheet row
/// `true_row`) its span and drop every covered position. `rows_from_here`
/// is the absolute row number of this and every later stored row: the IR
/// table holds only the rows the sheet stores, so a merge over sheet rows
/// 1-5 in a sheet whose next stored row is 6 spans one IR row, not five.
fn apply_merges(
    merges: &mut RangeSweep<()>,
    mut tcells: Vec<TableCell>,
    true_row: u32,
    rows_from_here: &[u32],
) -> Vec<TableCell> {
    let width = tcells.len();
    let mut covered = vec![false; width];
    let mut cost = 0u64;
    for (range, ()) in merges.row(true_row) {
        let lo = range.col_lo as usize;
        if lo >= width {
            continue;
        }
        let hi = (range.col_hi as usize).min(width - 1);
        if range.row_lo == true_row {
            let last = range.row_hi;
            let rows = rows_from_here
                .iter()
                .take_while(|&&r| r <= last)
                .count()
                .max(1);
            cost += rows as u64;
            let anchor = &mut tcells[lo];
            anchor.row_span = rows as u32;
            // Clipped to the table's width, like the row span to its rows.
            anchor.col_span = range.col_span().min((width - lo) as u64) as u32;
            covered[lo + 1..=hi].fill(true);
        } else {
            covered[lo..=hi].fill(true);
        }
        cost += (hi - lo + 1) as u64;
    }
    merges.charge(cost);
    if !covered.contains(&true) {
        return tcells;
    }
    tcells
        .into_iter()
        .zip(covered)
        .filter(|(_, c)| !c)
        .map(|(cell, _)| cell)
        .collect()
}

/// A grid position with no cell in the source.
/// True when a laid-out cell still holds nothing, so a later cell claiming
/// the same column may take the slot rather than be discarded.
fn cell_is_blank(cell: &TableCell) -> bool {
    cell.content.iter().all(|el| match el {
        Element::Paragraph(p) => p.content.is_empty(),
        _ => false,
    })
}

/// How many different columns the cells of a row occupy.
fn distinct_columns(cells: &[CellData]) -> usize {
    if cells.windows(2).all(|w| w[0].col < w[1].col) {
        return cells.len();
    }
    let mut cols: Vec<u32> = cells.iter().map(|cd| cd.col).collect();
    cols.sort_unstable();
    cols.dedup();
    cols.len()
}

fn empty_cell() -> TableCell {
    TableCell {
        content: vec![Element::Paragraph(Paragraph::default())],
        col_span: 1,
        row_span: 1,
        ..Default::default()
    }
}

/// The rich-text runs `cell` carries: those of its shared string, or its own
/// inline ones. Shared, not copied: a shared string names the same runs from
/// every cell that uses it.
fn rich_runs_of(
    doc: &crate::xlsx::XlsxDocument,
    cell: &crate::xlsx::Cell,
) -> Option<std::sync::Arc<[crate::xlsx::shared_strings::RichTextRun]>> {
    match &cell.value {
        crate::xlsx::cell::CellValue::SharedString(idx) => doc
            .shared_strings
            .get_shared(*idx)
            .and_then(|s| s.rich_text.clone()),
        // An inline (`t="inlineStr"`) string cell carries its own rich runs
        // directly on the raw `Cell`, not via the shared string table.
        crate::xlsx::cell::CellValue::String(_) => cell.rich_runs.clone(),
        _ => None,
    }
}

/// What the IR's spans for `runs` cost: each run becomes a `TextSpan` of its
/// own, with its own copy of the run's text (the displayed cell text stops at
/// 32 KB, the runs do not) and of the font name. The runs are
/// stored once, so without this a string of thousands of runs named by
/// thousands of cells is a span per run per cell for free.
fn rich_span_chars(runs: &[crate::xlsx::shared_strings::RichTextRun]) -> usize {
    runs.iter().fold(0usize, |total, r| {
        total
            .saturating_add(std::mem::size_of::<TextSpan>())
            .saturating_add(r.text.len())
            .saturating_add(r.font_name.as_ref().map_or(0, String::len))
    })
}

/// Build the styled span for a cell: display text, the cell font's size and
/// weight from the workbook stylesheet, and the sheet's hyperlink target if
/// the cell has one.
fn cell_span(doc: &crate::xlsx::XlsxDocument, cd: &CellData) -> TextSpan {
    let mut span = TextSpan::plain(display_text(cd).to_string());
    span.hyperlink = cd.hyperlink.as_deref().map(str::to_string);
    span.hyperlink_tooltip = cd.hyperlink_tooltip.as_deref().map(str::to_string);
    let Some(font) = cd.style_index.and_then(|idx| font_for(doc, idx)) else {
        return span;
    };
    if let Some(size_pt) = font.size {
        // XLSX cell font size is in points (`<font><sz val="N"/>` where N is
        // f32). IR uses half-points; same convention as the DOCX/PPTX paths.
        span.font_size_half_pt =
            Some(crate::core::units::HalfPoint::from_points_rounded(size_pt).0);
    }
    span.bold = font.bold;
    span.italic = font.italic;
    span
}

/// Build a cell's paragraph content: one `TextSpan` per rich-text run
/// when the cell's shared string carries per-run formatting (e.g. a
/// bold superscript footnote marker within otherwise-plain text) —
/// previously discarded entirely, flattening to a single unformatted
/// span regardless of how many differently-formatted runs the source
/// actually had. Falls back to the single-span `cell_span`
/// path for a plain string or any non-string cell.
fn cell_spans(doc: &crate::xlsx::XlsxDocument, cd: &CellData) -> Vec<InlineContent> {
    let Some(runs) = cd.rich_runs.as_ref().filter(|r| !r.is_empty()) else {
        return vec![InlineContent::Text(cell_span(doc, cd))];
    };
    runs.iter()
        .map(|r| {
            let mut span = TextSpan::plain(r.text.clone());
            span.hyperlink = cd.hyperlink.as_deref().map(str::to_string);
            span.hyperlink_tooltip = cd.hyperlink_tooltip.as_deref().map(str::to_string);
            span.bold = r.bold.unwrap_or(false);
            span.italic = r.italic.unwrap_or(false);
            if let Some(size_pt) = r.font_size {
                span.font_size_half_pt =
                    Some(crate::core::units::HalfPoint::from_points_rounded(size_pt).0);
            }
            span.font_name = r.font_name.clone();
            span.color = r
                .color
                .as_ref()
                .and_then(|c| c.resolve_opt(doc.theme.as_ref()))
                .map(|rgb| rgb.0);
            span.vertical_align = r.vert_align.clone();
            InlineContent::Text(span)
        })
        .collect()
}

/// A parsed spreadsheet cell carried through `xlsx_to_ir`: the rendered
/// display string plus the structured facts needed to populate the IR's
/// semantic `TableCell` fields.
struct CellData {
    /// 0-based grid column this cell occupies. Without it, cells were
    /// emitted in encounter order, so a row that skips a column (perfectly
    /// legal — XLSX stores only non-empty cells) shifted every later value
    /// one column left and filed it under the wrong header.
    col: u32,
    /// Hyperlink target for this cell, when the sheet declares one.
    hyperlink: Option<std::sync::Arc<str>>,
    /// That hyperlink's hover text (`hyperlink/@tooltip`).
    hyperlink_tooltip: Option<std::sync::Arc<str>>,
    /// Rendered display string (same text `write_cell_value_fast` produces).
    text: String,
    /// Cell format index (`s` attribute) — used for prose-mode font recovery.
    style_index: Option<u32>,
    /// Semantic type, or `None` for empty cells.
    data_type: Option<CellDataType>,
    /// Underlying numeric value for number/date/boolean cells.
    raw_number: Option<f64>,
    /// Number-format code, when the cell has a non-General custom format.
    number_format: Option<String>,
    /// Number-format ID, when the cell has a non-General format.
    number_format_id: Option<u32>,
    /// Formula text (`<f>` content, shared-formula followers already
    /// reconstructed by `xlsx::shared_formula`), when present.
    formula: Option<String>,
    /// Per-run rich-text formatting, when this cell's value is a shared
    /// string with `<r><rPr>…</rPr><t>…</t></r>` sub-runs (e.g. a bold
    /// superscript footnote marker within otherwise-plain text). `None`
    /// for a plain (non-rich) string or any non-string cell.
    rich_runs: Option<std::sync::Arc<[crate::xlsx::shared_strings::RichTextRun]>>,
}

/// The text a cell should show when it has no cached value: `=formula`
/// when one exists, otherwise empty. A formula cell with no `<v>` (common
/// output shape from closedxml and similar writers that never cache
/// values) used to render as a blank cell indistinguishable from a
/// genuinely empty one.
/// What the cell shows: the same text `plain_text()` renders (a formula
/// with no cached value is already `=formula` there). A formula whose
/// cached value is the empty string shows nothing, as in Excel, Tika and
/// calamine; rendering `=formula` for it here made `to_html()` of a
/// lookup-heavy workbook 10,000 words longer than its own `plain_text()`.
fn display_text(cd: &CellData) -> &str {
    &cd.text
}

/// Derive a cell's semantic type, raw numeric value, and number-format
/// metadata. Mirrors the type/date decisions in `write_cell_value_fast` but
/// preserves the structured facts instead of collapsing them to a string, so
/// `to_ir()` consumers (e.g. the WASM `toIr()` surface) can distinguish a
/// number or date cell from a text cell.
///
/// Format metadata is surfaced for every non-empty cell — not just numbers —
/// so a caller can tell, for instance, that a *text* cell sits in a
/// date-formatted column (the "is it a date/number column, and
/// formatting?"). The format string prefers the workbook's custom `<numFmts>`
/// entry and falls back to the canonical code for built-in IDs, which never
/// appear in `<numFmts>`.
fn cell_semantics(
    doc: &crate::xlsx::XlsxDocument,
    cell: &crate::xlsx::Cell,
    date_indices: &std::collections::HashSet<u32>,
    format_code_omitted: &mut bool,
) -> (Option<CellDataType>, Option<f64>, Option<String>, Option<u32>) {
    use crate::xlsx::CellValue;

    if matches!(cell.value, CellValue::Empty) {
        return (None, None, None, None);
    }

    // Number-format metadata (skip General / id 0 — carries no information).
    let fmt_id = cell
        .style_index
        .and_then(|idx| doc.styles.as_ref()?.number_format_id_for(idx))
        .filter(|&id| id != 0);
    let declared = cell
        .style_index
        .and_then(|idx| doc.styles.as_ref()?.number_format_for(idx));
    let max = crate::limits::MAX_NUMBER_FORMAT_CHARS;
    // The code is copied onto every cell of the style, so one longer than
    // Excel allows is left off rather than multiplied (and is never copied
    // to be measured). A character is at most four bytes, so only a short
    // code is counted.
    let fmt_str = match declared {
        Some(code) if code.len() > max * 4 || code.chars().count() > max => {
            *format_code_omitted = true;
            None
        },
        Some(code) => Some(code.to_string()),
        None => {
            fmt_id.and_then(|id| crate::xlsx::numfmt::builtin_format_code(id).map(str::to_string))
        },
    };

    let (data_type, raw_number) = match &cell.value {
        CellValue::Empty => unreachable!("handled above"),
        CellValue::Number(n) => {
            let is_date = cell.style_index.is_some_and(|i| date_indices.contains(&i));
            let ty = if is_date {
                CellDataType::Date
            } else {
                CellDataType::Number
            };
            (Some(ty), Some(*n))
        },
        // Dates almost always arrive as Number + a date format (handled above);
        // this variant is a defensive fallback and carries no serial to expose.
        CellValue::Date(_) => (Some(CellDataType::Date), None),
        CellValue::Boolean(b) => (Some(CellDataType::Boolean), Some(if *b { 1.0 } else { 0.0 })),
        CellValue::Error(_) => (Some(CellDataType::Error), None),
        CellValue::String(_) | CellValue::SharedString(_) => (Some(CellDataType::Text), None),
    };

    (data_type, raw_number, fmt_str, fmt_id)
}

/// Look up a cell's font through the workbook's stylesheet (if loaded).
/// `to_ir` runs after the document has been fully read; if styles weren't
/// parsed yet they remain `None` and we silently skip per-cell font
/// recovery rather than mutate the document during a `&self` traversal.
/// A cell's solid fill colour ([ECMA-376] §18.8.20 fill), resolved
/// through the theme when it names a theme slot.
fn cell_background(
    doc: &crate::xlsx::XlsxDocument,
    cd: &CellData,
    theme: Option<&crate::core::theme::Theme>,
) -> Option<[u8; 3]> {
    let fill = doc.styles.as_ref()?.fill_for(cd.style_index?)?;
    fill.solid_color()?.resolve_opt(theme).map(|rgb| rgb.0)
}

/// A cell's border ([ECMA-376] §18.8.4), one IR line per styled side.
fn cell_border(
    doc: &crate::xlsx::XlsxDocument,
    cd: &CellData,
    theme: Option<&crate::core::theme::Theme>,
) -> Option<TableBorder> {
    let b = doc.styles.as_ref()?.border_for(cd.style_index?)?;
    let line = |side: &Option<crate::xlsx::styles::BorderSide>| {
        let side = side.as_ref()?;
        let (style, size) = border_line_style(&side.style)?;
        Some(BorderLine {
            style,
            color: side
                .color
                .as_ref()
                .and_then(|c| c.resolve_opt(theme))
                .map(|rgb| rgb.0),
            size: Some(size),
            space: None,
        })
    };
    let border = TableBorder {
        top: line(&b.top),
        bottom: line(&b.bottom),
        left: line(&b.left),
        right: line(&b.right),
        inside_h: None,
        inside_v: None,
    };
    [&border.top, &border.bottom, &border.left, &border.right]
        .iter()
        .any(|l| l.is_some())
        .then_some(border)
}

/// An `ST_BorderStyle` ([ECMA-376] §18.18.3) as an IR line style and a
/// width in eighths of a point (thin 1/2 pt, medium 1 pt, thick 1 1/2 pt;
/// hair is thinner still). `none` is no line.
fn border_line_style(style: &str) -> Option<(BorderStyle, u32)> {
    Some(match style {
        "thin" => (BorderStyle::Single, 4),
        "medium" => (BorderStyle::Single, 8),
        "thick" => (BorderStyle::Thick, 12),
        "double" => (BorderStyle::Double, 4),
        "hair" => (BorderStyle::Dotted, 2),
        "dotted" => (BorderStyle::Dotted, 4),
        "dashed" | "dashDot" | "dashDotDot" => (BorderStyle::Dashed, 4),
        "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot" => {
            (BorderStyle::Dashed, 8)
        },
        _ => return None,
    })
}

fn cell_alignment<'d>(
    doc: &'d crate::xlsx::XlsxDocument,
    cd: &CellData,
) -> Option<&'d crate::xlsx::styles::CellAlignment> {
    doc.styles.as_ref()?.alignment_for(cd.style_index?)
}

/// `horizontal` (§18.18.40) as IR paragraph alignment; `general`, `fill`
/// and `centerContinuous` have no IR equivalent and stay unset.
fn cell_h_align(doc: &crate::xlsx::XlsxDocument, cd: &CellData) -> Option<ParagraphAlignment> {
    Some(match cell_alignment(doc, cd)?.horizontal.as_deref()? {
        "left" => ParagraphAlignment::Left,
        "center" => ParagraphAlignment::Center,
        "right" => ParagraphAlignment::Right,
        "justify" => ParagraphAlignment::Justify,
        "distributed" => ParagraphAlignment::Distribute,
        _ => return None,
    })
}

/// `vertical` (§18.18.88) as IR cell vertical alignment.
fn cell_v_align(doc: &crate::xlsx::XlsxDocument, cd: &CellData) -> Option<CellVerticalAlign> {
    Some(match cell_alignment(doc, cd)?.vertical.as_deref()? {
        "top" => CellVerticalAlign::Top,
        "center" => CellVerticalAlign::Center,
        "bottom" => CellVerticalAlign::Bottom,
        _ => return None,
    })
}

fn font_for(
    doc: &crate::xlsx::XlsxDocument,
    style_index: u32,
) -> Option<&crate::xlsx::styles::Font> {
    doc.styles.as_ref()?.font_for(style_index)
}

/// Map a lowercase file extension (`"png"`, `"jpeg"`, ...) to the
/// matching `ImageFormat` variant. Mirrors the PPTX helper. Returns
/// `None` for unrecognised extensions; the round-trip then carries
/// only the bytes (renderers usually sniff the format from the magic
/// header and ignore the missing variant).
fn image_format_from_ext(ext: &str) -> Option<ImageFormat> {
    match ext {
        "png" => Some(ImageFormat::Png),
        "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
        "gif" => Some(ImageFormat::Gif),
        "tif" | "tiff" => Some(ImageFormat::Tiff),
        "bmp" => Some(ImageFormat::Bmp),
        "emf" => Some(ImageFormat::Emf),
        "wmf" => Some(ImageFormat::Wmf),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::charge_picture;
    use crate::limits::TextBudget;

    /// A picture that one budget refuses costs the others nothing, whichever
    /// of them refuses it.
    #[test]
    fn test_a_refused_picture_costs_no_budget_anything() {
        let (mut text, mut image, mut total) = (
            TextBudget::with_limit(100),
            TextBudget::with_limit(10),
            TextBudget::with_limit(100),
        );
        assert!(!charge_picture(&mut text, &mut image, Some(&mut total), 60, 50));
        assert!(image.exhausted());
        assert!(!text.exhausted() && !total.exhausted());
        assert!(text.charge(100) && total.charge(100), "the refused picture spent budget");

        let (mut text, mut image, mut total) = (
            TextBudget::with_limit(100),
            TextBudget::with_limit(100),
            TextBudget::with_limit(10),
        );
        assert!(!charge_picture(&mut text, &mut image, Some(&mut total), 60, 50));
        assert!(total.exhausted());
        assert!(text.charge(100) && image.charge(100), "the refused picture spent budget");

        let (mut text, mut image) = (TextBudget::with_limit(10), TextBudget::with_limit(100));
        assert!(!charge_picture(&mut text, &mut image, None, 60, 50));
        assert!(text.exhausted());
        assert!(image.charge(100), "the refused picture spent budget");

        let (mut text, mut image) = (TextBudget::with_limit(100), TextBudget::with_limit(100));
        assert!(charge_picture(&mut text, &mut image, None, 60, 50));
        assert!(!text.charge(60) && image.charge(50), "an accepted picture is charged to each");
    }
}
