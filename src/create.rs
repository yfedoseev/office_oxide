//! Unified document creation from IR or markdown.

use std::io::{Seek, Write};
use std::path::Path;

use crate::Result;
use crate::format::DocumentFormat;
use crate::ir::*;

/// Create a document file by parsing Markdown and converting it to the target format.
///
/// This is the primary integration bridge between markdown-producing tools and Office documents.
///
/// # Example
///
/// ```rust,no_run
/// use office_oxide::create::create_from_markdown;
/// use office_oxide::format::DocumentFormat;
///
/// let markdown = "# Report\n\nThis is a paragraph.\n\n- item one\n- item two\n";
/// create_from_markdown(markdown, DocumentFormat::Docx, "report.docx").unwrap();
/// ```
pub fn create_from_markdown(
    markdown: &str,
    format: DocumentFormat,
    path: impl AsRef<Path>,
) -> Result<()> {
    let ir = DocumentIR::from_markdown(markdown, format);
    create_from_ir(&ir, format, path)
}

/// Create a document by parsing Markdown and writing to any `Write + Seek` destination.
pub fn create_from_markdown_to_writer<W: Write + Seek>(
    markdown: &str,
    format: DocumentFormat,
    writer: W,
) -> Result<()> {
    let ir = DocumentIR::from_markdown(markdown, format);
    create_from_ir_to_writer(&ir, format, writer)
}

/// Create a document file from a `DocumentIR`.
pub fn create_from_ir(
    ir: &DocumentIR,
    format: DocumentFormat,
    path: impl AsRef<Path>,
) -> Result<()> {
    match format {
        DocumentFormat::Docx => {
            let writer = ir_to_docx(ir);
            writer.save(path)?;
        },
        DocumentFormat::Xlsx => {
            let writer = ir_to_xlsx(ir);
            writer.save(path)?;
        },
        DocumentFormat::Pptx => {
            let writer = ir_to_pptx(ir);
            writer.save(path)?;
        },
        _ => return Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
    }
    Ok(())
}

/// Create a document from IR and write to any `Write + Seek` destination.
pub fn create_from_ir_to_writer<W: Write + Seek>(
    ir: &DocumentIR,
    format: DocumentFormat,
    writer: W,
) -> Result<()> {
    match format {
        DocumentFormat::Docx => {
            let w = ir_to_docx(ir);
            w.write_to(writer)?;
        },
        DocumentFormat::Xlsx => {
            let w = ir_to_xlsx(ir);
            w.write_to(writer)?;
        },
        DocumentFormat::Pptx => {
            let w = ir_to_pptx(ir);
            w.write_to(writer)?;
        },
        _ => return Err(crate::OfficeError::UnsupportedFormat(format!("{format:?}"))),
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// DOCX conversion
// ---------------------------------------------------------------------------

/// Build a `DocxWriter` from `DocumentIR`, exposed so callers can embed
/// extra parts (fonts, custom metadata) before serialization.
///
/// Content past `MAX_NESTING_DEPTH` is dropped rather than overflowing
/// the writer's stack (`DocumentIR` is `Deserialize`, so it can arrive
/// from anywhere, including untrusted input). Call
/// [`crate::docx::write::DocxWriter::truncated_subtrees`] on the
/// returned writer, after `write_to`/`save`, to learn whether that
/// happened — it used to be a silent `Ok` with only a `log::warn!`
/// (issue #218).
pub fn ir_to_docx(ir: &DocumentIR) -> crate::docx::write::DocxWriter {
    use crate::docx::write::{DocxWriter, IrParaProps, Run};

    // The counter is thread-local (see core::xml::DepthGuard) and shared
    // across every DocxWriter built on this thread; reset it here so a
    // fresh ir_to_docx call's count doesn't inherit a previous one's.
    // Building two DocxWriters concurrently, interleaved, on the same
    // thread would confuse the two counts — not a pattern this crate (or
    // realistically any caller) uses.
    crate::core::xml::reset_truncated_subtrees();

    let mut writer = DocxWriter::new();

    // Write metadata
    writer.set_metadata(&ir.metadata);

    if let Some(rgb) = ir.sections.first().and_then(|s| s.background_rgb) {
        writer.set_background_rgb(rgb);
    }

    for section in &ir.sections {
        // Section title becomes H1 — but skip it when the title is already
        // carried by one of the section's own heading elements. The DOCX
        // parser derives `section.title` FROM a heading (which it also
        // keeps in `elements`), so re-emitting the title here would
        // materialise a duplicate H1 on every write→parse cycle, breaking
        // round-trip idempotence. This used to check only
        // `section.elements.first()`, so a section whose heading wasn't
        // its literal first element (e.g. a byline or date line before
        // it) got the heading duplicated on every round trip (issue
        // #259) — checking the whole element list instead of just the
        // first entry fixes it without changing behavior for the common
        // leading-heading case.
        let title_already_present = section.elements.iter().any(|e| {
            matches!(e, Element::Heading(h)
                if section.title.as_deref().is_some_and(|t| inline_to_text(&h.content) == t))
        });
        if let Some(ref title) = section.title {
            if !title.is_empty() && !title_already_present {
                let runs = [Run::new(title)];
                let props = IrParaProps {
                    style: Some("Heading1".to_string()),
                    ..Default::default()
                };
                writer.add_ir_paragraph(&runs, Some(props));
            }
        }

        // Section headers/footers
        if let Some(ref hf) = section.header {
            writer.add_section_header(
                crate::docx::write::HfType::default_header(),
                hf.content.clone(),
            );
        }
        if let Some(ref hf) = section.footer {
            writer.add_section_header(
                crate::docx::write::HfType::default_footer(),
                hf.content.clone(),
            );
        }
        if let Some(ref hf) = section.first_page_header {
            writer.add_section_header(
                crate::docx::write::HfType::first_page_header(),
                hf.content.clone(),
            );
        }
        if let Some(ref hf) = section.first_page_footer {
            writer.add_section_header(
                crate::docx::write::HfType::first_page_footer(),
                hf.content.clone(),
            );
        }
        if let Some(ref hf) = section.even_page_header {
            writer.add_section_header(
                crate::docx::write::HfType::even_page_header(),
                hf.content.clone(),
            );
        }
        if let Some(ref hf) = section.even_page_footer {
            writer.add_section_header(
                crate::docx::write::HfType::even_page_footer(),
                hf.content.clone(),
            );
        }

        for elem in &section.elements {
            add_element_to_docx(&mut writer, elem);
        }

        // Section page setup / columns
        if section.page_setup.is_some()
            || section.columns.is_some()
            || section.break_type != SectionBreakType::Continuous
        {
            writer.set_section_props(
                section.page_setup.clone(),
                section.columns.clone(),
                section.break_type.clone(),
            );
        }
    }

    writer
}

fn add_element_to_docx(writer: &mut crate::docx::write::DocxWriter, elem: &Element) {
    use crate::docx::write::{IrParaProps, Run};

    match elem {
        Element::Heading(h) => {
            let level = h.clamped_level();
            let runs: Vec<Run> = ir_inline_to_runs(&h.content);
            let props = IrParaProps {
                style: Some(format!("Heading{level}")),
                alignment: h.alignment.clone(),
                indent_left_twips: h.indent_left_twips,
                indent_right_twips: h.indent_right_twips,
                first_line_indent_twips: h.first_line_indent_twips,
                space_before_twips: h.space_before_twips,
                space_after_twips: h.space_after_twips,
                line_spacing: h.line_spacing.clone(),
                tabs: h.tabs.clone(),
                frame_position: h.frame_position.clone(),
                keep_with_next: h.keep_with_next,
                keep_together: h.keep_together,
                page_break_before: h.page_break_before,
                background_color: h.background_color,
                border: h.border.clone(),
                ..Default::default()
            };
            writer.add_ir_paragraph(&runs, Some(props));
        },
        Element::Paragraph(p) => {
            let runs = ir_inline_to_runs(&p.content);
            if runs
                .iter()
                .any(|r| !r.text.is_empty() || r.footnote_ref.is_some() || r.endnote_ref.is_some())
            {
                let props = IrParaProps {
                    alignment: p.alignment.clone(),
                    indent_left_twips: p.indent_left_twips,
                    indent_right_twips: p.indent_right_twips,
                    first_line_indent_twips: p.first_line_indent_twips,
                    space_before_twips: p.space_before_twips,
                    space_after_twips: p.space_after_twips,
                    line_spacing: p.line_spacing.clone(),
                    tabs: p.tabs.clone(),
                    frame_position: p.frame_position.clone(),
                    keep_with_next: p.keep_with_next,
                    keep_together: p.keep_together,
                    page_break_before: p.page_break_before,
                    background_color: p.background_color,
                    outline_level: p.outline_level,
                    border: p.border.clone(),
                    ..Default::default()
                };
                writer.add_ir_paragraph(&runs, Some(props));
            }
        },
        Element::Table(t) => {
            writer.add_ir_table(t);
        },
        Element::List(l) => {
            writer.add_ir_list(l);
        },
        Element::Image(img) => {
            writer.add_ir_image(img);
        },
        Element::ThematicBreak => {
            // Emit as a blank paragraph with a single bottom border —
            // the conventional DOCX representation of a horizontal
            // rule. Word displays this as a thin black line under
            // the paragraph; on PDF→DOCX→IR re-parse the renderer
            // detects "empty paragraph with bottom-border-only" and
            // draws a horizontal rule.
            let border = crate::ir::ParagraphBorder {
                top: None,
                left: None,
                right: None,
                between: None,
                bottom: Some(crate::ir::BorderLine {
                    style: crate::ir::BorderStyle::Single,
                    color: Some([0, 0, 0]),
                    size: Some(6),
                    space: Some(1),
                }),
            };
            let props = IrParaProps {
                border: Some(border),
                ..Default::default()
            };
            writer.add_ir_paragraph(&[], Some(props));
        },
        Element::PageBreak => {
            writer.add_page_break();
        },
        Element::ColumnBreak => {
            writer.add_column_break();
        },
        Element::TextBox(tb) => {
            writer.add_text_box(tb);
        },
        Element::Footnote(n) => {
            writer.add_footnote(n.id, &n.content, n.marker.clone());
        },
        Element::Endnote(n) => {
            writer.add_endnote(n.id, &n.content, n.marker.clone());
        },
        Element::CodeBlock(cb) => {
            writer.add_code_block(&cb.content);
        },
        Element::Shape(_) => {
            // Vector shapes are written directly by the layout-preserving
            // DOCX writer (`pdf_oxide::converters::docx_layout`), not via
            // the markdown→IR→DOCX pipeline.
        },
    }
}

fn ir_inline_to_runs(content: &[InlineContent]) -> Vec<crate::docx::write::Run> {
    use crate::docx::write::Run;
    let mut runs: Vec<Run> = Vec::new();
    for item in content {
        match item {
            InlineContent::Text(span) => {
                let mut run = Run::new(&span.text);
                run.bold = span.bold;
                run.italic = span.italic;
                run.strikethrough = span.strikethrough;
                run.font_name = span.font_name.clone();
                run.hyperlink = span.hyperlink.clone();
                run.font_size_half_pt = span.font_size_half_pt;
                run.color_rgb = span.color;
                run.underline_style = span.underline.clone();
                run.highlight = span.highlight;
                run.vertical_align = span.vertical_align.clone();
                run.all_caps = span.all_caps;
                run.small_caps = span.small_caps;
                run.char_spacing_half_pt = span.char_spacing_half_pt;
                runs.push(run);
            },
            InlineContent::LineBreak => {
                runs.push(Run {
                    text: "\n".to_string(),
                    ..Default::default()
                });
            },
            InlineContent::FootnoteRef(r) => {
                runs.push(Run {
                    footnote_ref: Some(r.note_id),
                    note_ref_marker: r.marker.clone(),
                    ..Default::default()
                });
            },
            InlineContent::EndnoteRef(r) => {
                runs.push(Run {
                    endnote_ref: Some(r.note_id),
                    note_ref_marker: r.marker.clone(),
                    ..Default::default()
                });
            },
        }
    }
    coalesce_runs(runs)
}

/// Merge adjacent text runs that share identical run properties so the
/// emitted DOCX has one `<w:r>` per styling region instead of one per
/// PDF span. PDF text extraction returns ~1 span per word; without
/// this pass the document.xml balloons (~5× over the merged form),
/// search/replace breaks across word boundaries, and screen readers
/// stutter.
///
/// Footnote/endnote/field runs are never merged (they carry semantic
/// markers that must stay in their own `<w:r>` for Word to recognise
/// them as references).
fn coalesce_runs(runs: Vec<crate::docx::write::Run>) -> Vec<crate::docx::write::Run> {
    use crate::docx::write::Run;
    let mut out: Vec<Run> = Vec::with_capacity(runs.len());
    for r in runs {
        let mergeable = r.footnote_ref.is_none() && r.endnote_ref.is_none() && r.text != "\n";
        if mergeable {
            if let Some(last) = out.last_mut() {
                if last.footnote_ref.is_none()
                    && last.endnote_ref.is_none()
                    && last.text != "\n"
                    && run_props_equal(last, &r)
                {
                    last.text.push_str(&r.text);
                    continue;
                }
            }
        }
        out.push(r);
    }
    out
}

/// Compare two runs' style properties (everything except `text`,
/// `footnote_ref`, `endnote_ref`) for byte-equality.
fn run_props_equal(a: &crate::docx::write::Run, b: &crate::docx::write::Run) -> bool {
    // hyperlink is part of a run's identity: merging a linked run with an
    // unlinked one silently swallows the link.
    a.hyperlink == b.hyperlink
        && a.bold == b.bold
        && a.italic == b.italic
        && a.underline == b.underline
        && a.underline_style == b.underline_style
        && a.strikethrough == b.strikethrough
        && a.color == b.color
        && a.color_rgb == b.color_rgb
        && a.font_size_pt == b.font_size_pt
        && a.font_size_half_pt == b.font_size_half_pt
        && a.font_name == b.font_name
        && a.highlight == b.highlight
        && a.vertical_align == b.vertical_align
        && a.all_caps == b.all_caps
        && a.small_caps == b.small_caps
        && a.char_spacing_half_pt == b.char_spacing_half_pt
}

// ---------------------------------------------------------------------------
// XLSX conversion
// ---------------------------------------------------------------------------

/// Sanitise a worksheet name and ensure it doesn't clash with names
/// already used in the workbook. Excel limits names to 31 chars and
/// forbids `:\\/?*[]`; the spec also forbids the reserved name
/// "History". When the sanitised candidate is empty or already taken,
/// fall back to "Sheet<idx>" — and even that is post-checked so
/// pathological inputs can't collide.
fn unique_sheet_name(raw: &str, idx: usize, used: &std::collections::HashSet<String>) -> String {
    fn sanitise(s: &str) -> String {
        let mut out = String::with_capacity(s.len().min(31));
        for ch in s.chars() {
            if matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
                out.push('_');
            } else {
                out.push(ch);
            }
            if out.chars().count() >= 31 {
                break;
            }
        }
        out.trim().to_string()
    }
    let candidate = sanitise(raw);
    if !candidate.is_empty()
        && !candidate.eq_ignore_ascii_case("history")
        && !used.contains(&candidate)
    {
        return candidate;
    }
    // Fall back to indexed name.
    let mut fallback = format!("Sheet{idx}");
    let mut bump = idx;
    while used.contains(&fallback) {
        bump += 1;
        fallback = format!("Sheet{bump}");
    }
    fallback
}

/// Build an `XlsxWriter` from `DocumentIR`. Public so callers can embed
/// extra parts (fonts, custom metadata) before serialization. Mirrors
/// `ir_to_docx` and `ir_to_pptx`.
pub fn ir_to_xlsx(ir: &DocumentIR) -> crate::xlsx::write::XlsxWriter {
    use crate::xlsx::write::{CellData, CellStyle};

    let mut writer = crate::xlsx::write::XlsxWriter::new();
    writer.set_metadata(&ir.metadata);

    // Sheet names must be unique within a workbook (ECMA-376) and Excel
    // additionally rejects names > 31 chars, names containing `:\\/?*[]`,
    // and the literal "History". We sanitise + de-duplicate by
    // appending the 1-based index when a section's title would clash
    // (or when there's no title at all).
    let mut used_names: std::collections::HashSet<String> = std::collections::HashSet::new();
    for (idx, section) in ir.sections.iter().enumerate() {
        // Prefer the section's title; failing that, use the first
        // heading inside the section so each tab gets a meaningful
        // label (e.g. "1 Introduction", "Abstract") instead of the
        // anonymous "Sheet1..N".
        let raw_owned = section
            .title
            .clone()
            .or_else(|| first_heading_text(&section.elements))
            .unwrap_or_default();
        let raw = raw_owned.as_str();
        let name = unique_sheet_name(raw, idx + 1, &used_names);
        used_names.insert(name.clone());
        let mut sheet = writer.add_sheet(&name);

        // Propagate per-section page geometry so a PDF→XLSX→PDF round
        // trip preserves the source MediaBox. Without this each
        // worksheet falls back to default Letter portrait and a long
        // PDF (134 / 660 pages) flows onto far fewer pages because the
        // renderer uses a different page size on read-back.
        if let Some(ps) = section.page_setup.as_ref() {
            sheet.set_page_setup(crate::xlsx::write::PageSetup {
                width_twips: ps.width_twips,
                height_twips: ps.height_twips,
                margin_top_twips: ps.margin_top_twips,
                margin_bottom_twips: ps.margin_bottom_twips,
                margin_left_twips: ps.margin_left_twips,
                margin_right_twips: ps.margin_right_twips,
                header_distance_twips: ps.header_distance_twips,
                footer_distance_twips: ps.footer_distance_twips,
                landscape: ps.landscape,
            });
        }

        let mut row_cursor = 0usize;
        // Body paragraphs that aren't part of a table get split across
        // multiple rows when long, so a page-of-prose stays readable
        // instead of piling 1500 chars into a single clipped cell. We
        // also widen column A so the resulting rows have somewhere to
        // breathe. Short paragraphs (≤ 80 chars) and headings stay in
        // a single cell to preserve their visual identity.
        let mut body_paragraphs_seen = false;

        for elem in &section.elements {
            match elem {
                Element::Table(t) => {
                    for (ci, &twips) in t.column_widths_twips.iter().enumerate() {
                        if twips > 0 {
                            let w = (twips as f64) * 96.0 / (1440.0 * 7.0);
                            sheet.set_column_width(ci, w.clamp(3.0, 80.0));
                        }
                    }
                    // Columns currently claimed by a still-open vertical
                    // merge anchored in an earlier row, mapped to how many
                    // more rows (including the one about to be processed)
                    // it still covers. Without this, `col` advanced purely
                    // by summing `col_span` across each row's own cells —
                    // but a row-covering merge means the covered rows have
                    // that column excluded from their own `cells` entirely
                    // (the sparse, span-driven model every reader already
                    // uses), so every cell after the gap silently shifted
                    // left, misplacing content and hyperlinks onto the
                    // wrong column instead of skipping over the claimed
                    // one (issue #340).
                    let mut active_spans: std::collections::BTreeMap<usize, usize> =
                        std::collections::BTreeMap::new();
                    for row in &t.rows {
                        let mut col = 0usize;
                        for cell in &row.cells {
                            while active_spans.get(&col).is_some_and(|&remaining| remaining > 0) {
                                col += 1;
                            }
                            let text = cell_text(cell);
                            let data = ir_cell_to_cell_data(cell, &text);
                            if let Some(style) = xlsx_cell_style(
                                row.is_header,
                                cell.background_color,
                                cell.number_format.as_deref(),
                            ) {
                                sheet.set_cell_styled(row_cursor, col, data, style);
                            } else {
                                sheet.set_cell(row_cursor, col, data);
                            }
                            if let Some(url) = cell_hyperlink(cell) {
                                sheet.set_cell_hyperlink(row_cursor, col, url);
                            }
                            let cs = cell.col_span.max(1) as usize;
                            let rs = cell.row_span.max(1) as usize;
                            if cs > 1 || rs > 1 {
                                sheet.merge_cells(row_cursor, col, rs, cs);
                            }
                            if rs > 1 {
                                for c in col..col + cs {
                                    active_spans.insert(c, rs);
                                }
                            }
                            col += cs;
                        }
                        // This row is now consumed: every still-open span
                        // (including any minted just above) has one fewer
                        // row left to cover.
                        for v in active_spans.values_mut() {
                            *v = v.saturating_sub(1);
                        }
                        active_spans.retain(|_, v| *v > 0);
                        row_cursor += 1;
                    }
                },
                Element::Paragraph(p) => {
                    let text = inline_to_text(&p.content);
                    if !text.is_empty() {
                        body_paragraphs_seen = true;
                        // Persist the IR paragraph's font size onto the cell.
                        // This is what allows a PDF→IR→XLSX→IR→PDF round-trip
                        // to recover the original 9–10 pt body size instead of
                        // falling back to the 12 pt default and inflating the
                        // page count.
                        let mut style = CellStyle::new();
                        if let Some(size_pt) = crate::ir::first_inline_font_size_pt(&p.content) {
                            style = style.font_size(size_pt);
                        }
                        if let Some(name) = first_inline_font_name(&p.content) {
                            style = style.font_name(name);
                        }
                        // The first hyperlink anywhere in the paragraph,
                        // attached to the first resulting row — this
                        // fallback (non-tabular content split across
                        // cells) doesn't track which specific sub-line a
                        // hyperlink's own span covers, matching the
                        // simplification `cell_hyperlink` already makes
                        // for a real table cell's content (issue #262).
                        let hyperlink = p.content.iter().find_map(|c| match c {
                            InlineContent::Text(t) => t.hyperlink.clone(),
                            _ => None,
                        });
                        for (i, line) in split_paragraph_for_xlsx(&text).into_iter().enumerate() {
                            sheet.set_cell_styled(
                                row_cursor,
                                0,
                                CellData::String(line),
                                style.clone(),
                            );
                            if i == 0 {
                                if let Some(ref url) = hyperlink {
                                    sheet.set_cell_hyperlink(row_cursor, 0, url.clone());
                                }
                            }
                            row_cursor += 1;
                        }
                    }
                },
                Element::Image(img) => {
                    // Anchor any image carried by the IR onto this
                    // worksheet. EMU coordinates default to (0, 0) when
                    // the IR didn't carry per-image positioning — the
                    // round-trip still recovers the bytes, just stacked
                    // at the sheet origin. When position-aware writers
                    // wrap images in TextBox the outer branch below
                    // unwraps the EMU coords.
                    if let (Some(data), Some(fmt)) = (&img.data, &img.format) {
                        let cx = img.display_width_emu.unwrap_or(3_000_000) as i64;
                        let cy = img.display_height_emu.unwrap_or(2_000_000) as i64;
                        sheet.add_image(data.clone(), fmt.extension(), 0, 0, cx, cy);
                    }
                },
                Element::TextBox(tb) => {
                    // Positional wrapper: when the IR places an image
                    // inside a TextBox (PDF→IR can carry shape coords
                    // that way), forward the inner image bytes with the
                    // TextBox's anchor.
                    let x = tb.x_emu.unwrap_or(0);
                    let y = tb.y_emu.unwrap_or(0);
                    let cx = tb.width_emu.unwrap_or(0) as i64;
                    let cy = tb.height_emu.unwrap_or(0) as i64;
                    for inner in &tb.content {
                        if let Element::Image(img) = inner {
                            if let (Some(data), Some(fmt)) = (&img.data, &img.format) {
                                let icx = if cx > 0 {
                                    cx
                                } else {
                                    img.display_width_emu.unwrap_or(3_000_000) as i64
                                };
                                let icy = if cy > 0 {
                                    cy
                                } else {
                                    img.display_height_emu.unwrap_or(2_000_000) as i64
                                };
                                sheet.add_image(data.clone(), fmt.extension(), x, y, icx, icy);
                            }
                        }
                    }
                    // Everything in the box that is not an image still has
                    // text; the arm used to look for images and nothing else.
                    let mut rows = Vec::new();
                    for inner in &tb.content {
                        if !matches!(inner, Element::Image(_)) {
                            xlsx_text_rows(inner, &mut rows);
                        }
                    }
                    for line in rows {
                        body_paragraphs_seen = true;
                        sheet.set_cell(row_cursor, 0, CellData::String(line));
                        row_cursor += 1;
                    }
                },
                Element::Heading(h) => {
                    let text = inline_to_text(&h.content);
                    if !text.is_empty() {
                        let data = CellData::String(text);
                        let mut style = CellStyle::new().bold();
                        if let Some(size_pt) = crate::ir::first_inline_font_size_pt(&h.content) {
                            style = style.font_size(size_pt);
                        }
                        if let Some(name) = first_inline_font_name(&h.content) {
                            style = style.font_name(name);
                        }
                        sheet.set_cell_styled(row_cursor, 0, data, style);
                        row_cursor += 1;
                    }
                },
                Element::ThematicBreak
                | Element::PageBreak
                | Element::ColumnBreak
                | Element::Shape(_) => {},
                // An XLSX-sourced cell comment round-trips as
                // Element::Endnote (convert_xlsx.rs appends one per
                // comment, marker = "{cell_ref}" or "{cell_ref}
                // ({author})"). Written back as a real cell comment
                // instead of falling into the generic wildcard below,
                // which dumped every comment's text as a spurious extra
                // row with no cell reference at all (issue #344).
                Element::Endnote(n) => {
                    let placed = parse_xlsx_comment_marker(n.marker.as_deref())
                        .and_then(|(cell_ref, author)| {
                            crate::xlsx::CellRef::parse(&cell_ref).map(|r| (r, author))
                        })
                        .map(|(r, author)| {
                            let mut rows = Vec::new();
                            for inner in &n.content {
                                xlsx_text_rows(inner, &mut rows);
                            }
                            let text = rows.join("\n");
                            if !text.is_empty() {
                                sheet.set_cell_comment(r.row as usize, r.col as usize, author, text);
                            }
                        })
                        .is_some();
                    // A marker that doesn't match the XLSX "{cell_ref}
                    // ({author})" shape (an endnote from a non-XLSX
                    // source converted to a spreadsheet, say) still has
                    // real text — fall back to the generic row dump
                    // rather than silently dropping it.
                    if !placed {
                        let mut rows = Vec::new();
                        for inner in &n.content {
                            xlsx_text_rows(inner, &mut rows);
                        }
                        for line in rows {
                            body_paragraphs_seen = true;
                            sheet.set_cell(row_cursor, 0, CellData::String(line));
                            row_cursor += 1;
                        }
                    }
                },
                // Everything else has no native spreadsheet shape but does
                // have text, and dropping it silently is how a PPTX → XLSX
                // conversion lost 96% of its words.
                other => {
                    let mut rows = Vec::new();
                    xlsx_text_rows(other, &mut rows);
                    for line in rows {
                        body_paragraphs_seen = true;
                        sheet.set_cell(row_cursor, 0, CellData::String(line));
                        row_cursor += 1;
                    }
                },
            }
        }

        // If we emitted any body paragraphs (rather than just tables)
        // widen column A so multi-line prose has somewhere to render.
        // Tables manage their own per-column widths above so we leave
        // those alone.
        if body_paragraphs_seen {
            sheet.set_column_width(0, 80.0);
        }
    }

    writer
}

/// Split a long paragraph into ~120-char chunks at sentence boundaries
/// for XLSX rendering. Short paragraphs (≤ 80 chars) pass through as a
/// single chunk so they keep their compact look.
///
/// Operates on `char_indices` throughout so the byte indices we slice
/// at are always valid UTF-8 boundaries — paragraphs from PDFs often
/// contain multi-byte glyphs (mathematical italic, accented Latin,
/// CJK) and naive byte arithmetic blows up on them.
fn split_paragraph_for_xlsx(text: &str) -> Vec<String> {
    const SHORT_THRESHOLD: usize = 80;
    const TARGET_LINE_LEN: usize = 120;
    const SCAN_BACK_CHARS: usize = 60;

    if text.chars().count() <= SHORT_THRESHOLD {
        return vec![text.to_string()];
    }

    // Pre-compute char positions so all slicing happens on boundaries.
    let chars: Vec<(usize, char)> = text.char_indices().collect();
    let total_chars = chars.len();
    let total_bytes = text.len();

    let mut chunks: Vec<String> = Vec::new();
    let mut char_start: usize = 0; // index into `chars`

    while char_start < total_chars {
        let remaining_chars = total_chars - char_start;
        if remaining_chars <= TARGET_LINE_LEN {
            let head_byte = chars[char_start].0;
            let tail = text[head_byte..].trim();
            if !tail.is_empty() {
                chunks.push(tail.to_string());
            }
            break;
        }

        // The "minimum break point" is char_start + TARGET_LINE_LEN.
        let min_break_char = char_start + TARGET_LINE_LEN;
        let scan_back_char = min_break_char
            .saturating_sub(SCAN_BACK_CHARS)
            .max(char_start);

        // Find a sentence boundary: a `.` followed by ` ` followed by
        // an uppercase ASCII letter. Prefer breaks at or after the
        // target, then fall back to one slightly before.
        let mut break_char: Option<usize> = None;

        // Pass 1: at-or-after the cap.
        for i in min_break_char..total_chars.saturating_sub(2) {
            if chars[i].1 == '.' && chars[i + 1].1 == ' ' && chars[i + 2].1.is_ascii_uppercase() {
                break_char = Some(i + 2); // start of the next sentence
                break;
            }
        }

        // Pass 2: before the cap, within scan_back window.
        if break_char.is_none() {
            for i in scan_back_char..min_break_char.saturating_sub(2).max(scan_back_char) {
                if i + 2 >= total_chars {
                    break;
                }
                if chars[i].1 == '.' && chars[i + 1].1 == ' ' && chars[i + 2].1.is_ascii_uppercase()
                {
                    break_char = Some(i + 2);
                }
            }
        }

        // Pass 3: next whitespace at-or-after the cap.
        if break_char.is_none() {
            for i in min_break_char..total_chars {
                if chars[i].1 == ' ' {
                    break_char = Some(i + 1);
                    break;
                }
            }
        }

        let next_char = break_char.unwrap_or(total_chars);
        let head_byte = chars[char_start].0;
        let tail_byte = if next_char >= total_chars {
            total_bytes
        } else {
            chars[next_char].0
        };
        let head = text[head_byte..tail_byte].trim();
        if !head.is_empty() {
            chunks.push(head.to_string());
        }

        // Advance past any leading whitespace on the tail (we already
        // trimmed `head`, but `next_char` may sit right at the space).
        let mut cs = next_char;
        while cs < total_chars && chars[cs].1 == ' ' {
            cs += 1;
        }
        if cs <= char_start {
            // Defensive: ensure forward progress.
            cs = char_start + 1;
        }
        char_start = cs;
    }

    if chunks.is_empty() {
        chunks.push(text.to_string());
    }
    chunks
}

// ---------------------------------------------------------------------------
// PPTX conversion
// ---------------------------------------------------------------------------

/// Build a `PptxWriter` from `DocumentIR`. Public so callers can embed
/// extra parts (fonts, custom metadata) before serialization.
pub fn ir_to_pptx(ir: &DocumentIR) -> crate::pptx::write::PptxWriter {
    let mut writer = crate::pptx::write::PptxWriter::new();
    writer.set_metadata(&ir.metadata);

    if let Some(ps) = ir.sections.iter().find_map(|s| s.page_setup.as_ref()) {
        let cx = ps.width_twips as u64 * 914_400 / 1440;
        let cy = ps.height_twips as u64 * 914_400 / 1440;
        writer.set_presentation_size(cx, cy);
    }

    // PowerPoint shows a "found a problem with content. Do you want to
    // repair?" dialog and renders Slide Sorter very slowly when a deck
    // exceeds ~250 slides. For large PDFs (e.g. a 660-page CFR) the
    // historical 1-section-per-slide mapping produces decks that hit
    // both issues. When the IR has more sections than the threshold
    // we collapse consecutive sections into heading-bounded chunks of
    // at most ~12 paragraphs each and cap the total slide count.
    const MAX_SLIDES: usize = 250;
    const MAX_PARAGRAPHS_PER_SLIDE: usize = 12;

    if ir.sections.len() <= MAX_SLIDES {
        for section in &ir.sections {
            emit_pptx_slide_from_section(&mut writer, section);
        }
    } else {
        emit_pptx_slides_compacted(&mut writer, ir, MAX_SLIDES, MAX_PARAGRAPHS_PER_SLIDE);
    }

    writer
}

/// One IR section → one slide. Used for "small" decks where 1:1 paging
/// is still viable.
fn emit_pptx_slide_from_section(writer: &mut crate::pptx::write::PptxWriter, section: &Section) {
    let slide = writer.add_slide();

    // Speaker notes round-trip into ppt/notesSlides/, never onto the slide.
    if let Some(ref notes) = section.speaker_notes {
        if !notes.is_empty() {
            slide.set_notes(notes);
        }
    }

    if let Some(ref title) = section.title {
        if !title.is_empty() {
            slide.set_title(title);
        }
    }

    // The PPTX parser surfaces the slide-title placeholder as BOTH
    // `section.title` and a Heading element somewhere in `elements`. The
    // title placeholder is already set above, so re-emitting that
    // heading as body text would both pollute the body and duplicate it
    // on a write→parse cycle (breaking round-trip idempotence). This
    // used to only check `section.elements.first()` — the same bug
    // shape as #259 on the DOCX side — so a slide whose title heading
    // wasn't its literal first shape (a decorative or subtitle shape
    // ahead of it in z-order) got the title duplicated into the body on
    // every round trip. Find the matching heading wherever it is and
    // skip only that one occurrence, not every element after it.
    let title_heading_idx = section.title.as_deref().and_then(|t| {
        section
            .elements
            .iter()
            .position(|e| matches!(e, Element::Heading(h) if inline_to_text(&h.content) == t))
    });

    for (i, elem) in section.elements.iter().enumerate() {
        if Some(i) == title_heading_idx {
            continue;
        }
        emit_pptx_element(slide, elem);
    }
}

/// Marker text used to encode `Element::ThematicBreak` through PPTX
/// round-trip. The PPTX paragraph format has no `<a:pPr>` border the
/// way DOCX `<w:pBdr>` does; emitting a thin connector shape would
/// position the rule absolutely on the slide (wrong for flow
/// content). Instead we emit a centered paragraph of U+2500 (BOX
/// DRAWINGS LIGHT HORIZONTAL) characters; the renderer's pdf_oxide
/// side detects this exact pattern and re-emits a real
/// `page.horizontal_rule()`. Plain enough that any other consumer
/// (PowerPoint itself, a markdown export, a screen reader) sees a
/// visible horizontal-rule glyph string and treats it as a
/// separator.
pub(crate) const PPTX_THEMATIC_BREAK_MARKER: &str = "\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}\u{2500}";

fn emit_pptx_element(slide: &mut crate::pptx::write::SlideData, elem: &Element) {
    match elem {
        Element::ThematicBreak => {
            // Encode via the marker text + center alignment. The
            // pdf_oxide renderer recognises the all-U+2500 content
            // and draws a real `page.horizontal_rule()` instead of
            // rendering the box-drawing glyphs.
            let runs = vec![crate::pptx::write::Run::new(PPTX_THEMATIC_BREAK_MARKER)];
            slide.add_rich_text_aligned(&runs, Some(ParagraphAlignment::Center));
        },
        Element::Heading(h) => {
            if slide.title.is_none() {
                slide.set_title_aligned(&inline_to_text(&h.content), h.alignment.clone());
            } else {
                let runs = inline_to_pptx_runs(&h.content);
                if !runs.is_empty() {
                    slide.add_rich_text_aligned(&runs, h.alignment.clone());
                }
            }
        },
        Element::Paragraph(p) => {
            let runs = inline_to_pptx_runs(&p.content);
            // Always emit, including for runs.is_empty() — empty
            // spacer paragraphs (used by pdf_to_ir to preserve large
            // vertical gaps on source cover pages) need to round-trip
            // through PPTX as empty <a:p> elements so the rendered
            // PPTX→IR→PDF cycle reproduces the source's vertical
            // rhythm. `space_before_twips` from IR (twips) is
            // converted to PPTX `<a:spcPts>` hundredths-of-pt:
            // 1 twip = 1/1440 in = 1/20 pt → twips * 5 = pt*100.
            let space_before_hundredths_pt = p.space_before_twips.map(|t| t * 5);
            let props = crate::pptx::write::ParaProps {
                alignment: p.alignment.clone(),
                space_before_hundredths_pt,
            };
            slide.add_rich_text_with_props(&runs, props);
        },
        Element::List(l) => {
            // Flatten the whole tree: `ListItem::nested` was read by no
            // writer, so every item below level 0 was silently dropped.
            fn flatten(list: &crate::ir::List, level: u8, out: &mut Vec<(u8, String)>) {
                for item in &list.items {
                    let text = item
                        .content
                        .iter()
                        .map(|e| match e {
                            Element::Paragraph(p) => inline_to_text(&p.content),
                            _ => String::new(),
                        })
                        .collect::<Vec<_>>()
                        .join(" ");
                    if !text.is_empty() {
                        out.push((level, text));
                    }
                    if let Some(ref nested) = item.nested {
                        flatten(nested, level.saturating_add(1), out);
                    }
                }
            }
            let mut items: Vec<(u8, String)> = Vec::new();
            flatten(l, l.level, &mut items);
            slide.add_nested_bullet_list(&items);
        },
        Element::Table(t) => {
            // A real a:tbl, not tab-joined text: the previous form lost the
            // grid, every cell boundary and all per-cell formatting.
            let rows: Vec<Vec<String>> = t
                .rows
                .iter()
                .map(|row| row.cells.iter().map(cell_text).collect())
                .collect();
            slide.add_table(rows);
        },
        Element::Image(img) => {
            if let (Some(data), Some(fmt)) = (&img.data, &img.format) {
                let cx = img.display_width_emu.unwrap_or(3_000_000);
                let cy = img.display_height_emu.unwrap_or(2_000_000);
                slide.add_image_with_alt(
                    data.clone(),
                    fmt.clone(),
                    0,
                    0,
                    cx,
                    cy,
                    img.alt_text.clone(),
                );
            }
        },
        Element::CodeBlock(cb) => {
            let run = crate::pptx::write::Run::new(&cb.content).font("Courier New");
            slide.add_rich_text(&[run]);
        },
        Element::TextBox(tb) => {
            // A `TextBox` in the IR only ever exists because the reader
            // saw a real `<a:xfrm>` on the source shape (see
            // `push_positional_textbox` in convert_pptx.rs — content
            // without one flows as plain elements and is never wrapped
            // at all). This includes a body PLACEHOLDER that happens to
            // inherit a real position, which is why a `List`/`Table` can
            // show up as a TextBox's content too, not just a genuine
            // free-floating text box.
            //
            // Flattening everything into the generic slide-body stream
            // (the old behavior) threw the box's own position away and
            // merged multiple independent text boxes' content into one
            // undifferentiated stream — or conjured a spurious empty box
            // out of unrelated body content once that flattening was
            // "fixed" to stop dropping text (issue #264). Paragraph/
            // Heading content — the common "this really is just a text
            // box" case — is now emitted as its own positioned shape,
            // the same way `Element::Image` already does. A nested
            // `List`/`Table`/`TextBox` — which has no positioned
            // representation of its own here — still flattens into the
            // body stream as before, preserving its own structure rather
            // than losing it to plain text (and keeping write→parse
            // idempotent: a body placeholder wrapping a bulleted list
            // reads back into the same TextBox{List} shape either way).
            let mut paragraphs = Vec::new();
            for inner in &tb.content {
                match inner {
                    Element::Paragraph(_) | Element::Heading(_) => {
                        let single = std::slice::from_ref(inner);
                        paragraphs.extend(textbox_content_to_pptx_paragraphs(single));
                    },
                    // `Element::Image`'s own arm always anchors at (0, 0)
                    // — routing it through `emit_pptx_element` like the
                    // other fallback cases would silently move it there,
                    // discarding the *wrapping* text box's real position
                    // (issue #264 — confirmed on real corpus files: a
                    // positioned image inside a TextBox lost its position
                    // and stopped counting as a positioned shape at all).
                    Element::Image(img) => {
                        if let (Some(data), Some(fmt)) = (&img.data, &img.format) {
                            let cx =
                                img.display_width_emu.unwrap_or(tb.width_emu.unwrap_or(3_000_000));
                            let cy = img
                                .display_height_emu
                                .unwrap_or(tb.height_emu.unwrap_or(2_000_000));
                            slide.add_image_with_alt(
                                data.clone(),
                                fmt.clone(),
                                tb.x_emu.unwrap_or(0),
                                tb.y_emu.unwrap_or(0),
                                cx,
                                cy,
                                img.alt_text.clone(),
                            );
                        }
                    },
                    _ => emit_pptx_element(slide, inner),
                }
            }
            if !paragraphs.is_empty() {
                let x = tb.x_emu.unwrap_or(0);
                let y = tb.y_emu.unwrap_or(0);
                let cx = tb.width_emu.map(|w| w as i64).unwrap_or(3_000_000);
                let cy = tb.height_emu.map(|h| h as i64).unwrap_or(500_000);
                slide.add_multi_paragraph_text_box(paragraphs, x, y, cx, cy);
            }
        },
        // Footnote and endnote bodies have no slide equivalent, but their
        // text is unambiguous content — the catch-all used to drop it.
        Element::Footnote(n) | Element::Endnote(n) => {
            for inner in &n.content {
                emit_pptx_element(slide, inner);
            }
        },
        Element::PageBreak | Element::ColumnBreak | Element::Shape(_) => {},
    }
}

/// Heading-aware compaction for large IR section lists.
///
/// Strategy:
/// 1. Build a flat list of `(title, elements)` "groups" where every
///    H1/H2 boundary starts a new group and the section's own
///    elements between headings are concatenated.
/// 2. Each group becomes one or more slides, splitting at paragraph
///    boundaries when the body exceeds `max_paragraphs_per_slide`.
/// 3. After collecting candidate slides, if we still exceed
///    `max_slides`, fold trailing slides into the previous one until
///    the cap is met (preserves earlier headings/structure).
fn emit_pptx_slides_compacted(
    writer: &mut crate::pptx::write::PptxWriter,
    ir: &DocumentIR,
    max_slides: usize,
    max_paragraphs_per_slide: usize,
) {
    // Step 1: build heading-bounded groups. Each group's title is
    // (text, optional alignment); the alignment flows through to
    // `slide.set_title_aligned` in step 4 so source-PDF cover-page
    // headings keep their original alignment (typically Center).
    type TitleWithAlgn = (String, Option<ParagraphAlignment>);
    let mut groups: Vec<(Option<TitleWithAlgn>, Vec<Element>)> = Vec::new();
    let mut current_title: Option<TitleWithAlgn> = None;
    let mut current_elems: Vec<Element> = Vec::new();
    // Tracks whether the current group has accumulated any genuine
    // body content (non-heading element). When false, an incoming
    // H1/H2 is folded into the current slide as a subtitle instead of
    // starting a new one. This prevents cover pages — where each
    // title-block line is promoted to a heading by `pdf_to_ir` — from
    // exploding into one title-only slide per line.
    let mut current_has_body = false;

    let flush = |groups: &mut Vec<(Option<TitleWithAlgn>, Vec<Element>)>,
                 title: &mut Option<TitleWithAlgn>,
                 elems: &mut Vec<Element>| {
        if !elems.is_empty() || title.is_some() {
            groups.push((title.take(), std::mem::take(elems)));
        }
    };

    // Whether an element constitutes "body content" for compaction
    // purposes. Cover pages typically begin with a logo or seal Image
    // and a list of centered headings; flipping `current_has_body` on
    // the leading Image causes the first heading to fall into the
    // "real new section" branch and strand the image as a title-less
    // slide. Only text-bearing elements should anchor a slide as
    // having body content. Empty paragraphs used as vertical spacers
    // (no runs, no border) are skipped — they're layout glue, not
    // content; counting them as body causes cover pages to split
    // mid-block when pdf_to_ir injects gap spacers.
    fn is_body_content(elem: &Element) -> bool {
        match elem {
            Element::Paragraph(p) => p.content.iter().any(|ic| match ic {
                InlineContent::Text(s) => !s.text.is_empty(),
                _ => false,
            }),
            Element::List(_) | Element::CodeBlock(_) | Element::Table(_) => true,
            _ => false,
        }
    }

    for section in &ir.sections {
        for elem in &section.elements {
            if let Element::Heading(h) = elem {
                if h.level <= 2 {
                    let text = inline_to_text(&h.content);
                    let trimmed = text.trim();
                    if trimmed.is_empty() {
                        continue;
                    }

                    if !current_has_body {
                        // Cover-page fold: keep all consecutive
                        // headings on the same slide. First heading
                        // owns the slide title; subsequent headings
                        // become bold paragraphs so they stay visible.
                        if current_title.is_none() {
                            current_title = Some((trimmed.to_string(), h.alignment.clone()));
                        } else {
                            let mut span = TextSpan::plain(trimmed.to_string());
                            span.bold = true;
                            current_elems.push(Element::Paragraph(Paragraph {
                                content: vec![InlineContent::Text(span)],
                                alignment: h.alignment.clone(),
                                ..Default::default()
                            }));
                        }
                        continue;
                    }

                    // Real new section: flush and open a new group.
                    flush(&mut groups, &mut current_title, &mut current_elems);
                    current_has_body = false;
                    current_title = Some((trimmed.to_string(), h.alignment.clone()));
                    continue;
                }
            }
            current_elems.push(elem.clone());
            if is_body_content(elem) {
                current_has_body = true;
            }
        }
    }
    flush(&mut groups, &mut current_title, &mut current_elems);

    // If the IR had no H1/H2 headings at all we end up with a single
    // group holding everything. That would be one slide with all the
    // content packed in, which the renderer can't actually fit. Fall
    // back to a paragraph-count partition over the flattened element
    // stream.
    if groups.len() <= 1 {
        let mut all_elems: Vec<Element> = Vec::new();
        for section in &ir.sections {
            for elem in &section.elements {
                all_elems.push(elem.clone());
            }
        }
        groups = vec![(None, all_elems)];
    }

    // Step 2: split each group into slide-sized chunks.
    struct PendingSlide {
        title: Option<(String, Option<ParagraphAlignment>)>,
        elements: Vec<Element>,
    }
    let mut pending: Vec<PendingSlide> = Vec::new();

    for (title, elems) in groups {
        let mut chunk: Vec<Element> = Vec::new();
        let mut paragraph_count = 0usize;
        let mut first_chunk = true;
        for elem in elems {
            let is_paragraph_like =
                matches!(elem, Element::Paragraph(_) | Element::List(_) | Element::CodeBlock(_));
            if is_paragraph_like && paragraph_count >= max_paragraphs_per_slide {
                pending.push(PendingSlide {
                    title: if first_chunk { title.clone() } else { None },
                    elements: std::mem::take(&mut chunk),
                });
                paragraph_count = 0;
                first_chunk = false;
            }
            if is_paragraph_like {
                paragraph_count += 1;
            }
            chunk.push(elem);
        }
        if !chunk.is_empty() || (first_chunk && title.is_some()) {
            pending.push(PendingSlide {
                title: if first_chunk { title.clone() } else { None },
                elements: chunk,
            });
        }
    }

    // Step 3: enforce the slide cap by folding trailing slides into
    // the previous one. We always keep at least one slide.
    while pending.len() > max_slides {
        // Pop the last slide and append its elements to the previous.
        let tail = pending.pop().expect("pending non-empty");
        if let Some(prev) = pending.last_mut() {
            prev.elements.extend(tail.elements);
        } else {
            pending.push(tail);
            break;
        }
    }

    // Step 4: emit slides.
    for ps in pending {
        let slide = writer.add_slide();
        if let Some((t, algn)) = ps.title.as_ref() {
            if !t.is_empty() {
                slide.set_title_aligned(t, algn.clone());
            }
        }
        for elem in &ps.elements {
            emit_pptx_element(slide, elem);
        }
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn rgb_to_hex(rgb: [u8; 3]) -> String {
    format!("{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])
}

fn cell_text(cell: &TableCell) -> String {
    cell.content
        .iter()
        .map(|e| match e {
            Element::Paragraph(p) => inline_to_text(&p.content),
            _ => String::new(),
        })
        .collect::<Vec<_>>()
        .join(" ")
}

/// The first hyperlink URL found anywhere in a cell's content, if any
/// (issue #262 — xlsx::write had no hyperlink concept at all, so a
/// cell's `TextSpan.hyperlink` — the same field DOCX/PPTX runs already
/// use — was silently dropped on every write).
fn cell_hyperlink(cell: &TableCell) -> Option<String> {
    cell.content.iter().find_map(|e| match e {
        Element::Paragraph(p) => p.content.iter().find_map(|c| match c {
            InlineContent::Text(t) => t.hyperlink.clone(),
            _ => None,
        }),
        _ => None,
    })
}

/// Convert an IR cell to writer data, honouring the type the parser recorded.
///
/// Re-parsing the *rendered* string threw away `data_type`, `raw_number` and
/// `number_format`: "007" became the number 7, a currency cell became text,
/// and a cell whose text happens to read "inf" or "NaN" became an Excel error
/// cell. The typed fields are only populated by the XLSX reader; prose formats
/// leave them `None` and still fall back to sniffing the text.
fn ir_cell_to_cell_data(cell: &TableCell, text: &str) -> crate::xlsx::write::CellData {
    use crate::ir::CellDataType;
    use crate::xlsx::write::CellData;

    match cell.data_type {
        Some(CellDataType::Number) | Some(CellDataType::Date) => {
            if let Some(n) = cell.raw_number {
                return CellData::Number(n);
            }
        },
        Some(CellDataType::Boolean) => {
            let t = text.trim();
            if t.eq_ignore_ascii_case("true") || t == "1" {
                return CellData::Boolean(true);
            }
            if t.eq_ignore_ascii_case("false") || t == "0" {
                return CellData::Boolean(false);
            }
        },
        // Text and Error cells keep their rendered form verbatim; sniffing
        // would re-introduce the "007" and "inf" corruption.
        Some(CellDataType::Text) | Some(CellDataType::Error) => {
            return if text.is_empty() {
                CellData::Empty
            } else {
                CellData::String(text.to_string())
            };
        },
        None => {},
    }
    text_to_cell_data(text)
}

fn text_to_cell_data(text: &str) -> crate::xlsx::write::CellData {
    use crate::xlsx::write::CellData;
    if text.is_empty() {
        CellData::Empty
    } else if let Ok(n) = text.parse::<f64>() {
        CellData::Number(n)
    } else {
        CellData::String(text.to_string())
    }
}

/// Pluck the first `Element::Heading`'s plain text from a section's
/// element list. Used by `ir_to_xlsx` to derive a meaningful
/// worksheet tab label when the section itself doesn't carry a
/// title — typical for a PDF→IR conversion where heading detection
/// happens at the element level, not the section level.
fn first_heading_text(elements: &[Element]) -> Option<String> {
    for el in elements {
        if let Element::Heading(h) = el {
            let text = inline_to_text(&h.content);
            let trimmed = text.trim();
            if !trimmed.is_empty() {
                return Some(trimmed.to_string());
            }
        }
    }
    None
}

/// First explicit font name from inline content. Used by the XLSX
/// path so cell styles carry the source font instead of always
/// falling back to the writer's "Calibri" default. Mirrors the
/// `first_inline_font_size_pt` helper.
fn first_inline_font_name(content: &[InlineContent]) -> Option<String> {
    for ic in content {
        if let InlineContent::Text(span) = ic {
            if let Some(name) = &span.font_name {
                if !name.is_empty() {
                    return Some(name.clone());
                }
            }
        }
    }
    None
}

/// Flatten an element the XLSX writer has no native shape for into one text
/// row per logical line.
///
/// `ir_to_xlsx` used to end in `_ => {}`, silently swallowing `List`,
/// `CodeBlock`, `Footnote`, `Endnote` and everything inside a `TextBox` that
/// was not an image. Converting a presentation to a spreadsheet lost almost
/// all of its text that way.
/// Parse an XLSX-sourced comment's `Note::marker` — `"{cell_ref}"` or
/// `"{cell_ref} ({author})"`, exactly the shape `convert_xlsx.rs`
/// generates for every `Element::Endnote` it builds from a real cell
/// comment — back into `(cell_ref, author)`. Returns `None` for any
/// other shape (no marker, or one that doesn't parse this way), so the
/// caller can fall back to treating it as ordinary endnote content
/// (issue #344).
fn parse_xlsx_comment_marker(marker: Option<&str>) -> Option<(String, Option<String>)> {
    let marker = marker?;
    match marker.split_once(" (") {
        Some((cell_ref, rest)) => {
            let author = rest.strip_suffix(')')?;
            Some((cell_ref.to_string(), Some(author.to_string())))
        },
        None => Some((marker.to_string(), None)),
    }
}

fn xlsx_text_rows(elem: &Element, out: &mut Vec<String>) {
    match elem {
        Element::Paragraph(p) => {
            let t = inline_to_text(&p.content);
            if !t.is_empty() {
                out.push(t);
            }
        },
        Element::Heading(h) => {
            let t = inline_to_text(&h.content);
            if !t.is_empty() {
                out.push(t);
            }
        },
        Element::List(l) => {
            fn walk(list: &crate::ir::List, out: &mut Vec<String>) {
                for item in &list.items {
                    for e in &item.content {
                        xlsx_text_rows(e, out);
                    }
                    if let Some(ref nested) = item.nested {
                        walk(nested, out);
                    }
                }
            }
            walk(l, out);
        },
        Element::CodeBlock(cb) => {
            for line in cb.content.lines() {
                out.push(line.to_string());
            }
        },
        Element::Footnote(n) | Element::Endnote(n) => {
            for e in &n.content {
                xlsx_text_rows(e, out);
            }
        },
        Element::Table(t) => {
            for row in &t.rows {
                let line = row
                    .cells
                    .iter()
                    .map(cell_text)
                    .collect::<Vec<_>>()
                    .join("\t");
                if !line.trim().is_empty() {
                    out.push(line);
                }
            }
        },
        Element::TextBox(tb) => {
            for e in &tb.content {
                xlsx_text_rows(e, out);
            }
        },
        Element::Image(_)
        | Element::ThematicBreak
        | Element::PageBreak
        | Element::ColumnBreak
        | Element::Shape(_) => {},
    }
}

fn xlsx_cell_style(
    is_header: bool,
    bg: Option<[u8; 3]>,
    number_format: Option<&str>,
) -> Option<crate::xlsx::write::CellStyle> {
    use crate::xlsx::write::CellStyle;
    // A cell's number format is what makes a currency or date cell render as
    // one; dropping it turned every formatted number into a bare value.
    let with_fmt = |style: CellStyle| match number_format {
        Some(code) if !code.is_empty() => style.number_format_code(code),
        _ => style,
    };
    if is_header {
        let bg_hex = bg.map(rgb_to_hex).unwrap_or_else(|| "D3D3D3".to_string());
        Some(with_fmt(CellStyle::new().bold().background(bg_hex)))
    } else if let Some(c) = bg {
        Some(with_fmt(CellStyle::new().background(rgb_to_hex(c))))
    } else if number_format.is_some_and(|c| !c.is_empty()) {
        Some(with_fmt(CellStyle::new()))
    } else {
        None
    }
}

/// Convert a `TextBox`'s block content into the `(runs, ParaProps)`
/// pairs `SlideData::add_multi_paragraph_text_box` needs (issue #264).
///
/// Handles `Paragraph`/`Heading` (the common case — a real text box's
/// content is ordinary flowed paragraphs) as real paragraphs, each with
/// its own alignment/spacing carried through the same way
/// `emit_pptx_element`'s own `Element::Paragraph` arm does. Anything
/// else (a table, list, or nested text box inside a text box — rare)
/// falls back to one plain-text paragraph via `inline_to_text`-style
/// flattening, so it isn't silently dropped, at the cost of its own
/// rich formatting.
/// Convert `Paragraph`/`Heading` elements into the `(runs, ParaProps)`
/// pairs `SlideData::add_multi_paragraph_text_box` needs (issue #264).
/// Callers only pass paragraph-like content here — a nested `List`/
/// `Table`/`TextBox` inside a `TextBox` is emitted separately (see
/// `emit_pptx_element`'s own `Element::TextBox` arm), since it has no
/// positioned representation of its own and collapsing it to plain text
/// would both lose its structure and break write→parse idempotence for
/// the common "body placeholder wrapped in a TextBox" case.
fn textbox_content_to_pptx_paragraphs(
    content: &[Element],
) -> Vec<(Vec<crate::pptx::write::Run>, crate::pptx::write::ParaProps)> {
    let mut out = Vec::new();
    for elem in content {
        match elem {
            Element::Paragraph(p) => {
                let runs = inline_to_pptx_runs(&p.content);
                if !runs.is_empty() {
                    let props = crate::pptx::write::ParaProps {
                        alignment: p.alignment.clone(),
                        space_before_hundredths_pt: p.space_before_twips.map(|t| t * 5),
                    };
                    out.push((runs, props));
                }
            },
            Element::Heading(h) => {
                let runs = inline_to_pptx_runs(&h.content);
                if !runs.is_empty() {
                    let props = crate::pptx::write::ParaProps {
                        alignment: h.alignment.clone(),
                        ..Default::default()
                    };
                    out.push((runs, props));
                }
            },
            _ => {},
        }
    }
    out
}

fn inline_to_pptx_runs(content: &[InlineContent]) -> Vec<crate::pptx::write::Run> {
    use crate::pptx::write::Run;
    content
        .iter()
        .filter_map(|item| {
            match item {
                // A dropped line break silently joins the words on either
                // side of it ("LINEA" + "LINEB" renders as "LINEALINEB").
                InlineContent::LineBreak => return Some(Run::line_break()),
                InlineContent::Text(_) => {},
                _ => return None,
            }
            if let InlineContent::Text(span) = item {
                if span.text.is_empty() {
                    return None;
                }
                let mut run = Run::new(&span.text);
                if span.bold {
                    run = run.bold();
                }
                if span.italic {
                    run = run.italic();
                }
                if span.underline.is_some() {
                    run = run.underline();
                }
                if span.strikethrough {
                    run = run.strikethrough();
                }
                if let Some(half_pt) = span.font_size_half_pt {
                    run = run.font_size(half_pt as f64 / 2.0);
                }
                if let Some(c) = span.color {
                    run = run.color(rgb_to_hex(c));
                }
                if let Some(ref name) = span.font_name {
                    run = run.font(name.clone());
                }
                if let Some(ref url) = span.hyperlink {
                    run = run.hyperlink(url.clone());
                }
                Some(run)
            } else {
                None
            }
        })
        .collect()
}

#[cfg(test)]
mod xlsx_table_write_tests {
    use super::*;
    use std::io::Cursor;

    /// issue #340 — a row-spanning merge (`row_span > 1`) claims its
    /// column for every row it covers. The covered rows' own `cells`
    /// correctly exclude that column (the sparse, span-driven model
    /// every reader uses), but the writer used to compute each cell's
    /// output column by summing `col_span` in encounter order within
    /// that row's own (already-filtered) cell list — with no way to
    /// know a column was skipped, every cell after the gap silently
    /// shifted left, misplacing its content and hyperlink onto the
    /// wrong column instead of the one actually claimed by the merge.
    #[test]
    fn a_cell_after_a_row_spanning_merge_keeps_its_own_column_and_hyperlink() {
        let cell = |text: &str, url: &str, row_span: u32| TableCell {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: text.to_string(),
                    hyperlink: Some(url.to_string()),
                    ..Default::default()
                })],
                ..Default::default()
            })],
            col_span: 1,
            row_span,
            ..Default::default()
        };

        let table = Table {
            rows: vec![
                TableRow {
                    cells: vec![
                        cell("Anchor", "https://example.com/anchor", 3),
                        cell("B0", "https://example.com/b0", 1),
                    ],
                    ..Default::default()
                },
                // Column 0 is covered by the anchor's row_span=3 merge, so
                // this row's own `cells` correctly holds only column 1 —
                // the writer must place it at column 1, not column 0.
                TableRow { cells: vec![cell("B1", "https://example.com/b1", 1)], ..Default::default() },
                TableRow { cells: vec![cell("B2", "https://example.com/b2", 1)], ..Default::default() },
            ],
            ..Default::default()
        };

        let ir = DocumentIR {
            metadata: Metadata { format: DocumentFormat::Xlsx, ..Default::default() },
            sections: vec![Section { elements: vec![Element::Table(table)], ..Default::default() }],
            defined_names: Vec::new(),
        };

        let mut buf = Cursor::new(Vec::new());
        create_from_ir_to_writer(&ir, DocumentFormat::Xlsx, &mut buf).unwrap();
        buf.set_position(0);
        let doc = crate::Document::from_reader(buf, DocumentFormat::Xlsx).unwrap();
        let ir2 = doc.to_ir();

        let mut found: std::collections::HashMap<String, String> = std::collections::HashMap::new();
        let Element::Table(t) = &ir2.sections[0].elements[0] else {
            panic!("expected a table");
        };
        for row in &t.rows {
            for cell in &row.cells {
                for c in &cell.content {
                    if let Element::Paragraph(p) = c {
                        for inc in &p.content {
                            if let InlineContent::Text(span) = inc {
                                if let Some(url) = &span.hyperlink {
                                    found.insert(span.text.clone(), url.clone());
                                }
                            }
                        }
                    }
                }
            }
        }

        assert_eq!(found.get("Anchor").map(String::as_str), Some("https://example.com/anchor"));
        assert_eq!(found.get("B0").map(String::as_str), Some("https://example.com/b0"));
        assert_eq!(
            found.get("B1").map(String::as_str),
            Some("https://example.com/b1"),
            "B1's hyperlink must survive at its own column, not be lost to a left-shift: {found:?}"
        );
        assert_eq!(
            found.get("B2").map(String::as_str),
            Some("https://example.com/b2"),
            "B2's hyperlink must survive at its own column, not be lost to a left-shift: {found:?}"
        );
    }

    /// issue #344 — the XLSX reader turns every real cell comment into an
    /// `Element::Endnote` whose marker is `"{cell_ref}"` or
    /// `"{cell_ref} ({author})"` (see `convert_xlsx.rs`). The writer used
    /// to have nowhere to put endnote content in a spreadsheet and fell
    /// back to dumping each one's text as a spurious extra row with no
    /// cell reference at all, silently turning every commented cell into
    /// bogus sheet data instead of round-tripping the comment. It must
    /// now write a real `xl/comments*.xml` + `vmlDrawing*.vml` pair and
    /// leave the sheet's own grid untouched.
    #[test]
    fn a_cell_comment_round_trips_as_a_real_comment_not_an_extra_row() {
        let cell = |text: &str| TableCell {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan { text: text.to_string(), ..Default::default() })],
                ..Default::default()
            })],
            col_span: 1,
            row_span: 1,
            ..Default::default()
        };

        let table = Table {
            rows: vec![TableRow {
                cells: vec![cell("A1"), cell("B1")],
                ..Default::default()
            }],
            ..Default::default()
        };

        let comment = Element::Endnote(Note {
            id: 1,
            marker: Some("B1 (Jane Doe)".to_string()),
            author: None,
            content: vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(TextSpan {
                    text: "This needs review".to_string(),
                    ..Default::default()
                })],
                ..Default::default()
            })],
        });

        let ir = DocumentIR {
            metadata: Metadata { format: DocumentFormat::Xlsx, ..Default::default() },
            sections: vec![Section {
                elements: vec![Element::Table(table), comment],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };

        let mut buf = Cursor::new(Vec::new());
        create_from_ir_to_writer(&ir, DocumentFormat::Xlsx, &mut buf).unwrap();
        buf.set_position(0);

        // The written package must contain a real comments part and its
        // companion VML drawing, not just plain sheet data.
        let mut zip = zip::ZipArchive::new(buf.clone()).unwrap();
        let names: Vec<String> = (0..zip.len()).map(|i| zip.by_index(i).unwrap().name().to_string()).collect();
        assert!(
            names.iter().any(|n| n.starts_with("xl/comments") && n.ends_with(".xml")),
            "expected a comments part, got: {names:?}"
        );
        assert!(
            names.iter().any(|n| n.contains("vmlDrawing")),
            "expected a companion VML drawing part, got: {names:?}"
        );

        buf.set_position(0);
        let doc = crate::Document::from_reader(buf, DocumentFormat::Xlsx).unwrap();
        let ir2 = doc.to_ir();

        let Element::Table(t) = &ir2.sections[0].elements[0] else {
            panic!("expected a table as the first element");
        };
        assert_eq!(t.rows.len(), 1, "the comment must not appear as an extra sheet row");

        let has_comment_endnote = ir2.sections[0].elements.iter().any(|e| {
            matches!(e, Element::Endnote(n) if n.marker.as_deref() == Some("B1 (Jane Doe)"))
        });
        assert!(has_comment_endnote, "the comment must round-trip back onto cell B1: {:?}", ir2.sections[0].elements);
    }
}

#[cfg(test)]
mod docx_section_title_tests {
    use super::*;
    use std::io::Cursor;

    /// issue #338 — `convert_docx.rs` derives `section.title` from a
    /// heading using a narrower text extraction (only `InlineContent::
    /// Text`, silently dropping `LineBreak`) than `ir_to_docx`'s "is the
    /// title already present in the elements" check (`inline_to_text`,
    /// which turns `LineBreak` into `\n`). A heading containing a line
    /// break therefore never string-matched its own derived title, so
    /// the title-suppression check always failed and a second, spurious
    /// copy of the heading got written on every round trip. Both sides
    /// must now agree by construction — they share `ir::inline_to_text`.
    #[test]
    fn a_heading_with_a_line_break_is_not_duplicated_as_a_second_title() {
        let heading = Element::Heading(Heading {
            level: 1,
            content: vec![InlineContent::LineBreak, InlineContent::Text(TextSpan::plain("Individual Elements"))],
            ..Default::default()
        });
        let body = Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain("Body text."))],
            ..Default::default()
        });

        let ir = DocumentIR {
            metadata: Metadata { format: DocumentFormat::Docx, ..Default::default() },
            sections: vec![Section {
                title: Some(inline_to_text(&[
                    InlineContent::LineBreak,
                    InlineContent::Text(TextSpan::plain("Individual Elements")),
                ])),
                elements: vec![heading, body],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };

        let mut buf = Cursor::new(Vec::new());
        create_from_ir_to_writer(&ir, DocumentFormat::Docx, &mut buf).unwrap();
        buf.set_position(0);
        let doc = crate::Document::from_reader(buf, DocumentFormat::Docx).unwrap();
        let ir2 = doc.to_ir();

        let heading_count = ir2.sections[0]
            .elements
            .iter()
            .filter(|e| matches!(e, Element::Heading(h) if inline_to_text(&h.content).contains("Individual Elements")))
            .count();
        assert_eq!(
            heading_count, 1,
            "the heading must not be duplicated as a second title on write: {:?}",
            ir2.sections[0].elements
        );
    }
}
