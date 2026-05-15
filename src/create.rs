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
pub fn ir_to_docx(ir: &DocumentIR) -> crate::docx::write::DocxWriter {
    use crate::docx::write::{DocxWriter, IrParaProps, Run};

    let mut writer = DocxWriter::new();

    // Write metadata
    writer.set_metadata(&ir.metadata);

    for section in &ir.sections {
        // Section title becomes H1
        if let Some(ref title) = section.title {
            if !title.is_empty() {
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
            let level = h.level.clamp(1, 6);
            let runs: Vec<Run> = ir_inline_to_runs(&h.content);
            let props = IrParaProps {
                style: Some(format!("Heading{level}")),
                alignment: h.alignment.clone(),
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
            writer.add_footnote(n.id, &n.content);
        },
        Element::Endnote(n) => {
            writer.add_endnote(n.id, &n.content);
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
                    ..Default::default()
                });
            },
            InlineContent::EndnoteRef(r) => {
                runs.push(Run {
                    endnote_ref: Some(r.note_id),
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
    a.bold == b.bold
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
                    for row in &t.rows {
                        let mut col = 0usize;
                        for cell in &row.cells {
                            let text = cell_text(cell);
                            let data = text_to_cell_data(&text);
                            if let Some(style) =
                                xlsx_cell_style(row.is_header, cell.background_color)
                            {
                                sheet.set_cell_styled(row_cursor, col, data, style);
                            } else {
                                sheet.set_cell(row_cursor, col, data);
                            }
                            let cs = cell.col_span.max(1) as usize;
                            let rs = cell.row_span.max(1) as usize;
                            if cs > 1 || rs > 1 {
                                sheet.merge_cells(row_cursor, col, rs, cs);
                            }
                            col += cs;
                        }
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
                        for line in split_paragraph_for_xlsx(&text) {
                            sheet.set_cell_styled(
                                row_cursor,
                                0,
                                CellData::String(line),
                                style.clone(),
                            );
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
                _ => {},
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

    if let Some(ref title) = section.title {
        if !title.is_empty() {
            slide.set_title(title);
        }
    }

    for elem in &section.elements {
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
            let items: Vec<String> = l
                .items
                .iter()
                .map(|i| {
                    i.content
                        .iter()
                        .map(|e| match e {
                            Element::Paragraph(p) => inline_to_text(&p.content),
                            _ => String::new(),
                        })
                        .collect::<Vec<_>>()
                        .join(" ")
                })
                .collect();
            let item_refs: Vec<&str> = items.iter().map(|s| s.as_str()).collect();
            slide.add_bullet_list(&item_refs);
        },
        Element::Table(t) => {
            let text = t
                .rows
                .iter()
                .map(|row| {
                    row.cells
                        .iter()
                        .map(cell_text)
                        .collect::<Vec<_>>()
                        .join("\t")
                })
                .collect::<Vec<_>>()
                .join("\n");
            if !text.is_empty() {
                slide.add_text(&text);
            }
        },
        Element::Image(img) => {
            if let (Some(data), Some(fmt)) = (&img.data, &img.format) {
                let cx = img.display_width_emu.unwrap_or(3_000_000);
                let cy = img.display_height_emu.unwrap_or(2_000_000);
                slide.add_image(data.clone(), fmt.clone(), 0, 0, cx, cy);
            }
        },
        Element::CodeBlock(cb) => {
            let run = crate::pptx::write::Run::new(&cb.content).font("Courier New");
            slide.add_rich_text(&[run]);
        },
        _ => {},
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

fn inline_to_text(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => out.push_str(&span.text),
            InlineContent::LineBreak => out.push('\n'),
            InlineContent::FootnoteRef(_) | InlineContent::EndnoteRef(_) => {},
        }
    }
    out
}

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

fn xlsx_cell_style(is_header: bool, bg: Option<[u8; 3]>) -> Option<crate::xlsx::write::CellStyle> {
    use crate::xlsx::write::CellStyle;
    if is_header {
        let bg_hex = bg.map(rgb_to_hex).unwrap_or_else(|| "D3D3D3".to_string());
        Some(CellStyle::new().bold().background(bg_hex))
    } else {
        bg.map(|c| CellStyle::new().background(rgb_to_hex(c)))
    }
}

fn inline_to_pptx_runs(content: &[InlineContent]) -> Vec<crate::pptx::write::Run> {
    use crate::pptx::write::Run;
    content
        .iter()
        .filter_map(|item| {
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
                if let Some(half_pt) = span.font_size_half_pt {
                    run = run.font_size(half_pt as f64 / 2.0);
                }
                if let Some(c) = span.color {
                    run = run.color(rgb_to_hex(c));
                }
                if let Some(ref name) = span.font_name {
                    run = run.font(name.clone());
                }
                Some(run)
            } else {
                None
            }
        })
        .collect()
}
