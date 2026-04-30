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

fn ir_to_docx(ir: &DocumentIR) -> crate::docx::write::DocxWriter {
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
            // Emit a horizontal rule as a paragraph with bottom border
            let props = IrParaProps::default();
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
    runs
}

// ---------------------------------------------------------------------------
// XLSX conversion
// ---------------------------------------------------------------------------

fn ir_to_xlsx(ir: &DocumentIR) -> crate::xlsx::write::XlsxWriter {
    use crate::xlsx::write::CellData;

    let mut writer = crate::xlsx::write::XlsxWriter::new();

    for section in &ir.sections {
        let name = section.title.as_deref().unwrap_or("Sheet");
        let mut sheet = writer.add_sheet(name);

        for elem in &section.elements {
            match elem {
                Element::Table(t) => {
                    for row in &t.rows {
                        let cells: Vec<CellData> = row
                            .cells
                            .iter()
                            .map(|c| {
                                let text: String = c
                                    .content
                                    .iter()
                                    .map(|e| match e {
                                        Element::Paragraph(p) => inline_to_text(&p.content),
                                        _ => String::new(),
                                    })
                                    .collect::<Vec<_>>()
                                    .join(" ");
                                if text.is_empty() {
                                    CellData::Empty
                                } else if let Ok(n) = text.parse::<f64>() {
                                    CellData::Number(n)
                                } else {
                                    CellData::String(text)
                                }
                            })
                            .collect();
                        sheet.add_row(cells);
                    }
                },
                Element::Paragraph(p) => {
                    let text = inline_to_text(&p.content);
                    if !text.is_empty() {
                        sheet.add_row(vec![CellData::String(text)]);
                    }
                },
                Element::Heading(h) => {
                    let text = inline_to_text(&h.content);
                    if !text.is_empty() {
                        sheet.add_row(vec![CellData::String(text)]);
                    }
                },
                _ => {},
            }
        }
    }

    writer
}

// ---------------------------------------------------------------------------
// PPTX conversion
// ---------------------------------------------------------------------------

fn ir_to_pptx(ir: &DocumentIR) -> crate::pptx::write::PptxWriter {
    let mut writer = crate::pptx::write::PptxWriter::new();

    for section in &ir.sections {
        let slide = writer.add_slide();

        if let Some(ref title) = section.title {
            if !title.is_empty() {
                slide.set_title(title);
            }
        }

        for elem in &section.elements {
            match elem {
                Element::Heading(h) => {
                    let text = inline_to_text(&h.content);
                    if slide.title.is_none() {
                        slide.set_title(&text);
                    } else {
                        slide.add_text(&text);
                    }
                },
                Element::Paragraph(p) => {
                    let text = inline_to_text(&p.content);
                    if !text.is_empty() {
                        slide.add_text(&text);
                    }
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
                _ => {},
            }
        }
    }

    writer
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
