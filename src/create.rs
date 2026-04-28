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

/// Convert IR to a DOCX writer.
fn ir_to_docx(ir: &DocumentIR) -> crate::docx::write::DocxWriter {
    let mut writer = crate::docx::write::DocxWriter::new();

    for section in &ir.sections {
        if let Some(ref title) = section.title {
            if !title.is_empty() {
                writer.add_heading(title, 1);
            }
        }
        for elem in &section.elements {
            add_element_to_docx(&mut writer, elem);
        }
    }

    writer
}

fn add_element_to_docx(writer: &mut crate::docx::write::DocxWriter, elem: &Element) {
    match elem {
        Element::Heading(h) => {
            let text = inline_to_text(&h.content);
            writer.add_heading(&text, h.level);
        },
        Element::Paragraph(p) => {
            let text = inline_to_text(&p.content);
            if !text.is_empty() {
                writer.add_paragraph(&text);
            }
        },
        Element::Table(t) => {
            let rows: Vec<Vec<String>> = t
                .rows
                .iter()
                .map(|r| {
                    r.cells
                        .iter()
                        .map(|c| {
                            c.content
                                .iter()
                                .map(|e| match e {
                                    Element::Paragraph(p) => inline_to_text(&p.content),
                                    _ => String::new(),
                                })
                                .collect::<Vec<_>>()
                                .join(" ")
                        })
                        .collect()
                })
                .collect();
            let row_refs: Vec<Vec<&str>> = rows
                .iter()
                .map(|r| r.iter().map(String::as_str).collect())
                .collect();
            writer.add_table(&row_refs);
        },
        Element::List(l) => {
            let items: Vec<String> = l.items.iter().map(|i| inline_to_text(&i.content)).collect();
            let item_refs: Vec<&str> = items.iter().map(String::as_str).collect();
            writer.add_list(&item_refs, l.ordered);
        },
        Element::ThematicBreak | Element::Image(_) => {},
    }
}

/// Convert IR to an XLSX writer.
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

/// Convert IR to a PPTX writer.
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
                    // If no title yet, use heading as title
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
                    let items: Vec<String> =
                        l.items.iter().map(|i| inline_to_text(&i.content)).collect();
                    let item_refs: Vec<&str> = items.iter().map(|s| s.as_str()).collect();
                    slide.add_bullet_list(&item_refs);
                },
                _ => {},
            }
        }
    }

    writer
}

/// Extract plain text from inline content.
fn inline_to_text(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => out.push_str(&span.text),
            InlineContent::LineBreak => out.push('\n'),
        }
    }
    out
}
