//! Unified document creation from IR or markdown.

use std::io::{Seek, Write};
use std::path::Path;

use crate::format::DocumentFormat;
use crate::ir::*;
use crate::Result;

/// Create a document file from a `DocumentIR`.
pub fn create_from_ir(
    ir: &DocumentIR,
    format: DocumentFormat,
    path: impl AsRef<Path>,
) -> Result<()> {
    match format {
        #[cfg(feature = "docx")]
        DocumentFormat::Docx => {
            let writer = ir_to_docx(ir);
            writer.save(path)?;
        }
        #[cfg(feature = "xlsx")]
        DocumentFormat::Xlsx => {
            let writer = ir_to_xlsx(ir);
            writer.save(path)?;
        }
        #[cfg(feature = "pptx")]
        DocumentFormat::Pptx => {
            let writer = ir_to_pptx(ir);
            writer.save(path)?;
        }
        #[allow(unreachable_patterns)]
        _ => {
            return Err(crate::OfficeError::UnsupportedFormat(format!(
                "{format:?}"
            )))
        }
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
        #[cfg(feature = "docx")]
        DocumentFormat::Docx => {
            let w = ir_to_docx(ir);
            w.write_to(writer)?;
        }
        #[cfg(feature = "xlsx")]
        DocumentFormat::Xlsx => {
            let w = ir_to_xlsx(ir);
            w.write_to(writer)?;
        }
        #[cfg(feature = "pptx")]
        DocumentFormat::Pptx => {
            let w = ir_to_pptx(ir);
            w.write_to(writer)?;
        }
        #[allow(unreachable_patterns)]
        _ => {
            return Err(crate::OfficeError::UnsupportedFormat(format!(
                "{format:?}"
            )))
        }
    }
    Ok(())
}

/// Convert IR to a DOCX writer.
#[cfg(feature = "docx")]
fn ir_to_docx(ir: &DocumentIR) -> docx_oxide::write::DocxWriter {
    let mut writer = docx_oxide::write::DocxWriter::new();

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

#[cfg(feature = "docx")]
fn add_element_to_docx(writer: &mut docx_oxide::write::DocxWriter, elem: &Element) {
    match elem {
        Element::Heading(h) => {
            let text = inline_to_text(&h.content);
            writer.add_heading(&text, h.level);
        }
        Element::Paragraph(p) => {
            let text = inline_to_text(&p.content);
            if !text.is_empty() {
                writer.add_paragraph(&text);
            }
        }
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
        }
        Element::List(l) => {
            let items: Vec<String> = l.items.iter().map(|i| inline_to_text(&i.content)).collect();
            let item_refs: Vec<&str> = items.iter().map(String::as_str).collect();
            writer.add_list(&item_refs, l.ordered);
        }
        Element::ThematicBreak | Element::Image(_) => {}
    }
}

/// Convert IR to an XLSX writer.
#[cfg(feature = "xlsx")]
fn ir_to_xlsx(ir: &DocumentIR) -> xlsx_oxide::write::XlsxWriter {
    use xlsx_oxide::write::CellData;

    let mut writer = xlsx_oxide::write::XlsxWriter::new();

    for section in &ir.sections {
        let name = section.title.as_deref().unwrap_or("Sheet");
        let sheet = writer.add_sheet(name);

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
                }
                Element::Paragraph(p) => {
                    let text = inline_to_text(&p.content);
                    if !text.is_empty() {
                        sheet.add_row(vec![CellData::String(text)]);
                    }
                }
                Element::Heading(h) => {
                    let text = inline_to_text(&h.content);
                    if !text.is_empty() {
                        sheet.add_row(vec![CellData::String(text)]);
                    }
                }
                _ => {}
            }
        }
    }

    writer
}

/// Convert IR to a PPTX writer.
#[cfg(feature = "pptx")]
fn ir_to_pptx(ir: &DocumentIR) -> pptx_oxide::write::PptxWriter {
    let mut writer = pptx_oxide::write::PptxWriter::new();

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
                }
                Element::Paragraph(p) => {
                    let text = inline_to_text(&p.content);
                    if !text.is_empty() {
                        slide.add_text(&text);
                    }
                }
                Element::List(l) => {
                    let items: Vec<String> =
                        l.items.iter().map(|i| inline_to_text(&i.content)).collect();
                    let item_refs: Vec<&str> = items.iter().map(|s| s.as_str()).collect();
                    slide.add_bullet_list(&item_refs);
                }
                _ => {}
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
