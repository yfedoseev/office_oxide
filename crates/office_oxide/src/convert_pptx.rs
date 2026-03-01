use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn pptx_to_ir(doc: &pptx_oxide::PptxDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for slide in &doc.slides {
        let title = find_title_text(&slide.shapes);
        let mut elements = Vec::new();

        // Sort shapes spatially
        let mut shape_entries: Vec<(Option<&pptx_oxide::ShapePosition>, &pptx_oxide::Shape)> =
            Vec::new();
        collect_shape_entries(&slide.shapes, &mut shape_entries);
        shape_entries.sort_by(|a, b| spatial_cmp(a.0, b.0));

        for (_, shape) in &shape_entries {
            convert_shape(shape, &mut elements);
        }

        // Add notes as paragraphs at end
        if let Some(ref notes) = slide.notes {
            if !notes.is_empty() {
                elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan {
                        text: notes.clone(),
                        bold: false,
                        italic: false,
                        strikethrough: false,
                        hyperlink: None,
                    })],
                }));
            }
        }

        sections.push(Section {
            title: title.clone(),
            elements,
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Pptx,
            title,
        },
        sections,
    }
}

fn collect_shape_entries<'a>(
    shapes: &'a [pptx_oxide::Shape],
    entries: &mut Vec<(Option<&'a pptx_oxide::ShapePosition>, &'a pptx_oxide::Shape)>,
) {
    for shape in shapes {
        match shape {
            pptx_oxide::Shape::Group(grp) => {
                collect_shape_entries(&grp.children, entries);
            }
            pptx_oxide::Shape::AutoShape(auto) => {
                entries.push((auto.position.as_ref(), shape));
            }
            pptx_oxide::Shape::Picture(pic) => {
                entries.push((pic.position.as_ref(), shape));
            }
            pptx_oxide::Shape::GraphicFrame(gf) => {
                entries.push((gf.position.as_ref(), shape));
            }
            pptx_oxide::Shape::Connector(_) => {}
        }
    }
}

fn spatial_cmp(
    a: Option<&pptx_oxide::ShapePosition>,
    b: Option<&pptx_oxide::ShapePosition>,
) -> std::cmp::Ordering {
    match (a, b) {
        (Some(a), Some(b)) => a.y.cmp(&b.y).then(a.x.cmp(&b.x)),
        (Some(_), None) => std::cmp::Ordering::Less,
        (None, Some(_)) => std::cmp::Ordering::Greater,
        (None, None) => std::cmp::Ordering::Equal,
    }
}

fn is_title_placeholder(ph_type: Option<&str>) -> bool {
    matches!(ph_type, Some("title" | "ctrTitle"))
}

fn find_title_text(shapes: &[pptx_oxide::Shape]) -> Option<String> {
    for shape in shapes {
        match shape {
            pptx_oxide::Shape::AutoShape(auto) => {
                if auto
                    .placeholder
                    .as_ref()
                    .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref()))
                {
                    if let Some(ref tb) = auto.text_body {
                        let text = plain_text_from_body(tb);
                        if !text.is_empty() {
                            return Some(text);
                        }
                    }
                }
            }
            pptx_oxide::Shape::Group(grp) => {
                if let Some(title) = find_title_text(&grp.children) {
                    return Some(title);
                }
            }
            _ => {}
        }
    }
    None
}

fn plain_text_from_body(body: &pptx_oxide::TextBody) -> String {
    let mut parts = Vec::new();
    for para in &body.paragraphs {
        let mut text = String::new();
        for content in &para.content {
            match content {
                pptx_oxide::TextContent::Run(run) => text.push_str(&run.text),
                pptx_oxide::TextContent::LineBreak => text.push('\n'),
                pptx_oxide::TextContent::Field(field) => text.push_str(&field.text),
            }
        }
        parts.push(text);
    }
    parts.join("\n")
}

fn convert_shape(
    shape: &pptx_oxide::Shape,
    elements: &mut Vec<Element>,
) {
    match shape {
        pptx_oxide::Shape::AutoShape(auto) => {
            // Skip title placeholder — used as section title
            if auto
                .placeholder
                .as_ref()
                .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref()))
            {
                return;
            }

            if let Some(ref tb) = auto.text_body {
                convert_text_body(tb, elements);
            }
        }
        pptx_oxide::Shape::Picture(pic) => {
            elements.push(Element::Image(Image {
                alt_text: pic.alt_text.clone(),
            }));
        }
        pptx_oxide::Shape::Group(grp) => {
            for child in &grp.children {
                convert_shape(child, elements);
            }
        }
        pptx_oxide::Shape::GraphicFrame(gf) => {
            if let pptx_oxide::GraphicContent::Table(ref tbl) = gf.content {
                elements.push(convert_pptx_table(tbl));
            }
        }
        pptx_oxide::Shape::Connector(_) => {}
    }
}

fn convert_text_body(body: &pptx_oxide::TextBody, elements: &mut Vec<Element>) {
    // Check if any paragraph has level > 0 — treat as list
    let has_levels = body.paragraphs.iter().any(|p| p.level > 0);

    if has_levels {
        // Convert paragraphs with levels to list items
        let mut items = Vec::new();
        for para in &body.paragraphs {
            items.push((para.level as u8, convert_text_paragraph_inline(para)));
        }
        elements.push(Element::List(build_nested_list(false, &items, 0)));
    } else {
        for para in &body.paragraphs {
            let content = convert_text_paragraph_inline(para);
            if !content.is_empty() {
                elements.push(Element::Paragraph(Paragraph { content }));
            }
        }
    }
}

fn convert_text_paragraph_inline(
    para: &pptx_oxide::TextParagraph,
) -> Vec<InlineContent> {
    let mut content = Vec::new();
    for tc in &para.content {
        match tc {
            pptx_oxide::TextContent::Run(run) => {
                if !run.text.is_empty() {
                    let hyperlink = run.hyperlink.as_ref().and_then(|h| {
                        match &h.target {
                            pptx_oxide::HyperlinkTarget::External(url) => Some(url.clone()),
                            pptx_oxide::HyperlinkTarget::Internal(_) => None,
                        }
                    });
                    content.push(InlineContent::Text(TextSpan {
                        text: run.text.clone(),
                        bold: run.bold.unwrap_or(false),
                        italic: run.italic.unwrap_or(false),
                        strikethrough: run.strikethrough,
                        hyperlink,
                    }));
                }
            }
            pptx_oxide::TextContent::LineBreak => {
                content.push(InlineContent::LineBreak);
            }
            pptx_oxide::TextContent::Field(field) => {
                if !field.text.is_empty() {
                    content.push(InlineContent::Text(TextSpan {
                        text: field.text.clone(),
                        bold: false,
                        italic: false,
                        strikethrough: false,
                        hyperlink: None,
                    }));
                }
            }
        }
    }
    content
}

fn build_nested_list(
    ordered: bool,
    items: &[(u8, Vec<InlineContent>)],
    base_level: u8,
) -> List {
    let mut list_items = Vec::new();
    let mut idx = 0;

    while idx < items.len() {
        let (level, content) = &items[idx];
        if *level <= base_level {
            let nested_start = idx + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].0 > base_level {
                nested_end += 1;
            }
            let nested = if nested_end > nested_start {
                Some(build_nested_list(
                    ordered,
                    &items[nested_start..nested_end],
                    base_level + 1,
                ))
            } else {
                None
            };
            list_items.push(ListItem {
                content: content.clone(),
                nested,
            });
            idx = if nested_end > nested_start {
                nested_end
            } else {
                idx + 1
            };
        } else {
            list_items.push(ListItem {
                content: content.clone(),
                nested: None,
            });
            idx += 1;
        }
    }

    List {
        ordered,
        items: list_items,
    }
}

fn convert_pptx_table(table: &pptx_oxide::Table) -> Element {
    let mut ir_rows = Vec::new();

    for (row_idx, row) in table.rows.iter().enumerate() {
        let mut ir_cells = Vec::new();

        for cell in &row.cells {
            // Skip merged cells
            if cell.h_merge || cell.v_merge {
                continue;
            }

            let mut cell_elements = Vec::new();
            if let Some(ref tb) = cell.text_body {
                for para in &tb.paragraphs {
                    let content = convert_text_paragraph_inline(para);
                    if !content.is_empty() {
                        cell_elements.push(Element::Paragraph(Paragraph { content }));
                    }
                }
            }

            ir_cells.push(TableCell {
                content: cell_elements,
                col_span: cell.grid_span,
                row_span: cell.row_span,
            });
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header: row_idx == 0,
        });
    }

    Element::Table(Table { rows: ir_rows })
}
