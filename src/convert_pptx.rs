use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn pptx_to_ir(doc: &crate::pptx::PptxDocument) -> DocumentIR {
    // Slide size sits at presentation level — every slide in the
    // deck shares it. EMU → twips is /635 (914400 EMU per inch,
    // 1440 twips per inch → 914400/1440 = 635).
    let page_setup = doc.presentation.slide_size.as_ref().map(|sz| PageSetup {
        width_twips: (sz.cx.max(0) / 635) as u32,
        height_twips: (sz.cy.max(0) / 635) as u32,
        landscape: sz.cx > sz.cy,
        ..Default::default()
    });

    let mut sections = Vec::new();

    for slide in doc.slides.iter() {
        let title_with_algn = find_title(&slide.shapes);
        let title = title_with_algn.as_ref().map(|(t, _)| t.clone());
        let title_alignment = title_with_algn.as_ref().and_then(|(_, a)| a.clone());
        let mut elements = Vec::new();

        // Lead each slide with the title placeholder text as a
        // heading so it has visible demarcation in the rendered
        // PDF/HTML output. When the slide has no title we used to
        // synthesise "Slide N" — that was useful for markdown anchors
        // but pure visual noise in paginated output, where every
        // slide already starts on its own page via the NextPage break.
        // Worse, the synthesised heading rendered as 20 pt bold and
        // contributed ~50 pt of fixed vertical overhead per section,
        // which inflated PDF→PPTX→PDF round-trip page counts.
        if let Some(ref t) = title {
            elements.push(Element::Heading(Heading {
                level: 2,
                content: vec![InlineContent::Text(TextSpan::plain(t.clone()))],
                alignment: title_alignment.clone(),
                ..Default::default()
            }));
        }

        // Sort shapes spatially
        let mut shape_entries: Vec<(Option<&crate::pptx::ShapePosition>, &crate::pptx::Shape)> =
            Vec::new();
        collect_shape_entries(&slide.shapes, &mut shape_entries);
        shape_entries.sort_by(|a, b| spatial_cmp(a.0, b.0));

        for (_, shape) in &shape_entries {
            convert_shape(shape, &mut elements);
        }

        // Propagate slide background colour to the section so the
        // PDF renderer can paint a full-slide rectangle before laying
        // down shapes.
        let background_rgb = slide.background_rgb;

        // Add notes as paragraphs at end
        if let Some(ref notes) = slide.notes {
            if !notes.is_empty() {
                elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(notes.clone()))],
                    ..Default::default()
                }));
            }
        }

        // Each PPTX slide is its own page when rendered to PDF or
        // any paginated format. Default `Continuous` would let two
        // slides share a page, which is wrong for slide content.
        let break_type = if sections.is_empty() {
            SectionBreakType::Continuous
        } else {
            SectionBreakType::NextPage
        };

        sections.push(Section {
            title: title.clone(),
            elements,
            break_type,
            page_setup: page_setup.clone(),
            background_rgb,
            ..Default::default()
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Pptx,
            title,
            ..Default::default()
        },
        sections,
    }
}

fn collect_shape_entries<'a>(
    shapes: &'a [crate::pptx::Shape],
    entries: &mut Vec<(Option<&'a crate::pptx::ShapePosition>, &'a crate::pptx::Shape)>,
) {
    for shape in shapes {
        match shape {
            crate::pptx::Shape::Group(grp) => {
                collect_shape_entries(&grp.children, entries);
            },
            crate::pptx::Shape::AutoShape(auto) => {
                entries.push((auto.position.as_ref(), shape));
            },
            crate::pptx::Shape::Picture(pic) => {
                entries.push((pic.position.as_ref(), shape));
            },
            crate::pptx::Shape::GraphicFrame(gf) => {
                entries.push((gf.position.as_ref(), shape));
            },
            crate::pptx::Shape::Connector(_) => {},
        }
    }
}

fn spatial_cmp(
    a: Option<&crate::pptx::ShapePosition>,
    b: Option<&crate::pptx::ShapePosition>,
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

/// Locate the title placeholder and return its text together with the
/// alignment of the first paragraph. Used by `pptx_to_ir` to seed both
/// `Section.title` and the synthesised level-2 Heading's alignment.
fn find_title(shapes: &[crate::pptx::Shape]) -> Option<(String, Option<ParagraphAlignment>)> {
    for shape in shapes {
        match shape {
            crate::pptx::Shape::AutoShape(auto)
                if auto
                    .placeholder
                    .as_ref()
                    .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref())) =>
            {
                if let Some(ref tb) = auto.text_body {
                    let text = plain_text_from_body(tb);
                    if !text.is_empty() {
                        let algn = tb.paragraphs.first().and_then(|p| p.alignment.clone());
                        return Some((text, algn));
                    }
                }
            },
            crate::pptx::Shape::Group(grp) => {
                if let Some(t) = find_title(&grp.children) {
                    return Some(t);
                }
            },
            _ => {},
        }
    }
    None
}

fn plain_text_from_body(body: &crate::pptx::TextBody) -> String {
    let mut parts = Vec::new();
    for para in &body.paragraphs {
        let mut text = String::new();
        for content in &para.content {
            match content {
                crate::pptx::TextContent::Run(run) => text.push_str(&run.text),
                crate::pptx::TextContent::LineBreak => text.push('\n'),
                crate::pptx::TextContent::Field(field) => text.push_str(&field.text),
            }
        }
        parts.push(text);
    }
    parts.join("\n")
}

fn convert_shape(shape: &crate::pptx::Shape, elements: &mut Vec<Element>) {
    match shape {
        crate::pptx::Shape::AutoShape(auto) => {
            // Skip title placeholder — used as section title
            if auto
                .placeholder
                .as_ref()
                .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref()))
            {
                return;
            }

            if let Some(ref tb) = auto.text_body {
                let mut inner = Vec::new();
                convert_text_body(tb, &mut inner);
                if inner.is_empty() {
                    return;
                }
                push_positional_textbox(elements, inner, auto.position.as_ref());
            }
        },
        crate::pptx::Shape::Picture(pic) => {
            // Carry the resolved media bytes through so the PDF renderer
            // (`render_pptx_textbox_content`) can paint the actual
            // picture at its shape rectangle. `embed_rid` is preserved
            // as alt-text fallback only when the relationship couldn't
            // be resolved — we still want a placeholder element so the
            // shape's position survives in plain-text / markdown output.
            let format = pic.format.as_deref().and_then(image_format_from_ext);
            let (display_w, display_h) = pic
                .position
                .as_ref()
                .map(|p| (Some(p.cx.max(0) as u64), Some(p.cy.max(0) as u64)))
                .unwrap_or((None, None));
            let img_el = Element::Image(Image {
                alt_text: pic.alt_text.clone(),
                data: pic.data.clone(),
                format,
                display_width_emu: display_w,
                display_height_emu: display_h,
                ..Default::default()
            });
            push_positional_textbox(elements, vec![img_el], pic.position.as_ref());
        },
        crate::pptx::Shape::Group(grp) => {
            for child in &grp.children {
                convert_shape(child, elements);
            }
        },
        crate::pptx::Shape::GraphicFrame(gf) => {
            if let crate::pptx::GraphicContent::Table(ref tbl) = gf.content {
                let table_el = convert_pptx_table(tbl);
                push_positional_textbox(elements, vec![table_el], gf.position.as_ref());
            }
        },
        crate::pptx::Shape::Connector(_) => {},
    }
}

/// Wrap a shape's converted IR content in a `TextBox` carrying its
/// absolute `(x, y, cx, cy)` EMU rectangle. The PPTX renderer uses
/// these coordinates to paint each shape at its source position
/// instead of flowing them as a single long page.
///
/// When the source shape has no `<a:xfrm>` (rare — placeholders that
/// inherit geometry from a slide layout), the inner content is pushed
/// as flow elements so plain-text / markdown rendering still sees it.
fn push_positional_textbox(
    elements: &mut Vec<Element>,
    content: Vec<Element>,
    position: Option<&crate::pptx::ShapePosition>,
) {
    // Wrap in `Element::TextBox` only when the source shape carried a
    // *real* `<a:xfrm>`. Placeholders that inherit geometry from the
    // slide layout parse as `ShapePosition { x: 0, y: 0, cx: 0, cy: 0 }`
    // — wrapping those in TextBox tells the renderer "place this 0×0
    // rectangle at (0, 0)" which collapses every paragraph onto the
    // top-left corner. Treat all-zeros as "no position" so the
    // content flows normally instead.
    let real_position = position.filter(|p| p.cx > 0 && p.cy > 0);
    if let Some(pos) = real_position {
        elements.push(Element::TextBox(TextBox {
            content,
            x_emu: Some(pos.x),
            y_emu: Some(pos.y),
            width_emu: Some(pos.cx.max(0) as u64),
            height_emu: Some(pos.cy.max(0) as u64),
            ..Default::default()
        }));
    } else {
        elements.extend(content);
    }
}

fn convert_text_body(body: &crate::pptx::TextBody, elements: &mut Vec<Element>) {
    // Check if any paragraph has level > 0 — treat as list
    let has_levels = body.paragraphs.iter().any(|p| p.level > 0);

    if has_levels {
        // Convert paragraphs with levels to list items
        let mut items = Vec::new();
        for para in &body.paragraphs {
            items.push((para.level as u8, convert_text_paragraph_inline(para)));
        }
        elements.push(Element::List(crate::ir::build_nested_list(false, &items, 0)));
    } else {
        for para in &body.paragraphs {
            let content = convert_text_paragraph_inline(para);
            // Honour space_before from PPTX so spacer paragraphs
            // emitted by pdf_to_ir round-trip with their full vertical
            // gap. Convert hundredths-of-pt → twips: hundredths * 0.2
            // (1pt = 20 twips, so pt*100 → twips = (pt*100)/5).
            let space_before_twips = para.space_before_hundredths_pt.map(|h| h.div_ceil(5));
            // Empty paragraphs serve as vertical spacers — keep them
            // in the IR even when content is empty so the renderer
            // can advance the cursor by the requested amount.
            if !content.is_empty() || space_before_twips.is_some() {
                elements.push(Element::Paragraph(Paragraph {
                    content,
                    alignment: para.alignment.clone(),
                    space_before_twips,
                    ..Default::default()
                }));
            }
        }
    }
}

fn convert_text_paragraph_inline(para: &crate::pptx::TextParagraph) -> Vec<InlineContent> {
    let mut content = Vec::new();
    for tc in &para.content {
        match tc {
            crate::pptx::TextContent::Run(run) => {
                if !run.text.is_empty() {
                    let hyperlink = run.hyperlink.as_ref().and_then(|h| match &h.target {
                        crate::pptx::HyperlinkTarget::External(url) => Some(url.clone()),
                        crate::pptx::HyperlinkTarget::Internal(_) => None,
                    });
                    let font_size_half_pt = run.font_size_hundredths_pt.map(|hp| {
                        crate::core::units::HalfPoint::from_drawingml_sz(hp)
                            .0
                            .max(1)
                    });
                    content.push(InlineContent::Text(TextSpan {
                        text: run.text.clone(),
                        bold: run.bold.unwrap_or(false),
                        italic: run.italic.unwrap_or(false),
                        strikethrough: run.strikethrough,
                        hyperlink,
                        font_size_half_pt,
                        color: run.color_rgb,
                        ..Default::default()
                    }));
                }
            },
            crate::pptx::TextContent::LineBreak => {
                content.push(InlineContent::LineBreak);
            },
            crate::pptx::TextContent::Field(field) => {
                if !field.text.is_empty() {
                    content.push(InlineContent::Text(TextSpan::plain(field.text.clone())));
                }
            },
        }
    }
    content
}

fn convert_pptx_table(table: &crate::pptx::Table) -> Element {
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
                        cell_elements.push(Element::Paragraph(Paragraph {
                            content,
                            alignment: para.alignment.clone(),
                            ..Default::default()
                        }));
                    }
                }
            }

            ir_cells.push(TableCell {
                content: cell_elements,
                col_span: cell.grid_span,
                row_span: cell.row_span,
                ..Default::default()
            });
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header: row_idx == 0,
            ..Default::default()
        });
    }

    Element::Table(Table {
        rows: ir_rows,
        ..Default::default()
    })
}

/// Map a lowercase file extension (`"png"`, `"jpeg"`, `"emf"`, …) to
/// the matching `ImageFormat` variant. Used by `convert_shape` when
/// converting a parsed PPTX `<p:pic>` whose underlying media part the
/// PPTX reader resolved into bytes + extension.
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
