use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn docx_to_ir(doc: &crate::docx::DocxDocument) -> DocumentIR {
    // Build per-section block-element windows from `body.section_breaks`.
    // Each break index is the exclusive end of one section. Trailing
    // elements after the last break go into a final section described
    // by the body-level `<w:sectPr>`.
    let breaks = &doc.body.section_breaks;
    let total = doc.body.elements.len();

    let mut windows: Vec<(usize, usize)> = Vec::new();
    let mut prev = 0;
    for &b in breaks {
        let end = b.min(total);
        if end > prev {
            windows.push((prev, end));
        }
        prev = end;
    }
    if prev < total || windows.is_empty() {
        windows.push((prev, total));
    }

    // Bring page-level headers and footers into the IR. Without this any
    // downstream renderer (PDF, search, plain-text) loses non-body content
    // like "My header" / "My footer" / page numbers / running titles. The
    // split between header and footer uses the section ref counts (same
    // approach as `to_markdown`).
    let n_header_refs: usize = doc.sections.iter().map(|s| s.header_refs.len()).sum();
    let mut header_blocks: Vec<Element> = Vec::new();
    let mut footer_blocks: Vec<Element> = Vec::new();
    for (idx, hf) in doc.headers_footers.iter().enumerate() {
        let mut tmp: Vec<Element> = Vec::new();
        convert_block_elements(&hf.content, &mut tmp, doc);
        if tmp
            .iter()
            .all(|e| matches!(e, Element::Paragraph(p) if p.content.is_empty()))
        {
            continue;
        }
        if idx < n_header_refs {
            header_blocks.extend(tmp);
        } else {
            footer_blocks.extend(tmp);
        }
    }
    let header = if header_blocks.is_empty() {
        None
    } else {
        Some(HeaderFooter {
            content: header_blocks,
        })
    };
    let footer = if footer_blocks.is_empty() {
        None
    } else {
        Some(HeaderFooter {
            content: footer_blocks,
        })
    };

    let mut ir_sections: Vec<Section> = Vec::with_capacity(windows.len());
    let mut doc_title: Option<String> = None;

    for (idx, (start, end)) in windows.iter().copied().enumerate() {
        let mut elements = Vec::new();
        convert_block_elements(&doc.body.elements[start..end], &mut elements, doc);

        let title = elements.iter().find_map(|e| {
            if let Element::Heading(h) = e {
                Some(
                    h.content
                        .iter()
                        .filter_map(|c| match c {
                            InlineContent::Text(span) => Some(span.text.as_str()),
                            _ => None,
                        })
                        .collect::<String>(),
                )
            } else {
                None
            }
        });
        if doc_title.is_none() {
            doc_title = title.clone();
        }

        let page_setup = doc.sections.get(idx).map(section_props_to_page_setup);
        // Propagate the multi-column layout out of the source DOCX so
        // the IR carries `Section.columns` for the renderer. Without
        // this, a PDF→DOCX→PDF round-trip of a 2-column source paper
        // (arxiv preprints etc.) collapsed back to a single column on
        // read because the column count was dropped at this hop.
        let columns = doc
            .sections
            .get(idx)
            .and_then(|sp| sp.columns)
            .filter(|n| *n >= 2)
            .map(|n| ColumnLayout {
                count: n,
                ..Default::default()
            });

        let break_type = if idx == 0 {
            SectionBreakType::Continuous
        } else {
            SectionBreakType::NextPage
        };

        ir_sections.push(Section {
            title,
            elements,
            page_setup,
            break_type,
            columns,
            header: header.clone(),
            footer: footer.clone(),
            ..Default::default()
        });
    }

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            title: doc_title,
            ..Default::default()
        },
        sections: ir_sections,
    }
}

fn section_props_to_page_setup(sp: &crate::docx::SectionProperties) -> PageSetup {
    let mut ps = PageSetup::default();
    if let Some(size) = &sp.page_size {
        ps.width_twips = size.width.0.max(0) as u32;
        ps.height_twips = size.height.0.max(0) as u32;
        if let Some(crate::docx::PageOrientation::Landscape) = size.orient {
            ps.landscape = true;
        }
    }
    if let Some(m) = &sp.margins {
        ps.margin_top_twips = m.top.0.max(0) as u32;
        ps.margin_bottom_twips = m.bottom.0.max(0) as u32;
        ps.margin_left_twips = m.left.0.max(0) as u32;
        ps.margin_right_twips = m.right.0.max(0) as u32;
        if let Some(h) = m.header {
            ps.header_distance_twips = h.0.max(0) as u32;
        }
        if let Some(f) = m.footer {
            ps.footer_distance_twips = f.0.max(0) as u32;
        }
    }
    ps
}

fn convert_block_elements(
    blocks: &[crate::docx::BlockElement],
    elements: &mut Vec<Element>,
    doc: &crate::docx::DocxDocument,
) {
    let mut i = 0;
    while i < blocks.len() {
        match &blocks[i] {
            crate::docx::BlockElement::Paragraph(p) => {
                // Check if this is a list item — group consecutive list paragraphs
                if let Some(nr) = p
                    .properties
                    .as_ref()
                    .and_then(|pp| pp.numbering_ref.as_ref())
                {
                    let list_element = convert_list_group(blocks, &mut i, nr.num_id, doc);
                    elements.push(list_element);
                    continue;
                }

                // Check for heading
                let heading_level = resolve_heading_level(p, doc);
                let alignment = paragraph_alignment(p);

                // Detect "horizontal rule" encoding: empty paragraph
                // with a single bottom border. pdf_to_ir round-trips
                // ThematicBreak through DOCX as exactly this shape;
                // recover it here so the renderer draws a rule.
                let inline = convert_paragraph_inline(p, doc);
                let is_empty_para = inline.iter().all(|ic| {
                    matches!(ic,
                        crate::ir::InlineContent::Text(s) if s.text.is_empty()
                    )
                });
                let has_bottom_border =
                    p.properties.as_ref().is_some_and(|pp| pp.has_bottom_border);
                if is_empty_para && has_bottom_border {
                    elements.push(Element::ThematicBreak);
                    i += 1;
                    continue;
                }

                if let Some(level) = heading_level {
                    elements.push(Element::Heading(Heading {
                        level: (level + 1).min(6),
                        content: convert_paragraph_inline(p, doc),
                        frame_position: paragraph_frame_position(p),
                        alignment,
                    }));
                } else {
                    // Check for page break in runs
                    let (before_break, has_break) = split_at_page_break(p, doc);
                    let frame_pos = paragraph_frame_position(p);
                    if !before_break.is_empty() || !has_break {
                        elements.push(Element::Paragraph(Paragraph {
                            content: if before_break.is_empty() && !has_break {
                                convert_paragraph_inline(p, doc)
                            } else {
                                before_break
                            },
                            frame_position: frame_pos,
                            alignment,
                            ..Default::default()
                        }));
                    }
                    if has_break {
                        elements.push(Element::ThematicBreak);
                    }
                }
                // Promote any floating drawings (anchored images, vector
                // shapes) embedded in this paragraph to paragraph-sibling
                // IR elements so the positional renderer can lay them out
                // alongside the text frame.
                collect_paragraph_floats(p, doc, elements);
                // Promote inline drawings (`<wp:inline>` wrapper) to
                // paragraph-sibling Image elements as well. Without this
                // every embedded raster image (e.g. logos, figures, the
                // CFR federal seal) lost its bytes on the way through
                // the IR — the inline-content model has no Image
                // variant, so hoisting to a sibling Element is the
                // only way to carry the bitmap forward.
                collect_paragraph_inline_images(p, doc, elements);
                i += 1;
            },
            crate::docx::BlockElement::Table(t) => {
                elements.push(convert_table(t, doc));
                i += 1;
            },
        }
    }
}

/// Pull `<w:framePr>` data out of a paragraph's properties into the IR
/// position type. Returns `None` if the paragraph isn't absolutely
/// positioned (the common case).
/// Walk a paragraph's runs and emit one IR `Element` for every
/// floating (anchored) drawing — both raster pictures and vector
/// `<wps:wsp>` shapes. Inline drawings are left for the inline-content
/// path. Promoting floats to paragraph siblings keeps the positional
/// renderer simple: it can iterate a flat element list and place each
/// one at its absolute coordinates.
/// Walk a paragraph's runs and emit one IR `Element::Image` for every
/// inline drawing (`<wp:inline>` wrapper). Counterpart to
/// `collect_paragraph_floats` which handles `<wp:anchor>`-anchored
/// drawings. The IR's `InlineContent` enum has no Image variant so
/// inline drawings can't ride along with the rest of a paragraph's
/// runs; instead we hoist them as paragraph-sibling Element::Image
/// nodes right after the surrounding text paragraph.
fn collect_paragraph_inline_images(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
    out: &mut Vec<Element>,
) {
    for pc in &p.content {
        let runs: &[crate::docx::Run] = match pc {
            crate::docx::ParagraphContent::Run(r) => std::slice::from_ref(r),
            crate::docx::ParagraphContent::Hyperlink(hl) => &hl.runs,
        };
        for run in runs {
            for rc in &run.content {
                if let crate::docx::RunContent::Drawing(d) = rc {
                    if !d.inline {
                        continue;
                    }
                    if d.relationship_id.is_empty() {
                        continue;
                    }
                    let (data, ext) = match doc.images.get(&d.relationship_id).cloned() {
                        Some(v) => v,
                        None => continue,
                    };
                    let format =
                        ext.as_deref()
                            .and_then(|e| match e.to_ascii_lowercase().as_str() {
                                "png" => Some(ImageFormat::Png),
                                "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
                                "gif" => Some(ImageFormat::Gif),
                                _ => None,
                            });
                    out.push(Element::Image(Image {
                        alt_text: d.description.clone(),
                        data: Some(data),
                        format,
                        display_width_emu: Some(d.width.0.max(0) as u64),
                        display_height_emu: Some(d.height.0.max(0) as u64),
                        positioning: ImagePositioning::Inline,
                        ..Default::default()
                    }));
                }
            }
        }
    }
}

fn collect_paragraph_floats(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
    out: &mut Vec<Element>,
) {
    for pc in &p.content {
        let runs: &[crate::docx::Run] = match pc {
            crate::docx::ParagraphContent::Run(r) => std::slice::from_ref(r),
            crate::docx::ParagraphContent::Hyperlink(hl) => &hl.runs,
        };
        for run in runs {
            for rc in &run.content {
                if let crate::docx::RunContent::Drawing(d) = rc {
                    if d.inline {
                        continue;
                    }
                    if let Some(el) = drawing_to_float_element(d, doc) {
                        out.push(el);
                    }
                }
            }
        }
    }
}

fn drawing_to_float_element(
    d: &crate::docx::DrawingInfo,
    doc: &crate::docx::DocxDocument,
) -> Option<Element> {
    use crate::docx::{AnchorFrame, ShapeKind};

    let pos = d.anchor_position?;
    let to_ir_anchor = |f: AnchorFrame| match f {
        AnchorFrame::Page => FloatAnchor::Page,
        AnchorFrame::Margin => FloatAnchor::Margin,
        AnchorFrame::Column => FloatAnchor::Column,
        AnchorFrame::Paragraph => FloatAnchor::Paragraph,
        AnchorFrame::Line | AnchorFrame::Character => FloatAnchor::Page,
    };
    let h_anchor = to_ir_anchor(pos.h_relative_from);
    let v_anchor = to_ir_anchor(pos.v_relative_from);
    let width_emu = d.width.0.max(0) as u64;
    let height_emu = d.height.0.max(0) as u64;

    // Vector shape takes precedence: a `<wps:wsp>` with `prstGeom`
    // never carries a `<a:blip>`, so the relationship_id is empty.
    if let Some(shape) = &d.shape {
        let kind = match shape.kind {
            ShapeKind::Line => ShapeGeom::Line,
            ShapeKind::Rect => ShapeGeom::Rect,
        };
        return Some(Element::Shape(Shape {
            kind,
            x_emu: pos.x_emu,
            y_emu: pos.y_emu,
            width_emu,
            height_emu,
            h_anchor,
            v_anchor,
            stroke_rgb: shape.stroke_rgb.map(|(r, g, b)| [r, g, b]),
            fill_rgb: shape.fill_rgb.map(|(r, g, b)| [r, g, b]),
            stroke_w_emu: shape.stroke_w_emu,
        }));
    }

    if d.relationship_id.is_empty() {
        return None;
    }
    let (data, ext) = doc.images.get(&d.relationship_id).cloned()?;
    let format = ext.as_deref().and_then(|e| match e {
        "png" => Some(ImageFormat::Png),
        "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
        _ => None,
    });
    Some(Element::Image(Image {
        alt_text: d.description.clone(),
        data: Some(data),
        format,
        display_width_emu: Some(width_emu),
        display_height_emu: Some(height_emu),
        positioning: ImagePositioning::Floating(FloatingImage {
            x_emu: pos.x_emu,
            y_emu: pos.y_emu,
            width_emu,
            height_emu,
            h_anchor,
            v_anchor,
            text_wrap: TextWrap::default(),
            allow_overlap: true,
        }),
        ..Default::default()
    }))
}

/// Translate a paragraph's `<w:jc>` justification into the IR's
/// `ParagraphAlignment`. `Left` (and `Both`/`Distribute`) collapse
/// to `None` so the renderer uses default left-alignment without
/// emitting an explicit override.
fn paragraph_alignment(p: &crate::docx::Paragraph) -> Option<ParagraphAlignment> {
    let jc = p
        .properties
        .as_ref()
        .and_then(|pp| pp.justification.as_ref())?;
    match jc {
        crate::docx::Justification::Center => Some(ParagraphAlignment::Center),
        crate::docx::Justification::Right => Some(ParagraphAlignment::Right),
        crate::docx::Justification::Both => Some(ParagraphAlignment::Justify),
        crate::docx::Justification::Distribute => Some(ParagraphAlignment::Distribute),
        crate::docx::Justification::Left => None,
    }
}

fn paragraph_frame_position(p: &crate::docx::Paragraph) -> Option<FramePosition> {
    p.properties.as_ref().and_then(|props| {
        props.frame_position.as_ref().map(|f| FramePosition {
            x_twips: f.x_twips,
            y_twips: f.y_twips,
            width_twips: f.width_twips,
            height_twips: f.height_twips,
        })
    })
}

fn resolve_heading_level(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> Option<u8> {
    let props = p.properties.as_ref()?;
    // Direct outline level
    if let Some(lvl) = props.outline_level {
        return Some(lvl);
    }
    // Resolve via stylesheet
    let style_id = props.style_id.as_ref()?;
    let styles = doc.styles.as_ref()?;
    styles.resolve_outline_level(style_id)
}

fn convert_paragraph_inline(
    p: &crate::docx::Paragraph,
    _doc: &crate::docx::DocxDocument,
) -> Vec<InlineContent> {
    let mut content = Vec::new();
    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                convert_run(run, None, &mut content);
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                let url = match &hl.target {
                    crate::docx::HyperlinkTarget::External(url) => Some(url.clone()),
                    crate::docx::HyperlinkTarget::Internal(_) => None,
                };
                for run in &hl.runs {
                    convert_run(run, url.as_deref(), &mut content);
                }
            },
        }
    }
    content
}

fn convert_run(
    run: &crate::docx::Run,
    hyperlink_url: Option<&str>,
    content: &mut Vec<InlineContent>,
) {
    let bold = run
        .properties
        .as_ref()
        .and_then(|rp| rp.bold)
        .unwrap_or(false);
    let italic = run
        .properties
        .as_ref()
        .and_then(|rp| rp.italic)
        .unwrap_or(false);
    let strike = run
        .properties
        .as_ref()
        .and_then(|rp| rp.strike.or(rp.dstrike))
        .unwrap_or(false);
    // `<w:sz w:val="N"/>` is already in half-points; IR uses the
    // same encoding. See `crate::core::units::HalfPoint::from_word_sz`
    // for the cross-format invariant (also: PPTX hundredths-pt,
    // XLSX points-as-f32 must convert here).
    let font_size_half_pt = run.properties.as_ref().and_then(|rp| {
        rp.font_size
            .map(|hp| crate::core::units::HalfPoint::from_word_sz(hp.0).0)
    });
    // `<w:rFonts w:ascii="...">` carries the run's face name. Without
    // forwarding it onto `TextSpan.font_name`, the IR→PDF renderer
    // falls back to the page builder's default font (Helvetica) and
    // every PDF→DOCX→PDF round-trip loses every typeface — even when
    // the DOCX writer correctly embedded the source-PDF font program
    // under `word/fonts/`.
    let font_name = run.properties.as_ref().and_then(|rp| rp.font_name.clone());
    // Propagate `<w:color w:val="RRGGBB"/>` so PDF→DOCX→PDF round-trips
    // preserve coloured text (red "0" in `pdfs_pdfium/text_color.pdf`
    // and the like). Theme / system / auto colours fall through to
    // the renderer default for now — resolving them properly needs the
    // document's `theme.xml`, which the current convert path doesn't
    // thread in.
    let text_color = run
        .properties
        .as_ref()
        .and_then(|rp| rp.color.as_ref())
        .and_then(|c| match c {
            crate::core::theme::ColorRef::Rgb(rgb) => Some(rgb.0),
            _ => None,
        });

    for rc in &run.content {
        match rc {
            crate::docx::RunContent::Text(text) => {
                content.push(InlineContent::Text(TextSpan {
                    text: text.clone(),
                    bold,
                    italic,
                    strikethrough: strike,
                    hyperlink: hyperlink_url.map(|s| s.to_string()),
                    font_size_half_pt,
                    font_name: font_name.clone(),
                    color: text_color,
                    ..Default::default()
                }));
            },
            crate::docx::RunContent::Break(crate::docx::BreakType::Line) => {
                content.push(InlineContent::LineBreak);
            },
            crate::docx::RunContent::Break(
                crate::docx::BreakType::Page | crate::docx::BreakType::Column,
            ) => {
                // Page/column breaks handled at paragraph level
            },
            crate::docx::RunContent::Tab => {
                content.push(InlineContent::Text(TextSpan::plain("\t")));
            },
            crate::docx::RunContent::Drawing(drawing) => {
                // Inline drawings handled at the paragraph level via
                // `collect_paragraph_inline_images`. The inline-content
                // model has no Image variant; hoisting here would
                // require splitting paragraphs around each drawing,
                // which loses spans. Just record alt text so the
                // run's surrounding text doesn't lose semantic continuity.
                if let Some(alt) = drawing.description.clone() {
                    if !alt.is_empty() {
                        content.push(InlineContent::Text(TextSpan::plain(alt)));
                    }
                }
            },
        }
    }
}

fn split_at_page_break(
    p: &crate::docx::Paragraph,
    _doc: &crate::docx::DocxDocument,
) -> (Vec<InlineContent>, bool) {
    let mut content = Vec::new();
    let mut has_break = false;

    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                for rc in &run.content {
                    if matches!(rc, crate::docx::RunContent::Break(crate::docx::BreakType::Page)) {
                        has_break = true;
                    }
                }
                if !has_break {
                    convert_run(run, None, &mut content);
                }
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                if !has_break {
                    let url = match &hl.target {
                        crate::docx::HyperlinkTarget::External(url) => Some(url.clone()),
                        crate::docx::HyperlinkTarget::Internal(_) => None,
                    };
                    for run in &hl.runs {
                        convert_run(run, url.as_deref(), &mut content);
                    }
                }
            },
        }
    }
    (content, has_break)
}

// ---------------------------------------------------------------------------
// List conversion
// ---------------------------------------------------------------------------

fn convert_list_group(
    blocks: &[crate::docx::BlockElement],
    i: &mut usize,
    num_id: u32,
    doc: &crate::docx::DocxDocument,
) -> Element {
    let mut items = Vec::new();
    let mut is_ordered = false;

    while *i < blocks.len() {
        if let crate::docx::BlockElement::Paragraph(p) = &blocks[*i] {
            if let Some(nr) = p
                .properties
                .as_ref()
                .and_then(|pp| pp.numbering_ref.as_ref())
            {
                if nr.num_id != num_id {
                    break;
                }

                // Determine ordered/unordered from numbering format
                if let Some(numbering) = doc.numbering.as_ref() {
                    if let Some(level) = numbering.resolve_level(nr.num_id, nr.ilvl) {
                        is_ordered = !matches!(
                            level.format,
                            crate::docx::NumberFormat::Bullet | crate::docx::NumberFormat::None
                        );
                    }
                }

                items.push((nr.ilvl, convert_paragraph_inline(p, doc)));
                *i += 1;
                continue;
            }
        }
        break;
    }

    // Build nested list structure from flat (ilvl, content) pairs
    Element::List(crate::ir::build_nested_list(is_ordered, &items, 0))
}

// ---------------------------------------------------------------------------
// Table conversion
// ---------------------------------------------------------------------------

fn convert_table(table: &crate::docx::Table, doc: &crate::docx::DocxDocument) -> Element {
    // First pass: compute row_span from vMerge patterns
    let num_rows = table.rows.len();
    let num_cols = table
        .rows
        .iter()
        .map(|r| {
            r.cells
                .iter()
                .map(|c| c.properties.as_ref().and_then(|p| p.grid_span).unwrap_or(1) as usize)
                .sum::<usize>()
        })
        .max()
        .unwrap_or(0);

    // Build a grid of (is_continue, row_span) for vMerge tracking
    let mut row_spans: Vec<Vec<u32>> = vec![vec![1; num_cols]; num_rows];

    // Track vMerge: for each column, walk down from each Restart to count Continue cells
    for col in 0..num_cols {
        let mut row = 0;
        while row < num_rows {
            let cell = get_cell_at_grid_col(&table.rows[row], col);
            if let Some(cell) = cell {
                let vmerge = cell.properties.as_ref().and_then(|p| p.vertical_merge);
                if matches!(vmerge, Some(crate::docx::table::MergeType::Restart)) {
                    // Count continuation cells below
                    let mut span = 1u32;
                    let mut next = row + 1;
                    while next < num_rows {
                        let next_cell = get_cell_at_grid_col(&table.rows[next], col);
                        if let Some(nc) = next_cell {
                            if matches!(
                                nc.properties.as_ref().and_then(|p| p.vertical_merge),
                                Some(crate::docx::table::MergeType::Continue)
                            ) {
                                span += 1;
                                next += 1;
                                continue;
                            }
                        }
                        break;
                    }
                    if let Some(cell_span) = row_spans[row].get_mut(col) {
                        *cell_span = span;
                    }
                }
            }
            row += 1;
        }
    }

    let mut ir_rows = Vec::new();
    for (row_idx, row) in table.rows.iter().enumerate() {
        let is_header = row.properties.as_ref().is_some_and(|p| p.is_header);

        let mut ir_cells = Vec::new();
        let mut grid_col = 0;

        for cell in &row.cells {
            let col_span = cell
                .properties
                .as_ref()
                .and_then(|p| p.grid_span)
                .unwrap_or(1);

            // Skip vMerge continue cells
            let is_continue = cell
                .properties
                .as_ref()
                .and_then(|p| p.vertical_merge)
                .is_some_and(|m| matches!(m, crate::docx::table::MergeType::Continue));

            if is_continue {
                grid_col += col_span as usize;
                continue;
            }

            let row_span = if grid_col < num_cols {
                row_spans[row_idx][grid_col]
            } else {
                1
            };

            let mut cell_elements = Vec::new();
            convert_block_elements(&cell.content, &mut cell_elements, doc);

            ir_cells.push(TableCell {
                content: cell_elements,
                col_span,
                row_span,
                ..Default::default()
            });

            grid_col += col_span as usize;
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header,
            ..Default::default()
        });
    }

    Element::Table(Table {
        rows: ir_rows,
        ..Default::default()
    })
}

fn get_cell_at_grid_col(
    row: &crate::docx::TableRow,
    target_col: usize,
) -> Option<&crate::docx::TableCell> {
    let mut col = 0;
    for cell in &row.cells {
        let span = cell
            .properties
            .as_ref()
            .and_then(|p| p.grid_span)
            .unwrap_or(1) as usize;
        if col == target_col {
            return Some(cell);
        }
        col += span;
        if col > target_col {
            return None;
        }
    }
    None
}

// Also handle images at the block level by scanning for drawings in paragraphs
impl From<&crate::docx::DrawingInfo> for Image {
    fn from(d: &crate::docx::DrawingInfo) -> Self {
        Image {
            alt_text: d.description.clone(),
            ..Default::default()
        }
    }
}
