use crate::format::DocumentFormat;
use crate::ir::*;

pub(crate) fn docx_to_ir(doc: &crate::docx::DocxDocument) -> DocumentIR {
    let mut elements = Vec::new();
    convert_block_elements(&doc.body.elements, &mut elements, doc);

    // Extract title from first heading
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

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            title: title.clone(),
        },
        sections: vec![Section { title, elements }],
    }
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

                if let Some(level) = heading_level {
                    elements.push(Element::Heading(Heading {
                        level: (level + 1).min(6),
                        content: convert_paragraph_inline(p, doc),
                    }));
                } else {
                    // Check for page break in runs
                    let (before_break, has_break) = split_at_page_break(p, doc);
                    if !before_break.is_empty() || !has_break {
                        elements.push(Element::Paragraph(Paragraph {
                            content: if before_break.is_empty() && !has_break {
                                convert_paragraph_inline(p, doc)
                            } else {
                                before_break
                            },
                        }));
                    }
                    if has_break {
                        elements.push(Element::ThematicBreak);
                    }
                }
                i += 1;
            },
            crate::docx::BlockElement::Table(t) => {
                elements.push(convert_table(t, doc));
                i += 1;
            },
        }
    }
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

    for rc in &run.content {
        match rc {
            crate::docx::RunContent::Text(text) => {
                content.push(InlineContent::Text(TextSpan {
                    text: text.clone(),
                    bold,
                    italic,
                    strikethrough: strike,
                    hyperlink: hyperlink_url.map(|s| s.to_string()),
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
                // Emit as a separate image element — but we're in inline context,
                // so we just note the alt text inline
                if drawing.description.is_some() {
                    content.push(InlineContent::Text(TextSpan::plain("")));
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
    Element::List(build_nested_list(is_ordered, &items, 0))
}

fn build_nested_list(ordered: bool, items: &[(u8, Vec<InlineContent>)], base_level: u8) -> List {
    let mut list_items = Vec::new();
    let mut idx = 0;

    while idx < items.len() {
        let (ilvl, content) = &items[idx];
        if *ilvl == base_level {
            // Collect any nested items immediately following at deeper levels
            let mut nested = None;
            let nested_start = idx + 1;
            let mut nested_end = nested_start;
            while nested_end < items.len() && items[nested_end].0 > base_level {
                nested_end += 1;
            }
            if nested_end > nested_start {
                nested = Some(build_nested_list(
                    ordered,
                    &items[nested_start..nested_end],
                    base_level + 1,
                ));
            }
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
            // Item at unexpected level — just add it flat
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
            });

            grid_col += col_span as usize;
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header,
        });
    }

    Element::Table(Table { rows: ir_rows })
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
        }
    }
}
