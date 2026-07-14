use crate::format::DocumentFormat;
use crate::ir::*;

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

pub(crate) fn xlsx_to_ir(doc: &crate::xlsx::XlsxDocument) -> DocumentIR {
    // Pre-compute date style indices once — avoids re-scanning format strings per cell.
    let date_indices = doc.date_style_indices();

    // Single String buffer reused across all cells — clear() keeps the heap
    // allocation; std::mem::take() moves it into TextSpan for non-empty cells.
    let mut buf = String::new();

    let mut sections = Vec::new();

    for (ws_idx, ws) in doc.worksheets.iter().enumerate() {
        // First pass: parse all rows into `CellData` — the rendered display
        // string plus the structured facts (semantic type, raw value, number
        // format) that the grid path threads into the IR so `to_ir()`
        // consumers can tell numbers/dates from text (issue #72).
        let mut parsed_rows: Vec<Vec<CellData>> = Vec::with_capacity(ws.rows.len());
        for row in &ws.rows {
            let mut cells: Vec<CellData> = Vec::with_capacity(row.cells.len());
            for cell in &row.cells {
                buf.clear();
                doc.write_cell_value_fast(cell, &mut buf, &date_indices);
                let text = if buf.is_empty() {
                    String::new()
                } else {
                    std::mem::take(&mut buf)
                };
                let (data_type, raw_number, number_format, number_format_id) =
                    cell_semantics(doc, cell, &date_indices);
                cells.push(CellData {
                    text,
                    style_index: cell.style_index,
                    data_type,
                    raw_number,
                    number_format,
                    number_format_id,
                });
            }
            // Drop trailing empty cells.
            while cells.last().is_some_and(|cd| cd.text.is_empty()) {
                cells.pop();
            }
            parsed_rows.push(cells);
        }

        // Decide row layout: a worksheet whose rows mostly have at most one
        // non-empty cell is "document style" — flowing text laid out one
        // paragraph per row. Render those rows as Paragraphs (not as a
        // 1-column Table) so the downstream PDF renderer flows them like
        // body text and honours per-paragraph font sizes.
        //
        // We choose Paragraph mode when ≥80 % of non-empty rows have ≤1
        // non-empty cell. That's permissive enough to handle real
        // worksheets that mostly hold prose but still emit a Table when a
        // genuine grid is present.
        let mut prose_score = 0usize;
        let mut nonempty_rows = 0usize;
        for cells in &parsed_rows {
            let nc = cells.iter().filter(|cd| !cd.text.is_empty()).count();
            if nc == 0 {
                continue;
            }
            nonempty_rows += 1;
            if nc <= 1 {
                prose_score += 1;
            }
        }
        let prose_mode = nonempty_rows >= 3 && prose_score * 100 >= nonempty_rows * 80;

        // Materialise any pictures or text shapes anchored on the
        // worksheet as positional IR elements so they survive the
        // round-trip back to PDF. Pictures wrap an `Element::Image`
        // in an `Element::TextBox`; text shapes wrap a styled
        // paragraph the same way. The flow renderer then paints
        // both at their absolute EMU rectangle (see
        // `render_text_box`).
        let mut image_elements: Vec<Element> =
            Vec::with_capacity(ws.images.len() + ws.text_shapes.len());
        for ts in &ws.text_shapes {
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
        for pic in &ws.images {
            let format = image_format_from_ext(&pic.format);
            let img = Image {
                alt_text: pic.alt_text.clone(),
                data: Some(pic.data.clone()),
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
                // Find the first non-empty cell.
                let Some(cd) = cells.iter().find(|cd| !cd.text.is_empty()) else {
                    continue;
                };
                let mut span = TextSpan::plain(cd.text.clone());
                if let Some(idx) = cd.style_index {
                    if let Some(font) = font_for(doc, idx) {
                        if let Some(size_pt) = font.size {
                            // XLSX cell font size is in points (`<font><sz val="N"/>`
                            // where N is f32). IR uses half-points; same
                            // half-pt convention as DOCX/PPTX read paths.
                            span.font_size_half_pt =
                                Some(crate::core::units::HalfPoint::from_points_rounded(size_pt).0);
                        }
                        if font.bold {
                            span.bold = true;
                        }
                        if font.italic {
                            span.italic = true;
                        }
                    }
                }
                out.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(span)],
                    ..Default::default()
                }));
            }
            out
        } else {
            // Genuine grid → emit a Table.
            let mut rows: Vec<TableRow> = Vec::with_capacity(parsed_rows.len());
            for (row_idx, cells) in parsed_rows.iter().enumerate() {
                let mut tcells: Vec<TableCell> = Vec::with_capacity(cells.len());
                for cd in cells {
                    let content = if cd.text.is_empty() {
                        Vec::new()
                    } else {
                        vec![InlineContent::Text(TextSpan::plain(cd.text.clone()))]
                    };
                    tcells.push(TableCell {
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
                        ..Default::default()
                    });
                }
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
        let mut combined: Vec<Element> = image_elements;
        combined.extend(elements);

        sections.push(Section {
            title: Some(ws.name.clone()),
            elements: combined,
            page_setup,
            break_type,
            ..Default::default()
        });
    }

    // Append a section for chart content. We don't render charts as graphics;
    // capturing their text (titles, axis labels, series names, cached values)
    // ensures that all human-meaningful words in the workbook appear in the
    // IR and downstream conversions, even when the chart itself isn't drawn.
    if !doc.chart_text.is_empty() {
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
        sections.push(Section {
            title: Some("Charts".to_string()),
            elements: chart_elements,
            ..Default::default()
        });
    }

    let title = sections.first().and_then(|s| s.title.clone());

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Xlsx,
            title,
            ..Default::default()
        },
        sections,
    }
}

/// A parsed spreadsheet cell carried through `xlsx_to_ir`: the rendered
/// display string plus the structured facts needed to populate the IR's
/// semantic `TableCell` fields (issue #72).
struct CellData {
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
}

/// Derive a cell's semantic type, raw numeric value, and number-format
/// metadata. Mirrors the type/date decisions in `write_cell_value_fast` but
/// preserves the structured facts instead of collapsing them to a string, so
/// `to_ir()` consumers (e.g. the WASM `toIr()` surface) can distinguish a
/// number or date cell from a text cell.
fn cell_semantics(
    doc: &crate::xlsx::XlsxDocument,
    cell: &crate::xlsx::Cell,
    date_indices: &std::collections::HashSet<u32>,
) -> (Option<CellDataType>, Option<f64>, Option<String>, Option<u32>) {
    use crate::xlsx::CellValue;

    // Number-format metadata (skip General / id 0 — carries no information).
    let fmt_id = cell
        .style_index
        .and_then(|idx| doc.styles.as_ref()?.number_format_id_for(idx))
        .filter(|&id| id != 0);
    let fmt_str = cell
        .style_index
        .and_then(|idx| doc.styles.as_ref()?.number_format_for(idx))
        .map(|s| s.to_string());

    match &cell.value {
        CellValue::Empty => (None, None, None, None),
        CellValue::Number(n) => {
            let is_date = cell.style_index.is_some_and(|i| date_indices.contains(&i));
            let ty = if is_date {
                CellDataType::Date
            } else {
                CellDataType::Number
            };
            (Some(ty), Some(*n), fmt_str, fmt_id)
        },
        // Dates almost always arrive as Number + a date format (handled above);
        // this variant is a defensive fallback and carries no serial to expose.
        CellValue::Date(_) => (Some(CellDataType::Date), None, fmt_str, fmt_id),
        CellValue::Boolean(b) => {
            (Some(CellDataType::Boolean), Some(if *b { 1.0 } else { 0.0 }), None, None)
        },
        CellValue::Error(_) => (Some(CellDataType::Error), None, None, None),
        CellValue::String(_) | CellValue::SharedString(_) => {
            (Some(CellDataType::Text), None, None, None)
        },
    }
}

/// Look up a cell's font through the workbook's stylesheet (if loaded).
/// `to_ir` runs after the document has been fully read; if styles weren't
/// parsed yet they remain `None` and we silently skip per-cell font
/// recovery rather than mutate the document during a `&self` traversal.
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
