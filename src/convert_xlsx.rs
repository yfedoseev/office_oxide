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
        // First pass: parse all rows into (cells, style-indices).
        // Each entry is a Vec of (text, style_index_for_first_non_empty_cell)
        // — we keep style index for the first non-empty cell in each row
        // because that's what carries the font info we want to recover.
        let mut parsed_rows: Vec<Vec<(String, Option<u32>)>> = Vec::with_capacity(ws.rows.len());
        for row in &ws.rows {
            let mut cells: Vec<(String, Option<u32>)> = Vec::with_capacity(row.cells.len());
            for cell in &row.cells {
                buf.clear();
                doc.write_cell_value_fast(cell, &mut buf, &date_indices);
                let text = if buf.is_empty() {
                    String::new()
                } else {
                    std::mem::take(&mut buf)
                };
                cells.push((text, cell.style_index));
            }
            // Drop trailing empty cells.
            while cells.last().is_some_and(|(t, _)| t.is_empty()) {
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
            let nc = cells.iter().filter(|(t, _)| !t.is_empty()).count();
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
                let Some((text, style_idx)) = cells.iter().find(|(t, _)| !t.is_empty()).cloned()
                else {
                    continue;
                };
                let mut span = TextSpan::plain(text);
                if let Some(idx) = style_idx {
                    if let Some(font) = font_for(doc, idx) {
                        if let Some(size_pt) = font.size {
                            // XLSX cell font size is in points (`<font><sz val="N"/>`
                            // where N is f32). IR uses half-points; same
                            // half-pt convention as DOCX/PPTX read paths.
                            span.font_size_half_pt = Some(
                                crate::core::units::HalfPoint::from_points_rounded(size_pt)
                                    .0,
                            );
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
                for (text, _) in cells {
                    let content = if text.is_empty() {
                        Vec::new()
                    } else {
                        vec![InlineContent::Text(TextSpan::plain(text.clone()))]
                    };
                    tcells.push(TableCell {
                        content: vec![Element::Paragraph(Paragraph {
                            content,
                            ..Default::default()
                        })],
                        col_span: 1,
                        row_span: 1,
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
        let page_setup = ws.page_setup.and_then(|wsp| {
            // A worksheet that only had <pageMargins> (no dimensions) is
            // treated as "no geometry" so the renderer keeps its
            // OfficeConfig default page size.
            if wsp.width_twips == 0 || wsp.height_twips == 0 {
                return None;
            }
            Some(PageSetup {
                width_twips: wsp.width_twips,
                height_twips: wsp.height_twips,
                margin_top_twips: wsp.margin_top_twips,
                margin_bottom_twips: wsp.margin_bottom_twips,
                margin_left_twips: wsp.margin_left_twips,
                margin_right_twips: wsp.margin_right_twips,
                header_distance_twips: wsp.header_distance_twips,
                footer_distance_twips: wsp.footer_distance_twips,
                landscape: wsp.landscape,
            })
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
