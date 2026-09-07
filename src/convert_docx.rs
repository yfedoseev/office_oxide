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
    // `doc.headers_footers` is filled section by section, header refs
    // before footer refs within each section, and each entry records both
    // its role (`is_header`) and which pages it covers (`hf_type`). Walk
    // the section ref lists in the same order to recover the per-section
    // slice, then file each part into the matching IR slot. Splitting by a
    // cumulative index instead — the old approach — put footers in
    // `Section.header` as soon as any section had an unequal number of
    // header and footer refs, and gave every section every other section's
    // headers on top of that.
    let mut hf_iter = doc.headers_footers.iter();
    let mut per_section: Vec<SectionHeaders> = Vec::new();
    for sp in &doc.sections {
        let mut slot = SectionHeaders::default();
        for _ in 0..(sp.header_refs.len() + sp.footer_refs.len()) {
            let Some(hf) = hf_iter.next() else { break };
            let mut tmp: Vec<Element> = Vec::new();
            convert_block_elements(&hf.content, &mut tmp, doc);
            if tmp
                .iter()
                .all(|e| matches!(e, Element::Paragraph(p) if p.content.is_empty()))
            {
                continue;
            }
            let part = HeaderFooter { content: tmp };
            let target = match (hf.is_header, hf.hf_type) {
                (true, crate::docx::HeaderFooterType::First) => &mut slot.first_header,
                (true, crate::docx::HeaderFooterType::Even) => &mut slot.even_header,
                (true, crate::docx::HeaderFooterType::Default) => &mut slot.header,
                (false, crate::docx::HeaderFooterType::First) => &mut slot.first_footer,
                (false, crate::docx::HeaderFooterType::Even) => &mut slot.even_footer,
                (false, crate::docx::HeaderFooterType::Default) => &mut slot.footer,
            };
            *target = Some(part);
        }
        per_section.push(slot);
    }

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

        let page_setup = doc.sections.get(idx).and_then(section_props_to_page_setup);
        // Propagate the multi-column layout out of the source DOCX so
        // the IR carries `Section.columns` for the renderer. Without
        // this, a PDF→DOCX→PDF round-trip of a 2-column source paper
        // (arxiv preprints etc.) collapsed back to a single column on
        // read because the column count was dropped at this hop.
        let columns = doc
            .sections
            .get(idx)
            .and_then(|sp| sp.columns.filter(|n| *n >= 2).map(|n| (n, sp)))
            .map(|(n, sp)| {
                let defs = sp.column_layout.as_ref();
                ColumnLayout {
                    count: n,
                    space_twips: defs.and_then(|d| d.space),
                    separator: defs.is_some_and(|d| d.separator),
                    column_widths_twips: defs.map(|d| d.widths.clone()).unwrap_or_default(),
                }
            });

        // `<w:type>` states the break kind outright. Deriving it from the
        // section's ordinal instead inverted the common case: a document
        // whose first section is `nextPage` and whose second is
        // `continuous` came back with exactly the opposite pair. ECMA-376
        // makes `nextPage` the default when the element is absent.
        let break_type = doc
            .sections
            .get(idx)
            .and_then(|sp| sp.break_type)
            .map(|k| match k {
                crate::docx::SectionBreakKind::Continuous => SectionBreakType::Continuous,
                crate::docx::SectionBreakKind::EvenPage => SectionBreakType::EvenPage,
                crate::docx::SectionBreakKind::OddPage => SectionBreakType::OddPage,
                // The IR has no "next column" variant; it is a
                // within-page break, so `Continuous` is the closest fit.
                crate::docx::SectionBreakKind::NextColumn => SectionBreakType::Continuous,
                crate::docx::SectionBreakKind::NextPage => SectionBreakType::NextPage,
            })
            .unwrap_or(SectionBreakType::NextPage);

        let hf = per_section.get(idx).cloned().unwrap_or_default();
        ir_sections.push(Section {
            title,
            elements,
            page_setup,
            break_type,
            columns,
            header: hf.header,
            footer: hf.footer,
            first_page_header: hf.first_header,
            first_page_footer: hf.first_footer,
            even_page_header: hf.even_header,
            even_page_footer: hf.even_footer,
            ..Default::default()
        });
    }

    // Real document metadata beats a title guessed from the first heading,
    // but the guess stays as the fallback for files with no core properties.
    let cp = doc.core_properties.as_ref();
    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Docx,
            title: cp
                .and_then(|c| c.title.clone())
                .filter(|t| !t.is_empty())
                .or(doc_title),
            author: cp.and_then(|c| c.creator.clone()),
            subject: cp.and_then(|c| c.subject.clone()),
            keywords: cp
                .and_then(|c| c.keywords.as_deref())
                .map(split_keywords)
                .unwrap_or_default(),
            created: cp.and_then(|c| c.created.clone()),
            modified: cp.and_then(|c| c.modified.clone()),
            description: cp.and_then(|c| c.description.clone()),
        },
        sections: ir_sections,
    }
}

/// The six header/footer slots a single section can carry.
#[derive(Default, Clone)]
struct SectionHeaders {
    header: Option<HeaderFooter>,
    footer: Option<HeaderFooter>,
    first_header: Option<HeaderFooter>,
    first_footer: Option<HeaderFooter>,
    even_header: Option<HeaderFooter>,
    even_footer: Option<HeaderFooter>,
}

/// Split a `cp:keywords` value. OOXML has no formal separator; Word writes
/// comma- or semicolon-separated lists, and space-separated is also seen.
pub(crate) fn split_keywords(s: &str) -> Vec<String> {
    s.split([',', ';'])
        .flat_map(|part| part.split_whitespace())
        .map(|k| k.trim().to_string())
        .filter(|k| !k.is_empty())
        .collect()
}

// ---------------------------------------------------------------------------
// Formatting helpers
// ---------------------------------------------------------------------------

/// Parse an OOXML `RRGGBB` colour attribute. `auto` and malformed values
/// yield `None` so the renderer keeps its own default.
fn hex_to_rgb(s: &str) -> Option<[u8; 3]> {
    let h = s.trim().trim_start_matches('#');
    if h.eq_ignore_ascii_case("auto") || h.len() != 6 || !h.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    let v = u32::from_str_radix(h, 16).ok()?;
    Some([(v >> 16) as u8, (v >> 8) as u8, v as u8])
}

/// Map a `w:val` border style name onto the IR's `BorderStyle`. Unknown
/// styles fall back to `Single` — the same rendering Word gives an
/// unrecognised decorative border.
fn border_style_from_val(v: &str) -> BorderStyle {
    match v {
        "none" | "nil" => BorderStyle::None,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "dotted" => BorderStyle::Dotted,
        "dashed" => BorderStyle::Dashed,
        "wave" => BorderStyle::Wave,
        "dashSmallGap" => BorderStyle::DashSmallGap,
        "outset" => BorderStyle::Outset,
        "inset" => BorderStyle::Inset,
        _ => BorderStyle::Single,
    }
}

fn edge_to_ir(e: &crate::docx::BorderEdge) -> BorderLine {
    BorderLine {
        style: e
            .style
            .as_deref()
            .map(border_style_from_val)
            .unwrap_or(BorderStyle::Single),
        color: e.color.as_deref().and_then(hex_to_rgb),
        size: e.size,
        space: e.space,
    }
}

fn table_borders_to_ir(b: &crate::docx::TableBorders) -> TableBorder {
    TableBorder {
        top: b.top.as_ref().map(edge_to_ir),
        bottom: b.bottom.as_ref().map(edge_to_ir),
        left: b.left.as_ref().map(edge_to_ir),
        right: b.right.as_ref().map(edge_to_ir),
        inside_h: b.inside_h.as_ref().map(edge_to_ir),
        inside_v: b.inside_v.as_ref().map(edge_to_ir),
    }
}

fn para_borders_to_ir(b: &crate::docx::ParagraphBorders) -> ParagraphBorder {
    ParagraphBorder {
        top: b.top.as_ref().map(edge_to_ir),
        bottom: b.bottom.as_ref().map(edge_to_ir),
        left: b.left.as_ref().map(edge_to_ir),
        right: b.right.as_ref().map(edge_to_ir),
        between: b.between.as_ref().map(edge_to_ir),
    }
}

/// Resolve a `<w:highlight w:val="...">` colour name. Word's highlight
/// palette is a fixed 16-entry set, so the names map to exact RGB.
fn highlight_name_to_rgb(name: &str) -> Option<[u8; 3]> {
    Some(match name {
        "black" => [0x00, 0x00, 0x00],
        "blue" => [0x00, 0x00, 0xFF],
        "cyan" => [0x00, 0xFF, 0xFF],
        "green" => [0x00, 0xFF, 0x00],
        "magenta" => [0xFF, 0x00, 0xFF],
        "red" => [0xFF, 0x00, 0x00],
        "yellow" => [0xFF, 0xFF, 0x00],
        "white" => [0xFF, 0xFF, 0xFF],
        "darkBlue" => [0x00, 0x00, 0x80],
        "darkCyan" => [0x00, 0x80, 0x80],
        "darkGreen" => [0x00, 0x80, 0x00],
        "darkMagenta" => [0x80, 0x00, 0x80],
        "darkRed" => [0x80, 0x00, 0x00],
        "darkYellow" => [0x80, 0x80, 0x00],
        "darkGray" => [0x80, 0x80, 0x80],
        "lightGray" => [0xC0, 0xC0, 0xC0],
        _ => return None,
    })
}

/// Copy `<w:pPr>` geometry — indent, spacing, keep flags, borders,
/// shading, tabs and outline level — onto an IR paragraph. Previously
/// only `<w:jc>` survived the hop, so a fully formatted paragraph read
/// back as a bare left-aligned block.
fn apply_paragraph_properties(pp: &crate::docx::ParagraphProperties, out: &mut Paragraph) {
    if let Some(ind) = pp.indent.as_ref() {
        out.indent_left_twips = ind.left.map(|t| t.0);
        out.indent_right_twips = ind.right.map(|t| t.0);
        // A hanging indent is a negative first-line indent in the IR.
        out.first_line_indent_twips = match (ind.first_line, ind.hanging) {
            (_, Some(h)) if h.0 != 0 => Some(-h.0),
            (Some(f), _) => Some(f.0),
            _ => None,
        };
    }
    if let Some(sp) = pp.spacing.as_ref() {
        out.space_before_twips = sp.before.map(|t| t.0.max(0) as u32);
        out.space_after_twips = sp.after.map(|t| t.0.max(0) as u32);
        out.line_spacing = sp.line.as_ref().map(|l| {
            let v = l.value.max(0) as u32;
            match l.rule {
                Some(crate::docx::LineSpacingRule::Exact) => LineSpacing::Exact(v),
                Some(crate::docx::LineSpacingRule::AtLeast) => LineSpacing::AtLeast(v),
                _ => LineSpacing::Auto(v),
            }
        });
    }
    out.keep_with_next = pp.keep_next.unwrap_or(false);
    out.keep_together = pp.keep_lines.unwrap_or(false);
    out.page_break_before = pp.page_break_before.unwrap_or(false);
    out.outline_level = pp.outline_level;
    out.border = pp.borders.as_ref().map(para_borders_to_ir);
    out.background_color = pp
        .shading
        .as_ref()
        .and_then(|sh| sh.fill.as_deref())
        .and_then(hex_to_rgb);
    out.tabs = pp
        .tabs
        .iter()
        .map(|t| TabStop {
            position_twips: t.position_twips,
            alignment: match t.alignment.as_str() {
                "center" => TabAlignment::Center,
                "right" | "end" => TabAlignment::Right,
                "decimal" => TabAlignment::Decimal,
                "bar" => TabAlignment::Bar,
                _ => TabAlignment::Left,
            },
            leader: match t.leader.as_deref() {
                Some("dot") => TabLeader::Dot,
                Some("hyphen") => TabLeader::Hyphen,
                Some("underscore") => TabLeader::Underscore,
                Some("heavy") => TabLeader::Heavy,
                Some("middleDot") => TabLeader::MiddleDot,
                _ => TabLeader::None,
            },
        })
        .collect();
}

/// Build an IR `PageSetup` from a section's properties, or `None` when the
/// section states neither a page size nor margins. A `<w:sectPr>` that only
/// carries a break type or header references says nothing about the page,
/// and reporting `PageSetup::default()` for it would hand the consumer a
/// Letter-size page the document never claimed.
fn section_props_to_page_setup(sp: &crate::docx::SectionProperties) -> Option<PageSetup> {
    if sp.page_size.is_none() && sp.margins.is_none() {
        return None;
    }
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
    Some(ps)
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
                // Resolve document defaults and the style chain once; every
                // decision below (list membership, heading level, alignment,
                // geometry) reads the effective set rather than the direct
                // `w:pPr` alone.
                let eff = effective_paragraph_props(p, doc);
                let eff_ref = eff.as_ref();

                // Check if this is a list item — group consecutive list
                // paragraphs. `<w:numId w:val="0"/>` is the ECMA-376 way of
                // saying "this paragraph has no numbering" — usually a
                // style-level list switched off for one paragraph. Treating
                // it as a list turned an ordinary paragraph into a bullet.
                if let Some(nr) = eff_ref
                    .and_then(|pp| pp.numbering_ref.as_ref())
                    .filter(|nr| nr.num_id != 0)
                {
                    let list_element = convert_list_group(blocks, &mut i, nr.num_id, doc);
                    elements.push(list_element);
                    continue;
                }

                // Check for heading
                let heading_level = resolve_heading_level(p, doc);
                let alignment = eff_ref.and_then(paragraph_alignment);

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
                let has_bottom_border = eff_ref.is_some_and(|pp| pp.has_bottom_border);
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
                    let (before_break, hard_break) = split_at_page_break(p, doc);
                    let frame_pos = paragraph_frame_position(p);
                    if !before_break.is_empty() || hard_break.is_none() {
                        let mut para = Paragraph {
                            content: if before_break.is_empty() && hard_break.is_none() {
                                convert_paragraph_inline(p, doc)
                            } else {
                                before_break
                            },
                            frame_position: frame_pos,
                            alignment,
                            ..Default::default()
                        };
                        if let Some(pp) = eff_ref {
                            apply_paragraph_properties(pp, &mut para);
                        }
                        elements.push(Element::Paragraph(para));
                    }
                    // A `<w:br w:type="page"/>` is a page break, not a
                    // horizontal rule. Emitting `ThematicBreak` here used to
                    // put a `---` in the markdown of every paginated document
                    // and made `PageBreak`/`ColumnBreak` unreachable.
                    match hard_break {
                        Some(HardBreak::Page) => elements.push(Element::PageBreak),
                        Some(HardBreak::Column) => elements.push(Element::ColumnBreak),
                        None => {},
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
fn paragraph_alignment(pp: &crate::docx::ParagraphProperties) -> Option<ParagraphAlignment> {
    let jc = pp.justification.as_ref()?;
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

/// Fold document defaults and the paragraph's style chain into its
/// effective `w:pPr`. Falls back to the direct properties when the package
/// has no stylesheet.
fn effective_paragraph_props(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> Option<crate::docx::ParagraphProperties> {
    match doc.styles.as_ref() {
        Some(sheet) => Some(sheet.effective_paragraph_properties(p.properties.as_ref())),
        None => p.properties.clone(),
    }
}

/// Recognise the `Heading1`..`Heading9` style-id convention and the
/// matching `heading 1`..`heading 9` style names. Returns a 0-based
/// outline level to match `w:outlineLvl`.
fn heading_level_from_style_name(name: &str) -> Option<u8> {
    let lower = name.trim().to_ascii_lowercase();
    let digits = lower.strip_prefix("heading")?.trim();
    let n: u8 = digits.parse().ok()?;
    (1..=9).contains(&n).then(|| n - 1)
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
    if let Some(lvl) = styles.resolve_outline_level(style_id) {
        return Some(lvl);
    }
    // Fall back to the naming convention. Word's built-in heading styles
    // often carry no `<w:outlineLvl>` of their own, so an outline-level-only
    // check found no headings at all in documents that use them.
    if let Some(lvl) = heading_level_from_style_name(style_id) {
        return Some(lvl);
    }
    styles
        .styles
        .get(style_id)
        .and_then(|s| s.name.as_deref())
        .and_then(heading_level_from_style_name)
}

/// Everything `convert_run` needs from the enclosing document to resolve a
/// run's effective formatting.
struct RunContext<'a> {
    theme: Option<&'a crate::core::theme::Theme>,
    styles: Option<&'a crate::docx::StyleSheet>,
    paragraph_style_id: Option<&'a str>,
}

/// Resolve a hyperlink's target into a URL string. Internal `w:anchor`
/// links become fragment references (`#bookmark`) rather than being
/// dropped — a table-of-contents built from internal links used to lose
/// every one of its destinations.
fn hyperlink_url(hl: &crate::docx::Hyperlink) -> Option<String> {
    match &hl.target {
        crate::docx::HyperlinkTarget::External(url) => Some(url.clone()),
        crate::docx::HyperlinkTarget::Internal(anchor) if !anchor.is_empty() => {
            Some(format!("#{anchor}"))
        },
        crate::docx::HyperlinkTarget::Internal(_) => None,
    }
}

fn run_context<'a>(
    p: &'a crate::docx::Paragraph,
    doc: &'a crate::docx::DocxDocument,
) -> RunContext<'a> {
    RunContext {
        theme: doc.theme.as_ref(),
        styles: doc.styles.as_ref(),
        paragraph_style_id: p.properties.as_ref().and_then(|pp| pp.style_id.as_deref()),
    }
}

fn convert_paragraph_inline(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> Vec<InlineContent> {
    let ctx = run_context(p, doc);
    let mut content = Vec::new();
    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                convert_run(run, None, &ctx, &mut content);
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                let url = hyperlink_url(hl);
                for run in &hl.runs {
                    convert_run(run, url.as_deref(), &ctx, &mut content);
                }
            },
        }
    }
    content
}

fn convert_run(
    run: &crate::docx::Run,
    hyperlink_url: Option<&str>,
    ctx: &RunContext<'_>,
    content: &mut Vec<InlineContent>,
) {
    let theme = ctx.theme;
    // Fold document defaults, the paragraph style chain, the run's
    // character style and its direct `w:rPr` into one effective set.
    // Reading `run.properties` alone meant a document that keeps its
    // formatting in styles — every Word template does — came back with
    // none of it.
    let resolved;
    let effective: Option<&crate::docx::RunProperties> = match ctx.styles {
        Some(sheet) => {
            resolved =
                sheet.effective_run_properties(ctx.paragraph_style_id, run.properties.as_ref());
            Some(&resolved)
        },
        None => run.properties.as_ref(),
    };
    let bold = effective.and_then(|rp| rp.bold).unwrap_or(false);
    let italic = effective.and_then(|rp| rp.italic).unwrap_or(false);
    let strike = effective
        .and_then(|rp| rp.strike.or(rp.dstrike))
        .unwrap_or(false);
    // `<w:sz w:val="N"/>` is already in half-points; IR uses the
    // same encoding. See `crate::core::units::HalfPoint::from_word_sz`
    // for the cross-format invariant (also: PPTX hundredths-pt,
    // XLSX points-as-f32 must convert here).
    let font_size_half_pt = effective.and_then(|rp| {
        rp.font_size
            .map(|hp| crate::core::units::HalfPoint::from_word_sz(hp.0).0)
    });
    // `<w:rFonts w:ascii="...">` carries the run's face name. Without
    // forwarding it onto `TextSpan.font_name`, the IR→PDF renderer
    // falls back to the page builder's default font (Helvetica) and
    // every PDF→DOCX→PDF round-trip loses every typeface — even when
    // the DOCX writer correctly embedded the source-PDF font program
    // under `word/fonts/`.
    let font_name = effective.and_then(|rp| rp.font_name.clone());
    // Propagate `<w:color w:val="RRGGBB"/>` so PDF→DOCX→PDF round-trips
    // preserve coloured text (red "0" in `pdfs_pdfium/text_color.pdf`
    // and the like). Theme / system / auto colours fall through to
    // the renderer default for now — resolving them properly needs the
    // document's `theme.xml`, which the current convert path doesn't
    // thread in.
    // Resolve `<w:color>` through the document theme. Matching only
    // `ColorRef::Rgb` silently dropped every theme-coloured run — and the
    // `w:val` fallback OOXML writes next to `w:themeColor` was discarded at
    // parse time, so even a document with no theme part came out colourless.
    let text_color = effective
        .and_then(|rp| rp.color.as_ref())
        .and_then(|c| c.resolve_opt(theme))
        .map(|rgb| rgb.0);
    // Remaining `w:rPr` toggles. Half of `TextSpan`'s fields used to be
    // permanently empty for DOCX because these were parsed and then never
    // read; underline in particular is the single most common piece of
    // direct formatting after bold/italic.
    let rp = effective;
    let underline = rp.and_then(|rp| rp.underline.as_ref()).map(|u| match u {
        crate::docx::UnderlineType::Single => UnderlineStyle::Single,
        crate::docx::UnderlineType::Double => UnderlineStyle::Double,
        crate::docx::UnderlineType::Thick => UnderlineStyle::Thick,
        crate::docx::UnderlineType::Dotted => UnderlineStyle::Dotted,
        crate::docx::UnderlineType::Dash => UnderlineStyle::Dash,
        crate::docx::UnderlineType::DotDash => UnderlineStyle::DotDash,
        crate::docx::UnderlineType::DotDotDash => UnderlineStyle::DotDotDash,
        crate::docx::UnderlineType::Wave => UnderlineStyle::Wave,
        crate::docx::UnderlineType::Words => UnderlineStyle::Words,
        crate::docx::UnderlineType::None => UnderlineStyle::None,
        crate::docx::UnderlineType::Other(_) => UnderlineStyle::Single,
    });
    // The DOCX writer encodes `TextSpan::highlight` as `<w:shd w:fill>`
    // inside `w:rPr`, so read that first and fall back to Word's named
    // `<w:highlight>` palette for documents authored elsewhere.
    let highlight = rp
        .and_then(|rp| rp.shading_fill.as_deref())
        .and_then(hex_to_rgb)
        .or_else(|| {
            rp.and_then(|rp| rp.highlight.as_deref())
                .and_then(highlight_name_to_rgb)
        });
    let vertical_align = rp.and_then(|rp| rp.vertical_align).map(|va| match va {
        crate::docx::VerticalAlign::Superscript => VerticalAlign::Superscript,
        crate::docx::VerticalAlign::Subscript => VerticalAlign::Subscript,
        crate::docx::VerticalAlign::Baseline => VerticalAlign::Baseline,
    });
    let all_caps = rp.and_then(|rp| rp.caps).unwrap_or(false);
    let small_caps = rp.and_then(|rp| rp.small_caps).unwrap_or(false);
    let char_spacing_half_pt = rp.and_then(|rp| rp.char_spacing);

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
                    underline: underline.clone(),
                    highlight,
                    vertical_align: vertical_align.clone(),
                    all_caps,
                    small_caps,
                    char_spacing_half_pt,
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

/// The kind of hard break that terminated a paragraph's inline content.
#[derive(Clone, Copy, PartialEq, Eq)]
enum HardBreak {
    Page,
    Column,
}

fn split_at_page_break(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> (Vec<InlineContent>, Option<HardBreak>) {
    let ctx = run_context(p, doc);
    let mut content = Vec::new();
    let mut has_break: Option<HardBreak> = None;

    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                for rc in &run.content {
                    match rc {
                        crate::docx::RunContent::Break(crate::docx::BreakType::Page) => {
                            has_break.get_or_insert(HardBreak::Page);
                        },
                        crate::docx::RunContent::Break(crate::docx::BreakType::Column) => {
                            has_break.get_or_insert(HardBreak::Column);
                        },
                        _ => {},
                    }
                }
                if has_break.is_none() {
                    convert_run(run, None, &ctx, &mut content);
                }
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                if has_break.is_none() {
                    let url = hyperlink_url(hl);
                    for run in &hl.runs {
                        convert_run(run, url.as_deref(), &ctx, &mut content);
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
    // The marker style and start value of the *shallowest* level in the
    // group describe the list the IR is about to build. Both used to be
    // resolved and then thrown away, so a list starting at 5 rendered as
    // "1." and `a) b) c)` was indistinguishable from `1. 2. 3.`.
    let mut top_ilvl: Option<u8> = None;
    let mut start_number: Option<u32> = None;
    let mut style: Option<ListStyle> = None;

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
                        if top_ilvl.is_none_or(|t| nr.ilvl < t) {
                            top_ilvl = Some(nr.ilvl);
                            style = number_format_to_list_style(&level.format);
                            // `w:start` defaults to 1; only report an
                            // explicit non-default so renderers that ignore
                            // the field are not silently contradicted.
                            start_number = (level.start != 1).then_some(level.start);
                        }
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
    let mut list = crate::ir::build_nested_list(is_ordered, &items, 0);
    list.start_number = start_number;
    list.style = style;
    Element::List(list)
}

fn number_format_to_list_style(f: &crate::docx::NumberFormat) -> Option<ListStyle> {
    use crate::docx::NumberFormat as NF;
    Some(match f {
        NF::Decimal => ListStyle::Decimal,
        NF::Bullet => ListStyle::Bullet,
        NF::LowerLetter => ListStyle::LowerAlpha,
        NF::UpperLetter => ListStyle::UpperAlpha,
        NF::LowerRoman => ListStyle::LowerRoman,
        NF::UpperRoman => ListStyle::UpperRoman,
        NF::None => return None,
        NF::Other(_) => return None,
    })
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
        let rp = row.properties.as_ref();
        let is_header = rp.is_some_and(|p| p.is_header);

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

            let cp = cell.properties.as_ref();
            ir_cells.push(TableCell {
                content: cell_elements,
                col_span,
                row_span,
                background_color: cp
                    .and_then(|p| p.shading.as_ref())
                    .and_then(|sh| sh.fill.as_deref())
                    .and_then(hex_to_rgb),
                border: cp.and_then(|p| p.borders.as_ref()).map(table_borders_to_ir),
                vertical_align: cp.and_then(|p| p.v_align).map(|va| match va {
                    crate::docx::CellVAlign::Top => CellVerticalAlign::Top,
                    crate::docx::CellVAlign::Center => CellVerticalAlign::Center,
                    crate::docx::CellVAlign::Bottom => CellVerticalAlign::Bottom,
                }),
                width_twips: cp.and_then(|p| p.width.as_ref()).and_then(dxa_width),
                padding: cp.and_then(|p| p.margins.as_ref()).map(|m| CellPadding {
                    top_twips: m.top.map(|v| v.max(0) as u32),
                    bottom_twips: m.bottom.map(|v| v.max(0) as u32),
                    left_twips: m.left.map(|v| v.max(0) as u32),
                    right_twips: m.right.map(|v| v.max(0) as u32),
                }),
                text_direction: cp
                    .and_then(|p| p.text_direction.as_deref())
                    .map(|td| match td {
                        "tbRl" | "tbRlV" | "vert" => TextDirection::TbRl,
                        "btLr" | "lrTbV" | "vert270" => TextDirection::BtLr,
                        _ => TextDirection::LrTb,
                    }),
                ..Default::default()
            });

            grid_col += col_span as usize;
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header,
            height_twips: rp.and_then(|p| p.height).map(|h| h.max(0) as u32),
            allow_break: !rp.is_some_and(|p| p.cant_split),
            repeat_as_header: is_header,
        });
    }

    let tp = table.properties.as_ref();
    Element::Table(Table {
        rows: ir_rows,
        column_widths_twips: table.grid.iter().map(|t| t.0.max(0) as u32).collect(),
        border: tp.and_then(|p| p.borders.as_ref()).map(table_borders_to_ir),
        alignment: tp.and_then(|p| p.justification).map(|j| match j {
            crate::docx::Justification::Center => TableAlignment::Center,
            crate::docx::Justification::Right => TableAlignment::Right,
            _ => TableAlignment::Left,
        }),
        // The IR carries a single padding value; `w:tblCellMar` has four
        // edges. Word writes them equal in practice (and our own writer
        // always does), so take the left edge as the representative.
        cell_padding_twips: tp
            .and_then(|p| p.cell_margins.as_ref())
            .and_then(|m| m.left.or(m.top))
            .map(|v| v.max(0) as u32),
        caption: tp.and_then(|p| p.caption.clone()),
        width_twips: tp.and_then(|p| p.width.as_ref()).and_then(dxa_width),
        indent_left_twips: tp.and_then(|p| p.indent).map(|t| t.0),
    })
}

/// Read a `w:tblW` / `w:tcW` preferred width, but only when it is an
/// absolute twip measurement. `auto`, `nil` and percentage widths carry no
/// twip value, and reporting one would be a confidently wrong number.
fn dxa_width(w: &crate::docx::TableWidth) -> Option<u32> {
    if w.width_type == crate::docx::TableWidthType::Dxa && w.value > 0 {
        Some(w.value as u32)
    } else {
        None
    }
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
