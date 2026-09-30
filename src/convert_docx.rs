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
    // The body-level `<w:sectPr>` describes a final section even when
    // no element follows the last break: its headers and footers are
    // still that section's. Skipping the empty window shifted every
    // `doc.sections[idx]` lookup below and dropped the trailing
    // section's headers outright.
    if prev < total || windows.len() < doc.sections.len() || windows.is_empty() {
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
            // A first-page part in a section without `w:titlePg` is never
            // shown by Word; carrying it made every surface print it and
            // the writer, which emits `w:titlePg` for any first-page
            // part, switch it on.
            if !hf.active {
                continue;
            }
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

        // Must use the same extraction the write path's "is the title
        // already present in the elements" check uses (`inline_to_text`),
        // or a heading containing a `LineBreak` disagrees between the two
        // and gets duplicated on every write.
        let title = elements.iter().find_map(|e| {
            if let Element::Heading(h) = e {
                Some(inline_to_text(&h.content))
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
            footnote_settings: doc
                .sections
                .get(idx)
                .and_then(|sp| sp.footnote_properties.as_ref())
                .map(note_props_to_ir),
            endnote_settings: doc
                .sections
                .get(idx)
                .and_then(|sp| sp.endnote_properties.as_ref())
                .map(note_props_to_ir),
            ..Default::default()
        });
    }

    // Footnote, endnote and comment bodies are block-level content that
    // belongs to the document as a whole. Append them to the last section so
    // they reach every renderer instead of being dropped on the floor —
    // `Element::Footnote` / `Element::Endnote` were produced by no converter
    // before this.
    if let Some(last) = ir_sections.last_mut() {
        for n in &doc.footnotes {
            let (marker, rest) = extract_note_marker(&n.content, "FootnoteReference");
            let mut content = Vec::new();
            convert_block_elements(rest, &mut content, doc);
            if !content.is_empty() {
                last.elements.push(Element::Footnote(Note {
                    id: n.id,
                    content,
                    marker,
                    author: None,
                }));
            }
        }
        for n in &doc.endnotes {
            let (marker, rest) = extract_note_marker(&n.content, "EndnoteReference");
            let mut content = Vec::new();
            convert_block_elements(rest, &mut content, doc);
            if !content.is_empty() {
                last.elements.push(Element::Endnote(Note {
                    id: n.id,
                    content,
                    marker,
                    author: None,
                }));
            }
        }
        // Comments are annotations rather than body text; carry them as
        // endnotes with the author kept in the marker so nothing is lost,
        // and in the structured `author` field.
        for n in &doc.comments {
            let mut content = Vec::new();
            convert_block_elements(&n.content, &mut content, doc);
            if !content.is_empty() {
                // The same label the spreadsheet and slide converters use,
                // so every surface says "Comment (Author): …".
                let marker = match n.author.as_deref() {
                    Some(a) => format!("Comment ({a})"),
                    None => "Comment".to_string(),
                };
                last.elements.push(Element::Endnote(Note {
                    id: n.id,
                    content,
                    marker: Some(marker),
                    author: n.author.clone(),
                }));
            }
        }
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
            has_macros: doc.has_macros,
            text_truncated: false,
            ..crate::core::core_properties::ooxml_metadata_extras(
                cp,
                doc.app_properties.as_ref(),
                Some(&doc.package_properties),
            )
        },
        sections: ir_sections,
        defined_names: Vec::new(),
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

/// Split a `cp:keywords` value; see
/// [`crate::core::core_properties::split_keywords`].
pub(crate) fn split_keywords(s: &str) -> Vec<String> {
    crate::core::core_properties::split_keywords(s)
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

/// The single colour a `w:shd` (ECMA-376 §17.3.5) paints: its pattern
/// (`w:val`, ST_Shd) in `w:color` over `w:fill`. A percentage pattern is
/// exactly that share of the pattern colour; a stripe or cross pattern is
/// represented by the share of the area its lines cover, the average
/// colour the eye sees, since the IR holds one colour per area. An
/// automatic pattern colour is black and an automatic fill white (paper);
/// `clear` with an automatic fill, and `nil`, are no shading at all.
fn shading_rgb(sh: &crate::docx::table::Shading) -> Option<[u8; 3]> {
    let fill = sh.fill.as_deref().and_then(hex_to_rgb);
    let pattern = sh.pattern.as_deref().unwrap_or("clear");
    let coverage: f32 = match pattern {
        "nil" => return None,
        "clear" => return fill,
        "solid" => 1.0,
        "horzStripe" | "vertStripe" | "diagStripe" | "reverseDiagStripe" => 0.5,
        "thinHorzStripe" | "thinVertStripe" | "thinDiagStripe" | "thinReverseDiagStripe" => 0.25,
        "horzCross" | "diagCross" => 0.75,
        "thinHorzCross" | "thinDiagCross" => 0.4375,
        p => match p.strip_prefix("pct").and_then(|n| n.parse::<u8>().ok()) {
            Some(n) if n <= 100 => f32::from(n) / 100.0,
            // An unknown pattern: the fill is the one colour known.
            _ => return fill,
        },
    };
    let color = sh
        .color
        .as_deref()
        .and_then(hex_to_rgb)
        .unwrap_or([0, 0, 0]);
    let base = fill.unwrap_or([0xFF, 0xFF, 0xFF]);
    let mix =
        |b: u8, c: u8| (f32::from(b) * (1.0 - coverage) + f32::from(c) * coverage).round() as u8;
    Some([
        mix(base[0], color[0]),
        mix(base[1], color[1]),
        mix(base[2], color[2]),
    ])
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
            // saturating_neg: -i32::MIN overflows, and a hanging indent that
            // large is nonsense anyway.
            (_, Some(h)) if h.0 != 0 => Some(h.0.saturating_neg()),
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
    out.border = pp.borders.as_deref().map(para_borders_to_ir);
    out.background_color = pp.shading.as_deref().and_then(shading_rgb);
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

/// Split a custom footnote/endnote mark off the front of a note body, if
/// present. The writer (`docx/write.rs::generate_notes_xml`) puts the mark
/// in its own leading paragraph — a single run styled `style_name`
/// ("FootnoteReference"/"EndnoteReference") with the literal glyph as its
/// only content — so a real Word auto-number run (which carries the same
/// style but no `w:t`, just an empty `<w:footnoteRef/>`) is never mistaken
/// for a custom mark: `content` there stays empty, `Text` never appears.
fn extract_note_marker<'a>(
    content: &'a [crate::docx::BlockElement],
    style_name: &str,
) -> (Option<String>, &'a [crate::docx::BlockElement]) {
    let Some(crate::docx::BlockElement::Paragraph(p)) = content.first() else {
        return (None, content);
    };
    let [crate::docx::ParagraphContent::Run(run)] = p.content.as_slice() else {
        return (None, content);
    };
    if run
        .properties
        .as_ref()
        .and_then(|rp| rp.style_id.as_deref())
        != Some(style_name)
    {
        return (None, content);
    }
    let [crate::docx::RunContent::Text(text)] = run.content.as_slice() else {
        return (None, content);
    };
    if text.is_empty() {
        return (None, content);
    }
    (Some(text.clone()), &content[1..])
}

/// Same property set as [`apply_paragraph_properties`], but for a promoted
/// `Heading`. Outline-level promotion used to keep only the heading's level,
/// content, frame position and alignment — indent, spacing, line spacing,
/// keep-with-next/together, shading, borders and tabs all vanished in the
/// same step, since `Heading` had no fields to receive them.
fn apply_paragraph_properties_to_heading(pp: &crate::docx::ParagraphProperties, out: &mut Heading) {
    if let Some(ind) = pp.indent.as_ref() {
        out.indent_left_twips = ind.left.map(|t| t.0);
        out.indent_right_twips = ind.right.map(|t| t.0);
        out.first_line_indent_twips = match (ind.first_line, ind.hanging) {
            (_, Some(h)) if h.0 != 0 => Some(h.0.saturating_neg()),
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
    out.border = pp.borders.as_deref().map(para_borders_to_ir);
    out.background_color = pp.shading.as_deref().and_then(shading_rgb);
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

fn note_props_to_ir(np: &crate::docx::NoteProperties) -> NoteSettings {
    NoteSettings {
        position: np.position.clone(),
        number_format: np.number_format.clone(),
        start: np.start,
        restart: np.restart.clone(),
    }
}

/// Build an IR `PageSetup` from a section's properties, or `None` when the
/// section states neither a page size nor margins. A `<w:sectPr>` that only
/// carries a break type or header references says nothing about the page,
/// and reporting `PageSetup::default()` for it would hand the consumer a
/// Letter-size page the document never claimed.
fn section_props_to_page_setup(sp: &crate::docx::SectionProperties) -> Option<PageSetup> {
    if sp.page_size.is_none() && sp.margins.is_none() && sp.page_numbering.is_none() {
        return None;
    }
    let mut ps = PageSetup::default();
    if let Some(pn) = &sp.page_numbering {
        ps.page_number_start = pn.start;
        ps.page_number_format = pn.format.clone();
    }
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
        if let Some(g) = m.gutter {
            ps.gutter_twips = g.0.max(0) as u32;
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
    // A numId resumed later in the same block sequence (after a non-list
    // paragraph interrupts it) with no explicit override continues
    // counting from where it left off, per OOXML/Word semantics — not a
    // fresh 1. Tracks the next start number per numId across the several
    // `convert_list_group` calls this loop makes.
    let mut numbering_counts: std::collections::HashMap<u32, u32> =
        std::collections::HashMap::new();
    while i < blocks.len() {
        match &blocks[i] {
            crate::docx::BlockElement::Paragraph(p) => {
                // Resolve document defaults and the style chain once; every
                // decision below (list membership, heading level, alignment,
                // geometry) reads the effective set rather than the direct
                // `w:pPr` alone.
                let eff = effective_paragraph_props(p, doc);
                let eff_ref = eff.as_ref();

                // Heading level wins over list membership. Word's own
                // multilevel-list "Heading" gallery attaches numPr/ilfo to
                // the heading styles themselves, so a numbered heading
                // ("1. Introduction", "2.3 Scope") is the normal shape of
                // headings in real documents — checking list membership
                // first turned every one of them into a ListItem, leaving
                // the IR with no Headings, no Section.title, no guessed
                // metadata.title.
                let heading_level = resolve_heading_level(p, doc);

                // Check if this is a list item — group consecutive list
                // paragraphs. `<w:numId w:val="0"/>` is the ECMA-376 way of
                // saying "this paragraph has no numbering" — usually a
                // style-level list switched off for one paragraph. Treating
                // it as a list turned an ordinary paragraph into a bullet.
                if heading_level.is_none()
                    && let Some(num_id) = eff_ref
                        .and_then(|pp| pp.numbering_ref.as_ref())
                        .map(|nr| nr.num_id)
                        .filter(|&id| id != 0)
                {
                    let list_element =
                        convert_list_group(blocks, &mut i, num_id, doc, &mut numbering_counts, eff);
                    elements.push(list_element);
                    continue;
                }

                let alignment = eff_ref.and_then(paragraph_alignment);

                // Detect "horizontal rule" encoding: empty paragraph
                // with a single bottom border. pdf_to_ir round-trips
                // ThematicBreak through DOCX as exactly this shape;
                // recover it here so the renderer draws a rule.
                // The paragraph's inline content, cut at hard breaks —
                // converted once and used for every decision below.
                let mut segments = split_at_hard_breaks(p, doc);
                let is_empty_para = segments.iter().all(|(inline, _)| {
                    inline.iter().all(|ic| {
                        matches!(ic,
                            crate::ir::InlineContent::Text(s) if s.text.is_empty()
                        )
                    })
                });
                let has_bottom_border = eff_ref.is_some_and(|pp| pp.has_bottom_border);
                if is_empty_para && has_bottom_border {
                    elements.push(Element::ThematicBreak);
                    // The paragraph has no text, not no content: a heading
                    // style with a bottom border that holds only a chart,
                    // a picture or a text box used to `continue` past the
                    // collectors below and lose the drawing entirely.
                    collect_paragraph_floats(p, doc, elements);
                    collect_paragraph_inline_images(p, doc, elements);
                    collect_paragraph_text_boxes(p, doc, elements);
                    collect_paragraph_chart_text(p, elements);
                    i += 1;
                    continue;
                }

                if let Some(level) = heading_level {
                    let mut heading = Heading {
                        level: (level + 1).min(6),
                        content: convert_heading_inline(p, doc),
                        frame_position: paragraph_frame_position(p),
                        alignment,
                        ..Default::default()
                    };
                    if let Some(pp) = eff_ref {
                        apply_paragraph_properties_to_heading(pp, &mut heading);
                    }
                    elements.push(Element::Heading(heading));
                } else {
                    // A hard break splits the paragraph: the runs before it,
                    // the break, the runs after it. The runs after used to
                    // be dropped — and Word writes a manual page break as
                    // `<w:br w:type="page"/>` at the *start* of the next
                    // paragraph, so that paragraph's whole text vanished
                    // from every IR-backed surface while `plain_text()`
                    // kept it (75 corpus files).
                    let frame_pos = paragraph_frame_position(p);
                    let make_para = |content: Vec<InlineContent>| {
                        let mut para = Paragraph {
                            content,
                            frame_position: frame_pos.clone(),
                            alignment: alignment.clone(),
                            ..Default::default()
                        };
                        if let Some(pp) = eff_ref {
                            apply_paragraph_properties(pp, &mut para);
                        }
                        Element::Paragraph(para)
                    };
                    if segments.len() == 1 && segments[0].1.is_none() {
                        let (mut inline, _) = segments.swap_remove(0);
                        inline.shrink_to_fit();
                        elements.push(make_para(inline));
                    } else {
                        // The paragraph's own slot is the first segment when
                        // it holds text; later segments are paragraphs only
                        // when they hold text, so a paragraph that is just
                        // a page break stays just a page break.
                        for (content, brk) in segments.drain(..) {
                            if !content.is_empty() {
                                elements.push(make_para(content));
                            }
                            // A `<w:br w:type="page"/>` is a page break, not
                            // a horizontal rule. Emitting `ThematicBreak`
                            // here used to put a `---` in the markdown of
                            // every paginated document.
                            match brk {
                                Some(HardBreak::Page) => elements.push(Element::PageBreak),
                                Some(HardBreak::Column) => elements.push(Element::ColumnBreak),
                                None => {},
                            }
                        }
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
                // Text-box bodies are ordinary block content drawn in a
                // frame. Leaving them unread silently dropped most of the
                // prose in documents that lay text out with shapes.
                collect_paragraph_text_boxes(p, doc, elements);
                // Native charts keep every word they display in a separate
                // part (`word/charts/chartN.xml`). The reader resolves it
                // at open time; hoist the recovered lines to paragraph
                // siblings so the chart's title, categories, series names
                // and data values reach the IR.
                collect_paragraph_chart_text(p, elements);
                i += 1;
            },
            crate::docx::BlockElement::Table(t) => {
                let table_elem = convert_table(t, doc);
                // A table's accessibility caption (`w:tblCaption`) sits
                // in `Table.caption`, but the writer also emits it as a
                // visible "Caption"-styled paragraph immediately before
                // `<w:tbl>` (nothing else ever renders `Table.caption`
                // as visible text). Reading it back turned that
                // paragraph into an ordinary sibling `Element::Paragraph`
                // alongside the table's own `caption` field — the same
                // text represented twice — and the next write emitted
                // BOTH, growing by one duplicate paragraph every
                // round trip. Absorbing the immediately-preceding
                // matching paragraph here instead keeps the caption
                // represented exactly once.
                if let Element::Table(Table {
                    caption: Some(cap), ..
                }) = &table_elem
                {
                    let cap = cap.trim();
                    let last_matches = matches!(
                        elements.last(),
                        Some(Element::Paragraph(p)) if paragraph_plain_text(p).trim() == cap
                    );
                    if last_matches {
                        elements.pop();
                    }
                }
                elements.push(table_elem);
                i += 1;
            },
        }
    }
}

/// Plain-text content of a paragraph's inline runs, no formatting.
fn paragraph_plain_text(p: &Paragraph) -> String {
    let mut out = String::new();
    for content in &p.content {
        if let InlineContent::Text(span) = content {
            out.push_str(&span.text);
        }
    }
    out
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
                    // A linked picture has no bytes in the package, only
                    // its target; it used to be dropped without trace.
                    let (data, ext) = match embedded_image(d, doc) {
                        Some((data, ext)) => (Some(data), ext),
                        None if d.linked_image.is_some() => (None, None),
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
                        decorative: d.decorative,
                        data,
                        source_url: d.linked_image.clone(),
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

/// Emit one IR paragraph per line of text recovered from a chart or
/// SmartArt diagram part referenced by a drawing in this paragraph.
fn collect_paragraph_chart_text(p: &crate::docx::Paragraph, out: &mut Vec<Element>) {
    for pc in &p.content {
        let runs: &[crate::docx::Run] = match pc {
            crate::docx::ParagraphContent::Run(r) => std::slice::from_ref(r),
            crate::docx::ParagraphContent::Hyperlink(hl) => &hl.runs,
        };
        for run in runs {
            for rc in &run.content {
                let crate::docx::RunContent::Drawing(d) = rc else {
                    continue;
                };
                for line in d.chart_text.iter().chain(d.dgm_text.iter()) {
                    out.push(Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan::plain(line.clone()))],
                        ..Default::default()
                    }));
                }
            }
        }
    }
}

fn collect_paragraph_text_boxes(
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
                if let crate::docx::RunContent::TextBox(blocks) = rc {
                    let mut content = Vec::new();
                    convert_block_elements(blocks, &mut content, doc);
                    if content.is_empty() {
                        continue;
                    }
                    out.push(Element::TextBox(TextBox {
                        content,
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

    let (data, ext) = match embedded_image(d, doc) {
        Some((data, ext)) => (Some(data), ext),
        None if d.linked_image.is_some() => (None, None),
        None => return None,
    };
    let format = ext.as_deref().and_then(|e| match e {
        "png" => Some(ImageFormat::Png),
        "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
        _ => None,
    });
    Some(Element::Image(Image {
        alt_text: d.description.clone(),
        decorative: d.decorative,
        data,
        source_url: d.linked_image.clone(),
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

/// A drawing's embedded picture bytes and file extension, if the package
/// holds them.
fn embedded_image(
    d: &crate::docx::DrawingInfo,
    doc: &crate::docx::DocxDocument,
) -> Option<(Vec<u8>, Option<String>)> {
    if d.relationship_id.is_empty() {
        return None;
    }
    doc.images.get(&d.relationship_id).cloned()
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
    let mut eff = match doc.styles.as_ref() {
        Some(sheet) => Some(sheet.effective_paragraph_properties(p.properties.as_ref())),
        None => p.properties.clone(),
    };
    // A numbered paragraph's indentation comes from its numbering level's
    // `w:pPr/w:ind` (ECMA-376 §17.9) unless the paragraph sets its own;
    // Word applies it over the paragraph style's.
    if let Some(pp) = eff.as_mut()
        && p.properties.as_ref().is_none_or(|d| d.indent.is_none())
        && let Some(nr) = pp.numbering_ref.as_ref().filter(|nr| nr.num_id != 0)
        && let Some(ind) = doc
            .numbering
            .as_ref()
            .and_then(|n| n.resolve_level(nr.num_id, nr.ilvl))
            .and_then(|l| l.indent.clone())
    {
        pp.indent = Some(ind);
    }
    eff
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
    convert_inline_with(p, run_context(p, doc))
}

/// A heading's spans, without the formatting its own paragraph style
/// supplies. `Heading 1`'s `<w:b/>`/`<w:sz>` *are* the heading — folding
/// them into every span rendered `<h1><strong>…</strong></h1>` and
/// `# **…**` on every styled heading, which is not what Word shows and
/// not what pandoc or python-docx report. Direct formatting and character
/// styles still apply.
fn convert_heading_inline(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> Vec<InlineContent> {
    let ctx = RunContext {
        paragraph_style_id: None,
        ..run_context(p, doc)
    };
    convert_inline_with(p, ctx)
}

fn convert_inline_with(p: &crate::docx::Paragraph, ctx: RunContext<'_>) -> Vec<InlineContent> {
    let mut content = Vec::new();
    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                convert_run(run, None, &ctx, &mut content);
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                let url = hyperlink_url(hl);
                let link = url.as_deref().map(|u| (u, hl.tooltip.as_deref()));
                for run in &hl.runs {
                    convert_run(run, link, &ctx, &mut content);
                }
            },
        }
    }
    // Slack from `Vec`'s minimum capacity: an inline slot is ~100 bytes
    // and most paragraphs hold one span.
    content.shrink_to_fit();
    content
}

fn convert_run(
    run: &crate::docx::Run,
    // The enclosing hyperlink's URL and hover text.
    link: Option<(&str, Option<&str>)>,
    ctx: &RunContext<'_>,
    content: &mut Vec<InlineContent>,
) {
    convert_run_content(run.properties.as_ref(), &run.content, link, ctx, content);
}

/// [`convert_run`] on a run's properties and a slice of its content, so a
/// run can be converted in pieces (around a hard break) without copying it.
fn convert_run_content(
    run_properties: Option<&crate::docx::RunProperties>,
    run_content: &[crate::docx::RunContent],
    // The enclosing hyperlink's URL and hover text.
    link: Option<(&str, Option<&str>)>,
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
            resolved = sheet.effective_run_properties(ctx.paragraph_style_id, run_properties);
            Some(&resolved)
        },
        None => run_properties,
    };
    // `<w:vanish/>` — Word never renders this run at all. Excluding it
    // here (rather than carrying a `hidden` flag into the IR for every
    // renderer to filter separately) keeps plain_text/to_markdown/to_html
    // and the CLI's JSON projection automatically in agreement, instead
    // of risking a 5th instance of this crate's "two renderers disagree"
    // flaw.
    if effective.and_then(|rp| rp.hidden).unwrap_or(false) {
        return;
    }
    let no_properties;
    let eff: &crate::docx::RunProperties = match effective {
        Some(rp) => rp,
        None => {
            no_properties = crate::docx::RunProperties::default();
            &no_properties
        },
    };
    // `w:rFonts` names a face per script and complex-script text has its
    // own size and weight (`w:szCs`, `w:bCs`, `w:iCs`); see
    // `RunProperties::face_for`. Bold and italic are per run, taken from
    // the script of its text; face and size may change inside a run, so
    // each text item is split where they do.
    let run_face =
        eff.face_for(eff.run_script_class(run_content.iter().filter_map(|rc| match rc {
            crate::docx::RunContent::Text(t) => Some(t.as_str()),
            crate::docx::RunContent::FormField(ff) => ff.display_text.as_deref(),
            _ => None,
        })));
    let bold = run_face.bold;
    let italic = run_face.italic;
    // A theme reference (`w:asciiTheme`, the default in every Office
    // template) supersedes the literal face (ECMA-376 Part 1 §17.3.2.26)
    // and names a font in the theme's font scheme; without a theme to
    // resolve it, the literal face is the only name there is.
    let themed_ascii = eff
        .font_theme
        .zip(theme)
        .and_then(|(t, th)| t.resolve(&th.font_scheme));
    let ascii = themed_ascii.as_deref().or(eff.font_name.as_deref());
    let strike = eff.strike.or(eff.dstrike).unwrap_or(false);
    // Propagate `<w:color w:val="RRGGBB"/>` so PDF→DOCX→PDF round-trips
    // preserve coloured text. Resolve it through the document theme:
    // matching only `ColorRef::Rgb` silently dropped every theme-coloured
    // run — and the `w:val` fallback OOXML writes next to `w:themeColor`
    // was discarded at parse time, so even a document with no theme part
    // came out colourless.
    let text_color = eff
        .color
        .as_ref()
        .and_then(|c| c.resolve_opt(theme))
        .map(|rgb| rgb.0);
    // Remaining `w:rPr` toggles. Half of `TextSpan`'s fields used to be
    // permanently empty for DOCX because these were parsed and then never
    // read; underline in particular is the single most common piece of
    // direct formatting after bold/italic.
    let underline = eff.underline.as_ref().map(|u| match u {
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
    let highlight = eff
        .shading_fill
        .as_deref()
        .and_then(hex_to_rgb)
        .or_else(|| eff.highlight.as_deref().and_then(highlight_name_to_rgb));
    let vertical_align = eff.vertical_align.map(|va| match va {
        crate::docx::VerticalAlign::Superscript => VerticalAlign::Superscript,
        crate::docx::VerticalAlign::Subscript => VerticalAlign::Subscript,
        crate::docx::VerticalAlign::Baseline => VerticalAlign::Baseline,
    });
    let all_caps = eff.caps.unwrap_or(false);
    let small_caps = eff.small_caps.unwrap_or(false);
    let char_spacing_half_pt = eff.char_spacing;

    // One span per piece of `text` whose face and size stay the same.
    // The face name reaches `TextSpan.font_name` so the IR→PDF renderer
    // does not fall back to its default font; `<w:sz>` is already in
    // half-points, the IR's own encoding (see
    // `crate::core::units::HalfPoint::from_word_sz`).
    let push_text = |text: &str, content: &mut Vec<InlineContent>| {
        eff.for_each_script_segment(text, ascii, |piece, face| {
            content.push(InlineContent::Text(TextSpan {
                text: piece.to_string(),
                bold,
                italic,
                strikethrough: strike,
                hyperlink: link.map(|(url, _)| url.to_string()),
                hyperlink_tooltip: link.and_then(|(_, tip)| tip).map(str::to_string),
                font_size_half_pt: face
                    .font_size
                    .map(|hp| crate::core::units::HalfPoint::from_word_sz(hp.0).0),
                font_name: face.font_name.map(str::to_string),
                color: text_color,
                underline: underline.clone(),
                highlight,
                vertical_align: vertical_align.clone(),
                all_caps,
                small_caps,
                char_spacing_half_pt,
            }));
        });
    };

    for rc in run_content {
        match rc {
            crate::docx::RunContent::Text(text) => push_text(text, content),
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
            // Text boxes are hoisted to paragraph-sibling `Element::TextBox`
            // nodes by `collect_paragraph_text_boxes`; the inline-content
            // model has no block-container variant.
            crate::docx::RunContent::TextBox(_) => {},
            crate::docx::RunContent::Drawing(_) => {
                // Inline drawings are hoisted to paragraph-sibling
                // `Element::Image` nodes by `collect_paragraph_inline_images`,
                // which carries `alt_text` with them. Also emitting the alt
                // text here as a body-text span made the IR round-trip
                // unbounded: each generation wrote the alt text as a real
                // run *and* re-attached it to the image, so the document
                // grew every time it was read and written back.
            },
            // The citation point in body text — the note *body* is
            // converted separately into `Element::Footnote`/`Endnote`.
            // Carrying the reference mark here is the whole point:
            // before this, to_ir() had the note body but no record of
            // where it was cited.
            crate::docx::RunContent::FootnoteRef(id, _) => {
                content.push(InlineContent::FootnoteRef(FootnoteRef {
                    note_id: *id,
                    marker: None,
                }));
            },
            crate::docx::RunContent::EndnoteRef(id, _) => {
                content.push(InlineContent::EndnoteRef(FootnoteRef {
                    note_id: *id,
                    marker: None,
                }));
            },
            // A comment's anchor: its range start and citation point. The
            // body reaches the IR as the `Element::Endnote` labelled
            // "Comment (…)" with the same id.
            crate::docx::RunContent::CommentRangeStart(id) => {
                content.push(InlineContent::CommentStart(CommentAnchor { comment_id: *id }));
            },
            crate::docx::RunContent::CommentRef(id) => {
                content.push(InlineContent::CommentRef(CommentAnchor { comment_id: *id }));
            },
            crate::docx::RunContent::FormField(ff) => {
                if let Some(text) = &ff.display_text {
                    push_text(text, content);
                }
            },
            // Resolved into a TextBox sibling during from_opc when the
            // reference could be followed (SmartArt, embedded
            // package); an unresolvable one is dropped, matching the
            // type's documented intent.
            crate::docx::RunContent::DeferredPart(_) => {},
        }
    }
}

/// The kind of hard break that terminated a paragraph's inline content.
#[derive(Clone, Copy, PartialEq, Eq)]
enum HardBreak {
    Page,
    Column,
}

/// The paragraph's inline content cut at every hard (page/column) break:
/// `(content before the break, the break)`, with the final segment's
/// break `None`. A paragraph without one is a single segment.
fn split_at_hard_breaks(
    p: &crate::docx::Paragraph,
    doc: &crate::docx::DocxDocument,
) -> Vec<(Vec<InlineContent>, Option<HardBreak>)> {
    let ctx = run_context(p, doc);
    let mut segments: Vec<(Vec<InlineContent>, Option<HardBreak>)> = Vec::new();
    let mut content = Vec::new();
    let convert_run_split =
        |run: &crate::docx::Run,
         url: Option<(&str, Option<&str>)>,
         content: &mut Vec<InlineContent>,
         segments: &mut Vec<(Vec<InlineContent>, Option<HardBreak>)>| {
            // A run holding a hard break is converted around it: the run's
            // pieces before and after the break belong to different
            // segments. The pieces are slices of the run; cloning the whole
            // run to empty it again cost a copy of every run in the
            // document.
            let props = run.properties.as_ref();
            let mut piece_start = 0;
            for (i, rc) in run.content.iter().enumerate() {
                let brk = match rc {
                    crate::docx::RunContent::Break(crate::docx::BreakType::Page) => {
                        Some(HardBreak::Page)
                    },
                    crate::docx::RunContent::Break(crate::docx::BreakType::Column) => {
                        Some(HardBreak::Column)
                    },
                    _ => None,
                };
                if let Some(b) = brk {
                    if piece_start < i {
                        convert_run_content(
                            props,
                            &run.content[piece_start..i],
                            url,
                            &ctx,
                            content,
                        );
                    }
                    segments.push((std::mem::take(content), Some(b)));
                    piece_start = i + 1;
                }
            }
            if piece_start < run.content.len() {
                let piece = &run.content[piece_start..];
                convert_run_content(props, piece, url, &ctx, content);
            }
        };
    for pc in &p.content {
        match pc {
            crate::docx::ParagraphContent::Run(run) => {
                convert_run_split(run, None, &mut content, &mut segments);
            },
            crate::docx::ParagraphContent::Hyperlink(hl) => {
                let url = hyperlink_url(hl);
                let link = url.as_deref().map(|u| (u, hl.tooltip.as_deref()));
                for run in &hl.runs {
                    convert_run_split(run, link, &mut content, &mut segments);
                }
            },
        }
    }
    segments.push((content, None));
    segments
}

// ---------------------------------------------------------------------------
// List conversion
// ---------------------------------------------------------------------------

fn convert_list_group(
    blocks: &[crate::docx::BlockElement],
    i: &mut usize,
    num_id: u32,
    doc: &crate::docx::DocxDocument,
    numbering_counts: &mut std::collections::HashMap<u32, u32>,
    // The first paragraph's effective properties, which the caller already
    // resolved to decide this is a list.
    first_properties: Option<crate::docx::ParagraphProperties>,
) -> Element {
    let mut first_properties = Some(first_properties);
    let mut items = Vec::new();
    let mut is_ordered = false;
    // How many items at the group's own (shallowest) level this group
    // contributes — used to advance `numbering_counts` for
    // interrupted-list continuation.
    let mut top_level_item_count: u32 = 0;
    // The marker style and start value of the *shallowest* level in the
    // group describe the list the IR is about to build. Both used to be
    // resolved and then thrown away, so a list starting at 5 rendered as
    // "1." and `a) b) c)` was indistinguishable from `1. 2. 3.`.
    let mut top_ilvl: Option<u8> = None;
    let mut start_number: Option<u32> = None;
    let mut style: Option<ListStyle> = None;

    let start_index = *i;

    while *i < blocks.len() {
        if let crate::docx::BlockElement::Paragraph(p) = &blocks[*i] {
            // Membership must be decided on the *effective* properties, the
            // same set the caller used to decide this run is a list.
            // Testing the direct `w:pPr` here instead meant a paragraph
            // whose `w:numPr` comes from its style matched at the call site
            // and not here: the group consumed nothing, `*i` never advanced,
            // and the caller looped forever appending empty lists until the
            // process was killed.
            let eff = match first_properties.take() {
                Some(eff) => eff,
                None => effective_paragraph_props(p, doc),
            };
            if let Some(nr) = eff.as_ref().and_then(|pp| pp.numbering_ref.as_ref()) {
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
                            style = number_format_to_list_style(&level.format)
                                .map(|s| bullet_glyph_style(s, &level.level_text));
                            // Honour this instance's own `<w:startOverride>`
                            // when present, falling back to the
                            // abstract level's own `<w:start>`. `w:start`
                            // defaults to 1; only report an explicit
                            // non-default so renderers that ignore the
                            // field are not silently contradicted.
                            let effective_start = numbering
                                .resolve_start(nr.num_id, nr.ilvl)
                                .unwrap_or(level.start);
                            start_number = (effective_start != 1).then_some(effective_start);
                        }
                    }
                }

                if top_ilvl == Some(nr.ilvl) {
                    top_level_item_count += 1;
                }
                items.push((nr.ilvl, convert_paragraph_inline(p, doc)));
                *i += 1;
                continue;
            }
        }
        break;
    }

    // Guarantee forward progress. The caller advances only through this
    // function, so a group that consumes nothing is an infinite loop; make
    // that impossible here rather than relying on the two membership tests
    // agreeing forever.
    if *i == start_index {
        *i += 1;
    }

    // A numId seen earlier in this same block sequence, with no explicit
    // `w:start`/`w:startOverride` this time, continues counting from where
    // the previous group left off rather than restarting at 1.
    if start_number.is_none() {
        if let Some(&prev_count) = numbering_counts.get(&num_id) {
            start_number = Some(prev_count + 1);
        }
    }
    let resumed_from = start_number.unwrap_or(1);
    numbering_counts.insert(num_id, resumed_from + top_level_item_count.saturating_sub(1));

    // Build nested list structure from flat (ilvl, content) pairs
    let mut list = crate::ir::build_nested_list(is_ordered, &items, 0);
    list.start_number = start_number;
    list.style = style;
    Element::List(list)
}

/// Refine a bullet level's style by the glyph its `w:lvlText` draws: the
/// writer encodes `Square`/`Circle`/`Dash` as ▪/○/– and every glyph used
/// to read back as a plain bullet. Word's own gallery bullets are private
/// use code points of the Symbol/Wingdings fonts (U+F0A7 is Wingdings'
/// square, U+F0B7 Symbol's round bullet) and Courier New "o" is its
/// open-circle bullet.
fn bullet_glyph_style(style: ListStyle, level_text: &str) -> ListStyle {
    if style != ListStyle::Bullet {
        return style;
    }
    match level_text.trim() {
        "\u{25AA}" | "\u{25A0}" | "\u{25FE}" | "\u{25FC}" | "\u{F0A7}" | "\u{F06E}" => {
            ListStyle::Square
        },
        "\u{25CB}" | "\u{25E6}" | "o" | "\u{F06F}" => ListStyle::Circle,
        "\u{2013}" | "\u{2014}" | "-" | "\u{2212}" => ListStyle::Dash,
        _ => ListStyle::Bullet,
    }
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

/// Upper bound for a single `w:gridSpan`. Word's own table limit is 63
/// columns; this leaves generous headroom while keeping the value bounded.
const MAX_GRID_SPAN: u32 = 1_000;

/// A cell's `w:gridSpan`, clamped to `1..=MAX_GRID_SPAN`.
fn cell_grid_span(cell: &crate::docx::TableCell) -> u32 {
    cell.properties
        .as_ref()
        .and_then(|p| p.grid_span)
        .unwrap_or(1)
        .clamp(1, MAX_GRID_SPAN)
}

fn convert_table(table: &crate::docx::Table, doc: &crate::docx::DocxDocument) -> Element {
    // A table cell can hold another table (`convert_block_elements` ->
    // `convert_table` -> `convert_block_elements` -> ...), so this
    // recurses on whatever it's given — including a tree the XML parser
    // already bounded to `MAX_NESTING_DEPTH`. That bound protects parsing
    // (which runs on its own larger stack), but to_ir() runs on whatever
    // stack the caller has, and re-walking a tree that deep overflowed it
    // (the same defect class as an unguarded XML parse, one
    // layer downstream of it).
    let Some(_depth) = crate::core::xml::DepthGuard::enter() else {
        return Element::Table(Table {
            rows: vec![TableRow {
                cells: vec![TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: vec![InlineContent::Text(TextSpan::plain(format!(
                            "[nested table deeper than {} levels not shown — \
                             document truncated]",
                            crate::core::xml::MAX_NESTING_DEPTH
                        )))],
                        ..Default::default()
                    })],
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        });
    };

    // First pass: compute row_span from vMerge patterns.
    //
    // `w:gridSpan` is attacker-controlled and parsed as an unbounded u32,
    // so every span is clamped (`cell_grid_span`). The vMerge pass works
    // per *cell*, not per grid position: a dense `rows x columns` grid
    // made every narrow row cost the width of the table's widest row, and
    // one wide row over many ordinary ones was quadratic in the input.
    let num_rows = table.rows.len();
    // Starting grid column of every cell, row by row (ascending, so a
    // lookup by column is a binary search).
    let starts: Vec<Vec<usize>> = table
        .rows
        .iter()
        .map(|r| {
            let mut col = 0usize;
            r.cells
                .iter()
                .map(|c| {
                    let start = col;
                    col = col.saturating_add(cell_grid_span(c) as usize);
                    start
                })
                .collect()
        })
        .collect();
    let vmerge_at = |row: usize, col: usize| {
        let i = starts[row].binary_search(&col).ok()?;
        table.rows[row].cells[i]
            .properties
            .as_ref()
            .and_then(|p| p.vertical_merge)
    };
    let mut row_spans: Vec<Vec<u32>> = table.rows.iter().map(|r| vec![1; r.cells.len()]).collect();
    for (row, cells) in table.rows.iter().enumerate() {
        for (i, cell) in cells.cells.iter().enumerate() {
            let vmerge = cell.properties.as_ref().and_then(|p| p.vertical_merge);
            if !matches!(vmerge, Some(crate::docx::table::MergeType::Restart)) {
                continue;
            }
            // Count the continuation cells below. Each continuation cell
            // is reached from at most one restart, so this is linear.
            let col = starts[row][i];
            let mut span = 1u32;
            let mut next = row + 1;
            while next < num_rows
                && matches!(vmerge_at(next, col), Some(crate::docx::table::MergeType::Continue))
            {
                span = span.saturating_add(1);
                next += 1;
            }
            row_spans[row][i] = span;
        }
    }

    // The table's style: borders it does not set itself, and the
    // conditional (header, total, first/last column, banded) shading its
    // `w:tblLook` switches on.
    let tp = table.properties.as_ref();
    let style = tp
        .and_then(|p| p.style_id.as_deref())
        .zip(doc.styles.as_ref())
        .map(|(id, sheet)| sheet.resolve_table_style(id))
        .unwrap_or_default();
    let grid_width = starts
        .iter()
        .zip(&table.rows)
        .map(|(s, r)| match (s.last(), r.cells.last()) {
            (Some(&start), Some(cell)) => start.saturating_add(cell_grid_span(cell) as usize),
            _ => 0,
        })
        .max()
        .unwrap_or(0);
    let layout = TableStyleLayout {
        look: tp.and_then(|p| p.look),
        num_rows,
        grid_width,
        row_band: tp
            .and_then(|p| p.row_band_size)
            .or(style.row_band_size)
            .unwrap_or(1) as usize,
        col_band: tp
            .and_then(|p| p.col_band_size)
            .or(style.col_band_size)
            .unwrap_or(1) as usize,
    };

    let mut ir_rows = Vec::new();
    for (row_idx, row) in table.rows.iter().enumerate() {
        let rp = row.properties.as_ref();
        let is_header = rp.is_some_and(|p| p.is_header);

        let mut ir_cells = Vec::new();

        for (cell_idx, cell) in row.cells.iter().enumerate() {
            // Clamped here, where the IR cell is built, so every consumer is
            // covered: ir_render sizes a grid from the summed col_spans, and
            // the DOCX writer loops over them. An unbounded value from the
            // file reached both.
            let col_span = cell_grid_span(cell);

            // Skip vMerge continue cells
            let is_continue = cell
                .properties
                .as_ref()
                .and_then(|p| p.vertical_merge)
                .is_some_and(|m| matches!(m, crate::docx::table::MergeType::Continue));

            // A cell deleted via tracked changes (`w:cellDel`) is excluded
            // from the accepted view — same policy already applied to
            // run-level `w:del`. Its grid position is still accounted for
            // (spans are resolved per cell above), exactly like a
            // vMerge-continue cell's, so later real cells in the row don't
            // shift into the wrong column.
            let is_deleted = cell.properties.as_ref().is_some_and(|p| p.deleted);
            if is_continue || is_deleted {
                continue;
            }

            let row_span = row_spans[row_idx][cell_idx];

            let mut cell_elements = Vec::new();
            convert_block_elements(&cell.content, &mut cell_elements, doc);
            // One paragraph per cell is the norm; an `Element` slot is
            // ~270 bytes and a fresh `Vec` reserves four of them.
            cell_elements.shrink_to_fit();

            // The writer already emits `<w:jc>` inside a cell's paragraph
            // from `TableCell::text_align` (`docx/write.rs`), but nothing
            // read it back — take the first paragraph's alignment as the
            // cell's own, the same convention the writer uses when it
            // stamps every paragraph in the cell with this one value.
            let text_align = cell_elements.iter().find_map(|e| match e {
                Element::Paragraph(p) => p.alignment.clone(),
                _ => None,
            });

            let cp = cell.properties.as_ref();
            ir_cells.push(TableCell {
                content: cell_elements,
                col_span,
                row_span,
                text_align,
                // Direct cell shading wins over the table style's; a direct
                // `w:shd`, even `nil`, is the cell's own choice.
                background_color: match cp.and_then(|p| p.shading.as_deref()) {
                    Some(sh) => shading_rgb(sh),
                    None => {
                        let start = starts[row_idx][cell_idx];
                        let end = start.saturating_add(col_span as usize);
                        layout
                            .regions(row_idx, start, end)
                            .iter()
                            .rev()
                            .find_map(|kind| {
                                style
                                    .conditionals
                                    .iter()
                                    .find(|c| c.kind == *kind)
                                    .and_then(|c| c.cell_properties.as_ref())
                                    .and_then(|p| p.shading.as_deref())
                                    .and_then(shading_rgb)
                            })
                    },
                },
                border: cp
                    .and_then(|p| p.borders.as_deref())
                    .map(table_borders_to_ir),
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
        }

        ir_rows.push(TableRow {
            cells: ir_cells,
            is_header,
            height_twips: rp.and_then(|p| p.height).map(|h| h.max(0) as u32),
            height_rule: rp.and_then(|p| p.height_rule).map(|r| match r {
                crate::docx::table::RowHeightRule::AtLeast => RowHeightRule::AtLeast,
                crate::docx::table::RowHeightRule::Exact => RowHeightRule::Exact,
                crate::docx::table::RowHeightRule::Auto => RowHeightRule::Auto,
            }),
            allow_break: !rp.is_some_and(|p| p.cant_split),
            repeat_as_header: is_header,
        });
    }

    Element::Table(Table {
        rows: ir_rows,
        column_widths_twips: table.grid.iter().map(|t| t.0.max(0) as u32).collect(),
        border: tp
            .and_then(|p| p.borders.as_ref())
            .or(style.borders)
            .map(table_borders_to_ir),
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

/// Where a cell sits relative to the regions a table style formats.
struct TableStyleLayout {
    /// `None` when the table has no `w:tblLook`: only whole-table
    /// formatting applies, rather than guessing which regions are on.
    look: Option<crate::docx::table::TableLook>,
    num_rows: usize,
    grid_width: usize,
    row_band: usize,
    col_band: usize,
}

impl TableStyleLayout {
    /// The `w:tblStylePr` regions covering the cell at `row` spanning grid
    /// columns `start..end`, in increasing precedence (ECMA-376 §17.7.6:
    /// whole table, banded columns, banded rows, first/last column,
    /// first/last row, corner cells).
    fn regions(&self, row: usize, start: usize, end: usize) -> Vec<&'static str> {
        let mut out = vec!["wholeTable"];
        let Some(look) = self.look else {
            return out;
        };
        let first_row = look.first_row && row == 0;
        let last_row = look.last_row && row + 1 == self.num_rows;
        let first_col = look.first_column && start == 0;
        let last_col = look.last_column && end >= self.grid_width;
        if !look.no_v_band && !first_col && !last_col {
            let data_col = start.saturating_sub(usize::from(look.first_column));
            out.push(if (data_col / self.col_band.max(1)).is_multiple_of(2) {
                "band1Vert"
            } else {
                "band2Vert"
            });
        }
        if !look.no_h_band && !first_row && !last_row {
            let data_row = row.saturating_sub(usize::from(look.first_row));
            out.push(if (data_row / self.row_band.max(1)).is_multiple_of(2) {
                "band1Horz"
            } else {
                "band2Horz"
            });
        }
        if first_col {
            out.push("firstCol");
        }
        if last_col {
            out.push("lastCol");
        }
        if first_row {
            out.push("firstRow");
        }
        if last_row {
            out.push("lastRow");
        }
        match (first_row, last_row, first_col, last_col) {
            (true, _, true, _) => out.push("nwCell"),
            (true, _, _, true) => out.push("neCell"),
            (_, true, true, _) => out.push("swCell"),
            (_, true, _, true) => out.push("seCell"),
            _ => {},
        }
        out
    }
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

// Also handle images at the block level by scanning for drawings in paragraphs
impl From<&crate::docx::DrawingInfo> for Image {
    fn from(d: &crate::docx::DrawingInfo) -> Self {
        Image {
            alt_text: d.description.clone(),
            decorative: d.decorative,
            ..Default::default()
        }
    }
}
