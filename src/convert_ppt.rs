use crate::format::DocumentFormat;
use crate::ir::*;
use crate::ppt::{CharFormatSpan, ParaFormatSpan, TextRun, TextType};

/// Split `text` into paragraphs at bare `\r` / `\n` delimiters.
///
/// [MS-PPT] text bodies use a lone `\r` (0x0D) as the paragraph separator
/// *inside a single `TextBytesAtom`/`TextCharsAtom`* — see the "Paragraph
/// Formatting" section's own worked example ("a sunny day\rthe blue
/// sky\rsome green grass" is ONE atom containing three paragraphs). Rust's
/// `str::lines()` only splits on `\n` or `\r\n`, so a lone `\r` was never
/// being treated as a paragraph break at all: the whole multi-bullet body
/// was reaching the IR as a single paragraph/line with literal `\r`
/// characters embedded in it. Fixed here (found while implementing #254,
/// since `TextPFRun` paragraph boundaries are defined in terms of these
/// same delimiters) and filed as its own issue since the root cause
/// (`.lines()` on raw PPT text) is distinct from "no formatting is ever
/// read".
///
/// Returns `(start_char, end_char_excl_delimiter)` pairs — character
/// index ranges into `text.chars()`, *not* including the delimiter itself
/// — so callers can slice both the rendered text and any character-range
/// formatting spans (which are indexed the same way) consistently.
fn split_paragraphs(text: &str) -> Vec<(usize, usize)> {
    let chars: Vec<char> = text.chars().collect();
    let mut result = Vec::new();
    let mut start = 0usize;
    for (i, &c) in chars.iter().enumerate() {
        if c == '\r' || c == '\n' {
            result.push((start, i));
            start = i + 1;
        }
    }
    if start < chars.len() || result.is_empty() {
        result.push((start, chars.len()));
    }
    result
}

/// Build the `InlineContent` spans for one paragraph's `[start, end)`
/// character range, splitting further at any `char_formats` boundary that
/// falls inside the range and applying that run's real formatting.
///
/// Leading/trailing whitespace within the range is trimmed (matching the
/// old whole-text-trim behavior for the common case of no embedded
/// formatting, while keeping character offsets — and therefore formatting
/// span alignment — correct for the untrimmed middle).
fn spans_for_range(
    text_chars: &[char],
    start: usize,
    end: usize,
    char_formats: &[CharFormatSpan],
    hyperlink: Option<&str>,
) -> Vec<InlineContent> {
    let end = end.min(text_chars.len());
    let start = start.min(end);
    let mut s = start;
    let mut e = end;
    while s < e && text_chars[s].is_whitespace() {
        s += 1;
    }
    while e > s && text_chars[e - 1].is_whitespace() {
        e -= 1;
    }
    if s >= e {
        return Vec::new();
    }

    let mut cuts = std::collections::BTreeSet::new();
    cuts.insert(s);
    cuts.insert(e);
    for f in char_formats {
        if f.start > s && f.start < e {
            cuts.insert(f.start);
        }
        if f.end > s && f.end < e {
            cuts.insert(f.end);
        }
    }
    let cuts: Vec<usize> = cuts.into_iter().collect();

    let mut spans = Vec::with_capacity(cuts.len().saturating_sub(1));
    for w in cuts.windows(2) {
        let (a, b) = (w[0], w[1]);
        if a >= b {
            continue;
        }
        let seg: String = text_chars[a..b].iter().collect();
        let fmt = char_formats.iter().find(|f| f.start <= a && b <= f.end);
        let mut span = TextSpan::plain(seg);
        span.hyperlink = hyperlink.map(str::to_string);
        if let Some(f) = fmt {
            if let Some(bold) = f.format.bold {
                span.bold = bold;
            }
            if let Some(italic) = f.format.italic {
                span.italic = italic;
            }
            if let Some(underline) = f.format.underline {
                span.underline = if underline { Some(UnderlineStyle::Single) } else { None };
            }
            if let Some(size) = f.format.font_size {
                span.font_size_half_pt = Some((size.max(0) as u32) * 2);
            }
            if let Some(color) = f.format.color {
                span.color = Some(color);
            }
            if let Some(position) = f.format.position {
                span.vertical_align = match position.cmp(&0) {
                    std::cmp::Ordering::Greater => Some(VerticalAlign::Superscript),
                    std::cmp::Ordering::Less => Some(VerticalAlign::Subscript),
                    std::cmp::Ordering::Equal => None,
                };
            }
        }
        spans.push(InlineContent::Text(span));
    }
    spans
}

/// Map [MS-PPT]'s raw `TextAlignmentEnum` value to the IR's
/// `ParagraphAlignment`. `Distributed` (4) and `ThaiDistributed` (5) both
/// map to `Distribute`; `JustifyLow` (6, kashida justification) maps to
/// `Justify` as the closest IR equivalent — the IR has no kashida concept.
fn map_alignment(v: u16) -> Option<ParagraphAlignment> {
    match v {
        0 => Some(ParagraphAlignment::Left),
        1 => Some(ParagraphAlignment::Center),
        2 => Some(ParagraphAlignment::Right),
        3 | 6 => Some(ParagraphAlignment::Justify),
        4 | 5 => Some(ParagraphAlignment::Distribute),
        _ => None,
    }
}

/// Resolve the alignment that applies at character offset `at`, from
/// whichever `para_formats` span (if any) contains it.
fn alignment_at(para_formats: &[ParaFormatSpan], at: usize) -> Option<ParagraphAlignment> {
    para_formats
        .iter()
        .find(|p| p.start <= at && at < p.end.max(p.start + 1))
        .and_then(|p| p.format.alignment)
        .and_then(map_alignment)
}

/// Build a table cell's block content from its shape's own text runs,
/// reusing the same paragraph/formatting-span logic as ordinary body text
/// (issue #255).
fn table_cell_content(runs: &[TextRun]) -> Vec<Element> {
    let mut elements = Vec::new();
    for run in runs {
        if run.text.trim().is_empty() {
            continue;
        }
        let text_chars: Vec<char> = run.text.chars().collect();
        for &(start, end) in &split_paragraphs(&run.text) {
            let content = spans_for_range(
                &text_chars,
                start,
                end,
                &run.char_formats,
                run.hyperlink.as_deref(),
            );
            if !content.is_empty() {
                elements.push(Element::Paragraph(Paragraph {
                    content,
                    alignment: alignment_at(&run.para_formats, start),
                    ..Default::default()
                }));
            }
        }
    }
    elements
}

/// Convert a reconstructed grid-of-shapes table (issue #255) into
/// `Element::Table`.
fn table_block_to_element(table: &crate::ppt::TableBlock) -> Element {
    let rows = table
        .rows
        .iter()
        .map(|row| TableRow {
            cells: row
                .iter()
                .map(|cell_runs| TableCell {
                    content: table_cell_content(cell_runs),
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        })
        .collect();
    Element::Table(Table { rows, ..Default::default() })
}

pub(crate) fn ppt_to_ir(doc: &crate::ppt::PptDocument) -> DocumentIR {
    let mut sections = Vec::new();

    for (slide_idx, slide) in doc.slides.iter().enumerate() {
        let mut elements = Vec::new();
        let mut slide_title: Option<String> = None;
        // Presenter-only text, kept out of `elements` (which is "what the
        // audience sees") and routed to the dedicated field instead — the
        // PPTX side of this was already fixed in #203; #238 is the same
        // defect on the legacy binary .ppt path, which had never been
        // ported to route TextType::Notes there at all.
        let mut notes_lines: Vec<&str> = Vec::new();

        for run in &slide.text_runs {
            if run.text.trim().is_empty() {
                continue;
            }
            let text_chars: Vec<char> = run.text.chars().collect();
            let paragraphs = split_paragraphs(&run.text);

            match run.text_type {
                TextType::Title | TextType::CenterTitle => {
                    // A heading is rendered as one block; use the first
                    // non-empty paragraph's real formatting when direct
                    // formatting was read, falling back to the old
                    // synthetic "always bold" only when no
                    // `StyleTextPropAtom` was present at all (e.g. the
                    // many hand-built test fixtures that never set
                    // `char_formats`).
                    let mut content = Vec::new();
                    for &(start, end) in &paragraphs {
                        content.extend(spans_for_range(
                            &text_chars,
                            start,
                            end,
                            &run.char_formats,
                            run.hyperlink.as_deref(),
                        ));
                    }
                    if content.is_empty() {
                        continue;
                    }
                    if run.char_formats.is_empty() {
                        for c in &mut content {
                            if let InlineContent::Text(t) = c {
                                t.bold = true;
                            }
                        }
                    }
                    if slide_title.is_none() {
                        let joined: String = content
                            .iter()
                            .filter_map(|c| match c {
                                InlineContent::Text(t) => Some(t.text.as_str()),
                                _ => None,
                            })
                            .collect::<Vec<_>>()
                            .join("");
                        slide_title = Some(joined);
                    }
                    elements.push(Element::Heading(Heading {
                        level: 1,
                        content,
                        alignment: alignment_at(&run.para_formats, paragraphs.first().map_or(0, |p| p.0)),
                        ..Default::default()
                    }));
                },
                TextType::Body | TextType::HalfBody | TextType::QuarterBody => {
                    for &(start, end) in &paragraphs {
                        let content =
                            spans_for_range(&text_chars, start, end, &run.char_formats, run.hyperlink.as_deref());
                        if !content.is_empty() {
                            elements.push(Element::Paragraph(Paragraph {
                                content,
                                alignment: alignment_at(&run.para_formats, start),
                                ..Default::default()
                            }));
                        }
                    }
                },
                TextType::Notes => {
                    notes_lines.push(run.text.trim());
                },
                _ => {
                    let mut content = Vec::new();
                    for &(start, end) in &paragraphs {
                        content.extend(spans_for_range(
                            &text_chars,
                            start,
                            end,
                            &run.char_formats,
                            run.hyperlink.as_deref(),
                        ));
                    }
                    if !content.is_empty() {
                        elements.push(Element::Paragraph(Paragraph {
                            content,
                            alignment: alignment_at(&run.para_formats, paragraphs.first().map_or(0, |p| p.0)),
                            ..Default::default()
                        }));
                    }
                },
            }
        }

        // Reconstructed grid-of-shapes tables (issue #255) always land
        // after the slide's ordinary text — the binary format has no
        // single reading-order concept spanning both, so this is a
        // deliberate simplification rather than a claim of true order.
        for table in &slide.tables {
            elements.push(table_block_to_element(table));
        }

        let title = slide_title.unwrap_or_else(|| format!("Slide {}", slide_idx + 1));
        let speaker_notes = if notes_lines.is_empty() {
            None
        } else {
            Some(notes_lines.join("\n").trim().to_string()).filter(|s| !s.is_empty())
        };

        sections.push(Section {
            title: Some(title),
            elements,
            speaker_notes,
            ..Default::default()
        });
    }

    // Extracted pictures never reached the IR, so every image in a legacy
    // deck was silently dropped on conversion.
    crate::convert_xls::append_legacy_images(&mut sections, doc.images());

    // The deck's own declared title (from `\x05SummaryInformation`) beats
    // the first slide's own title — a slide title is not a document
    // title, it's just the only thing that was ever there to fall back
    // to (issue #244).
    let summary = doc.summary_properties();
    let title = summary
        .and_then(|s| s.title.clone())
        .filter(|t| !t.is_empty())
        .or_else(|| sections.first().and_then(|s| s.title.clone()));

    DocumentIR {
        metadata: Metadata {
            format: DocumentFormat::Ppt,
            title,
            author: summary.and_then(|s| s.author.clone()).filter(|s| !s.is_empty()),
            subject: summary.and_then(|s| s.subject.clone()).filter(|s| !s.is_empty()),
            keywords: summary
                .and_then(|s| s.keywords.as_deref())
                .map(crate::convert_docx::split_keywords)
                .unwrap_or_default(),
            description: summary.and_then(|s| s.comments.clone()).filter(|s| !s.is_empty()),
            created: summary.and_then(|s| s.created.clone()),
            modified: summary.and_then(|s| s.modified.clone()),
            has_macros: doc.has_macros(),
            ..Default::default()
        },
        sections,
        defined_names: Vec::new(),
    }
}
