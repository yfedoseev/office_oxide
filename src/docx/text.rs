use super::DocxDocument;
use super::document::BlockElement;
use super::hyperlink::HyperlinkTarget;
use super::image::DrawingInfo;
use super::numbering::NumberingDefinitions;
use super::paragraph::{BreakType, ParagraphContent, Run, RunContent};
use super::styles::StyleSheet;
use super::table::Table;

// ---------------------------------------------------------------------------
// Plain text extraction
// ---------------------------------------------------------------------------

impl DocxDocument {
    /// Extract all text as a plain string. Paragraphs are separated by newlines.
    ///
    /// Headers and footers are included, matching `to_markdown` and the IR
    /// renderers. All three used to disagree, so a consumer's word count
    /// changed depending on which method they called.
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        for hf in self.headers_footers.iter().filter(|h| h.is_header) {
            plain_text_blocks(&hf.content, &mut out);
        }
        plain_text_blocks(&self.body.elements, &mut out);
        for hf in self.headers_footers.iter().filter(|h| !h.is_header) {
            plain_text_blocks(&hf.content, &mut out);
        }
        // Trim trailing newlines
        while out.ends_with('\n') {
            out.pop();
        }
        out
    }

    /// Convert the document to Markdown.
    ///
    /// Includes headers and footers around the body so a downstream
    /// renderer (PDF, HTML, search index) sees the full visible content
    /// of every page. Without this, simple-but-meaningful artefacts like
    /// `My header` / `My footer` are silently dropped.
    pub fn to_markdown(&self) -> String {
        let mut out = String::new();
        let ctx = MarkdownCtx {
            styles: self.styles.as_ref(),
            numbering: self.numbering.as_ref(),
        };

        let (header_texts, footer_texts) = split_headers_footers(self, &ctx);
        for h in &header_texts {
            out.push_str(h);
            out.push_str("\n\n");
        }

        markdown_blocks(&self.body.elements, &ctx, &mut out, 0);

        for f in &footer_texts {
            if !out.ends_with("\n\n") {
                out.push_str("\n\n");
            }
            out.push_str(f);
            out.push('\n');
        }

        while out.ends_with('\n') {
            out.pop();
        }
        out
    }
}

/// Split parsed `HeaderFooter` entries into headers vs footers and
/// return them as deduplicated markdown-string vectors. Role is read
/// directly from each entry's `is_header` field (set at parse time),
/// so this is correct regardless of how many sections the document
/// has or how the entries are interleaved.
fn split_headers_footers(doc: &DocxDocument, ctx: &MarkdownCtx) -> (Vec<String>, Vec<String>) {
    let mut headers: Vec<String> = Vec::new();
    let mut footers: Vec<String> = Vec::new();
    let mut header_seen: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut footer_seen: std::collections::HashSet<String> = std::collections::HashSet::new();

    for hf in &doc.headers_footers {
        let mut buf = String::new();
        markdown_blocks(&hf.content, ctx, &mut buf, 0);
        let t = buf.trim().to_string();
        if t.is_empty() {
            continue;
        }
        if hf.is_header {
            if header_seen.insert(t.clone()) {
                headers.push(t);
            }
        } else if footer_seen.insert(t.clone()) {
            footers.push(t);
        }
    }
    (headers, footers)
}

fn plain_text_blocks(elements: &[BlockElement], out: &mut String) {
    for elem in elements {
        match elem {
            BlockElement::Paragraph(p) => {
                for content in &p.content {
                    match content {
                        ParagraphContent::Run(run) => plain_text_run(run, out),
                        ParagraphContent::Hyperlink(hl) => {
                            for run in &hl.runs {
                                plain_text_run(run, out);
                            }
                        },
                    }
                }
                out.push('\n');
            },
            BlockElement::Table(table) => {
                plain_text_table(table, out);
            },
        }
    }
}

fn plain_text_run(run: &Run, out: &mut String) {
    for content in &run.content {
        match content {
            RunContent::Text(text) => out.push_str(text),
            RunContent::Break(BreakType::Line) => out.push('\n'),
            RunContent::Break(BreakType::Page | BreakType::Column) => out.push('\n'),
            RunContent::Tab => out.push('\t'),
            RunContent::Drawing(_) => {},
            // Text-box prose is document content — in some real files it is
            // most of the document (issue #102). It is *block* content, so
            // it must be separated from the surrounding run: pasting it in
            // bare fused the last word of a text box to the first word after
            // it (`Linz` + `ANTRAG` -> `LinzANTRAG`).
            RunContent::TextBox(blocks) => {
                let mut inner = String::new();
                plain_text_blocks(blocks, &mut inner);
                let inner = inner.trim();
                if !inner.is_empty() {
                    if !out.is_empty() && !out.ends_with(['\n', ' ', '\t']) {
                        out.push('\n');
                    }
                    out.push_str(inner);
                    out.push('\n');
                }
            },
        }
    }
}

fn plain_text_table(table: &Table, out: &mut String) {
    for row in &table.rows {
        for (i, cell) in row.cells.iter().enumerate() {
            if i > 0 {
                out.push('\t');
            }
            let mut cell_text = String::new();
            plain_text_blocks(&cell.content, &mut cell_text);
            // Replace internal newlines with spaces for table cell text
            out.push_str(&cell_text.trim_end_matches('\n').replace('\n', " "));
        }
        out.push('\n');
    }
}

// ---------------------------------------------------------------------------
// Markdown extraction
// ---------------------------------------------------------------------------

struct MarkdownCtx<'a> {
    styles: Option<&'a StyleSheet>,
    numbering: Option<&'a NumberingDefinitions>,
}

fn markdown_blocks(elements: &[BlockElement], ctx: &MarkdownCtx, out: &mut String, _depth: usize) {
    for elem in elements {
        match elem {
            BlockElement::Paragraph(p) => {
                // Determine heading level from outline_level or style
                let heading_level = p
                    .properties
                    .as_ref()
                    .and_then(|pp| {
                        pp.outline_level.or_else(|| {
                            pp.style_id
                                .as_ref()
                                .and_then(|sid| ctx.styles?.resolve_outline_level(sid))
                        })
                    })
                    .map(|lvl| (lvl as usize) + 1);

                // Check for numbering
                let list_prefix = p.properties.as_ref().and_then(|pp| {
                    let nr = pp.numbering_ref.as_ref()?;
                    let numbering = ctx.numbering?;
                    let level = numbering.resolve_level(nr.num_id, nr.ilvl)?;
                    let indent = "  ".repeat(nr.ilvl as usize);
                    use super::numbering::NumberFormat;
                    let marker = match &level.format {
                        NumberFormat::Bullet => "- ".to_string(),
                        NumberFormat::Decimal => format!("{}. ", level.start),
                        NumberFormat::LowerLetter => format!("{}. ", level.start),
                        NumberFormat::UpperLetter => format!("{}. ", level.start),
                        NumberFormat::LowerRoman => format!("{}. ", level.start),
                        NumberFormat::UpperRoman => format!("{}. ", level.start),
                        NumberFormat::None => String::new(),
                        NumberFormat::Other(_) => "- ".to_string(),
                    };
                    Some(format!("{indent}{marker}"))
                });

                // Heading prefix
                if let Some(level) = heading_level {
                    // One definition of the valid heading range, shared with the IR
                    // renderers: `min(9)` emitted up to nine `#`, which no
                    // markdown reader treats as a heading at all.
                    let hashes = "#".repeat(level.clamp(1, 6));
                    out.push_str(&hashes);
                    out.push(' ');
                } else if let Some(ref prefix) = list_prefix {
                    out.push_str(prefix);
                }

                // Render paragraph content with inline formatting.
                // Consecutive runs sharing formatting are merged, because
                // Word splits one visually-bold phrase into several runs
                // routinely and wrapping each separately emits `****`,
                // which CommonMark reads as four literal asterisks.
                let mut pending: Option<(RunStyle, String)> = None;
                for content in &p.content {
                    match content {
                        ParagraphContent::Run(run) => {
                            let style = RunStyle::of(run);
                            let mut text = String::new();
                            markdown_run_text(run, ctx, &mut text);
                            if text.is_empty() {
                                continue;
                            }
                            match pending.as_mut() {
                                Some((cur, buf)) if *cur == style => buf.push_str(&text),
                                _ => {
                                    flush_run(&mut pending, out);
                                    pending = Some((style, text));
                                },
                            }
                        },
                        ParagraphContent::Hyperlink(hl) => {
                            flush_run(&mut pending, out);
                            let text = runs_to_plain_text(&hl.runs);
                            match &hl.target {
                                HyperlinkTarget::External(url) => {
                                    out.push('[');
                                    out.push_str(&text);
                                    out.push_str("](");
                                    out.push_str(url);
                                    out.push(')');
                                },
                                HyperlinkTarget::Internal(anchor) => {
                                    out.push('[');
                                    out.push_str(&text);
                                    out.push_str("](#");
                                    out.push_str(anchor);
                                    out.push(')');
                                },
                            }
                        },
                    }
                }
                flush_run(&mut pending, out);
                out.push('\n');

                // Add extra newline after headings for readability
                if heading_level.is_some() {
                    out.push('\n');
                }
            },
            BlockElement::Table(table) => {
                markdown_table(table, ctx, out);
            },
        }
    }
}

/// The run formatting that decides which markdown delimiters apply.
#[derive(PartialEq)]
struct RunStyle {
    bold: bool,
    italic: bool,
    strike: bool,
    vertical_align: Option<super::formatting::VerticalAlign>,
}

impl RunStyle {
    fn of(run: &Run) -> Self {
        let rp = run.properties.as_ref();
        Self {
            bold: rp.and_then(|rp| rp.bold).unwrap_or(false),
            italic: rp.and_then(|rp| rp.italic).unwrap_or(false),
            strike: rp.and_then(|rp| rp.strike.or(rp.dstrike)).unwrap_or(false),
            vertical_align: rp.and_then(|rp| rp.vertical_align),
        }
    }
}

/// Emit an accumulated run group with one set of delimiters.
fn flush_run(pending: &mut Option<(RunStyle, String)>, out: &mut String) {
    let Some((style, text)) = pending.take() else {
        return;
    };
    if text.is_empty() {
        return;
    }
    // Leading/trailing whitespace must sit outside the delimiters:
    // CommonMark does not open emphasis on `** text**`. An all-whitespace
    // run has no core to emphasise — and computing the two spans
    // independently makes them overlap, which inverts the slice range and
    // panics, so take the trailing span from what is left after the
    // leading one.
    let core_str = text.trim();
    if core_str.is_empty() {
        out.push_str(&text);
        return;
    }
    let lead_len = text.len() - text.trim_start().len();
    let lead = &text[..lead_len];
    let trail = &text[lead_len + core_str.len()..];
    let core = core_str;

    let mut body = core.to_string();
    // Super/subscript have no markdown syntax; HTML is the conventional
    // fallback. `vertAlign` was previously dropped entirely here.
    match style.vertical_align {
        Some(super::formatting::VerticalAlign::Superscript) => {
            body = format!("<sup>{body}</sup>");
        },
        Some(super::formatting::VerticalAlign::Subscript) => {
            body = format!("<sub>{body}</sub>");
        },
        _ => {},
    }
    if style.strike {
        body = format!("~~{body}~~");
    }
    if style.bold && style.italic {
        body = format!("***{body}***");
    } else if style.bold {
        body = format!("**{body}**");
    } else if style.italic {
        body = format!("*{body}*");
    }
    out.push_str(lead);
    out.push_str(&body);
    out.push_str(trail);
}

/// Collect a run's text content (no emphasis delimiters).
fn markdown_run_text(run: &Run, ctx: &MarkdownCtx, text: &mut String) {
    for content in &run.content {
        match content {
            RunContent::Text(t) => text.push_str(t),
            RunContent::Break(BreakType::Line) => text.push_str("  \n"),
            RunContent::Break(BreakType::Page | BreakType::Column) => {
                text.push_str("\n\n---\n\n");
            },
            RunContent::Tab => text.push('\t'),
            RunContent::Drawing(drawing) => {
                markdown_drawing(drawing, text);
            },
            // See `plain_text_run`: block content needs a separator or it
            // fuses with the run that follows it.
            RunContent::TextBox(blocks) => {
                let mut inner = String::new();
                markdown_blocks(blocks, ctx, &mut inner, 0);
                let inner = inner.trim();
                if !inner.is_empty() {
                    if !text.is_empty() && !text.ends_with(['\n', ' ', '\t']) {
                        text.push('\n');
                    }
                    text.push_str(inner);
                    text.push('\n');
                }
            },
        }
    }
}

fn markdown_drawing(drawing: &DrawingInfo, out: &mut String) {
    out.push_str("![");
    if let Some(ref desc) = drawing.description {
        out.push_str(desc);
    }
    out.push_str("](");
    out.push_str(&drawing.relationship_id);
    out.push(')');
}

fn markdown_table(table: &Table, _ctx: &MarkdownCtx, out: &mut String) {
    if table.rows.is_empty() {
        return;
    }

    // Collect all cell texts
    let mut row_texts: Vec<Vec<String>> = Vec::new();
    let mut max_cols = 0usize;

    for row in &table.rows {
        let mut cells: Vec<String> = Vec::new();
        for cell in &row.cells {
            let mut cell_text = String::new();
            plain_text_blocks(&cell.content, &mut cell_text);
            let cell_text = cell_text.trim().replace('\n', " ");
            cells.push(cell_text);
        }
        max_cols = max_cols.max(cells.len());
        row_texts.push(cells);
    }

    // Pad rows to max_cols
    for row in &mut row_texts {
        while row.len() < max_cols {
            row.push(String::new());
        }
    }

    // Output header row
    if let Some(first) = row_texts.first() {
        out.push('|');
        for cell in first {
            out.push(' ');
            out.push_str(cell);
            out.push_str(" |");
        }
        out.push('\n');

        // Separator row
        out.push('|');
        for _ in 0..max_cols {
            out.push_str(" --- |");
        }
        out.push('\n');

        // Data rows
        for row in row_texts.iter().skip(1) {
            out.push('|');
            for cell in row {
                out.push(' ');
                out.push_str(cell);
                out.push_str(" |");
            }
            out.push('\n');
        }
    }
    out.push('\n');
}

fn runs_to_plain_text(runs: &[Run]) -> String {
    let mut text = String::new();
    for run in runs {
        plain_text_run(run, &mut text);
    }
    text
}
