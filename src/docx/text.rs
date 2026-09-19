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
        // Footnote/endnote/comment bodies are real document content that
        // to_ir() already carries (as Element::Footnote/Endnote) — without
        // walking them here too, a footnote-only document returned "" even
        // though it visibly has text, and a document's word count changed
        // depending on which of plain_text()/to_markdown()/to_ir() a caller
        // used.
        for n in self.footnotes.iter().chain(self.endnotes.iter()).chain(self.comments.iter()) {
            plain_text_blocks(&n.content, &mut out);
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

        // See the identical note on plain_text(): footnote,
        // endnote and comment bodies are real content that to_ir() already
        // carries, and dropping them here made this renderer disagree with
        // that one.
        for n in self.footnotes.iter().chain(self.endnotes.iter()).chain(self.comments.iter()) {
            let mut note_buf = String::new();
            markdown_blocks(&n.content, &ctx, &mut note_buf, 0);
            let note_text = note_buf.trim();
            if !note_text.is_empty() {
                if !out.ends_with("\n\n") && !out.is_empty() {
                    out.push_str("\n\n");
                }
                out.push_str(note_text);
                out.push('\n');
            }
        }

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
    // `<w:vanish/>` — Word never renders this run at all.
    if run.properties.as_ref().and_then(|rp| rp.hidden).unwrap_or(false) {
        return;
    }
    for content in &run.content {
        match content {
            RunContent::Text(text) => out.push_str(text),
            RunContent::Break(BreakType::Line) => out.push('\n'),
            RunContent::Break(BreakType::Page | BreakType::Column) => out.push('\n'),
            RunContent::Tab => out.push('\t'),
            // A drawing has no text of its own — except a native chart or
            // SmartArt diagram, whose text was resolved from the separate
            // referenced part at open time.
            RunContent::Drawing(d) => {
                let lines: Vec<&str> = d
                    .chart_text
                    .iter()
                    .chain(d.dgm_text.iter())
                    .map(String::as_str)
                    .collect();
                if !lines.is_empty() {
                    if !out.is_empty() && !out.ends_with(['\n', ' ', '\t']) {
                        out.push('\n');
                    }
                    out.push_str(&lines.join("\n"));
                    out.push('\n');
                }
            },
            // Text-box prose is document content — in some real files it is
            // most of the document. It is *block* content, so
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
            // The reference mark itself carries no text of its own — the
            // note *body* is walked separately; this is only
            // the citation point, nothing to render here.
            RunContent::FootnoteRef(..) | RunContent::EndnoteRef(..) | RunContent::CommentRef(_) => {},
            RunContent::FormField(ff) => {
                if let Some(text) = &ff.display_text {
                    out.push_str(text);
                }
            },
            // Resolved into a TextBox during from_opc when the reference
            // could be followed; an unresolvable one is dropped (matches
            // the type's own documented intent).
            RunContent::DeferredPart(_) => {},
        }
    }
}

fn plain_text_table(table: &Table, out: &mut String) {
    // A table cell can hold another table, so this recurses (via
    // `plain_text_blocks` -> `plain_text_table` -> `plain_text_blocks` ...)
    // on whatever it's given — including a tree the XML parser already
    // bounded to `MAX_NESTING_DEPTH`, which on the small default thread
    // stack `plain_text()`/`to_markdown()` run on (unlike parsing itself,
    // which gets its own larger stack) is still deep enough to overflow.
    // Past the cap, stop descending rather than crash the whole process —
    // the same defect class as an unguarded XML parse, just one layer
    // downstream of it.
    let Some(_depth) = crate::core::xml::DepthGuard::enter() else {
        out.push_str(&format!(
            "[nested table deeper than {} levels not shown — document truncated]\n",
            crate::core::xml::MAX_NESTING_DEPTH
        ));
        return;
    };
    for row in &table.rows {
        // A cell deleted via tracked changes is excluded from the
        // accepted view, same policy already applied to run-level
        // `w:del`.
        let cells: Vec<_> = row
            .cells
            .iter()
            .filter(|c| !c.properties.as_ref().is_some_and(|p| p.deleted))
            .collect();
        for (i, cell) in cells.iter().enumerate() {
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
    // Keyed by (num_id, ilvl): the next number to print for that level.
    // This renderer processes one paragraph at a time with no notion of
    // "list group" the way convert_docx.rs's IR path has, so each level's
    // count is simply "one more than last time this exact level was
    // seen" — matching the per-numId continuation semantics of the reader.
    // Without this every item printed the abstract level's bare
    // `<w:start>` value forever (the 4th instance of this
    // crate's "two renderers disagree" flaw).
    let mut numbering_counts: std::collections::HashMap<(u32, u8), u32> =
        std::collections::HashMap::new();
    markdown_blocks_inner(elements, ctx, out, _depth, &mut numbering_counts);
}

fn markdown_blocks_inner(
    elements: &[BlockElement],
    ctx: &MarkdownCtx,
    out: &mut String,
    _depth: usize,
    numbering_counts: &mut std::collections::HashMap<(u32, u8), u32>,
) {
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
                    // The printed number for an ordered format: one more
                    // than the last time this exact (num_id, ilvl) was
                    // seen, or the effective start (honoring
                    // startOverride) on first encounter — not the bare
                    // abstract-level start reprinted forever.
                    let mut next_ordinal = || {
                        let key = (nr.num_id, nr.ilvl);
                        let next = match numbering_counts.get(&key) {
                            Some(&prev) => prev + 1,
                            None => numbering
                                .resolve_start(nr.num_id, nr.ilvl)
                                .unwrap_or(level.start),
                        };
                        numbering_counts.insert(key, next);
                        next
                    };
                    let marker = match &level.format {
                        NumberFormat::Bullet => "- ".to_string(),
                        NumberFormat::Decimal => format!("{}. ", next_ordinal()),
                        NumberFormat::LowerLetter => format!("{}. ", next_ordinal()),
                        NumberFormat::UpperLetter => format!("{}. ", next_ordinal()),
                        NumberFormat::LowerRoman => format!("{}. ", next_ordinal()),
                        NumberFormat::UpperRoman => format!("{}. ", next_ordinal()),
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
    // `<w:vanish/>` — Word never renders this run at all.
    if run.properties.as_ref().and_then(|rp| rp.hidden).unwrap_or(false) {
        return;
    }
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
            RunContent::FootnoteRef(..) | RunContent::EndnoteRef(..) | RunContent::CommentRef(_) => {},
            RunContent::FormField(ff) => {
                if let Some(t) = &ff.display_text {
                    text.push_str(t);
                }
            },
            RunContent::DeferredPart(_) => {},
        }
    }
}

fn markdown_drawing(drawing: &DrawingInfo, out: &mut String) {
    // A chart or diagram is not an image: render the text we recovered
    // from the referenced part instead of an empty image link with no
    // target.
    let lines: Vec<&str> = drawing
        .chart_text
        .iter()
        .chain(drawing.dgm_text.iter())
        .map(String::as_str)
        .collect();
    if !lines.is_empty() {
        if !out.is_empty() && !out.ends_with(['\n', ' ', '\t']) {
            out.push('\n');
        }
        out.push_str(&lines.join("  \n"));
        out.push('\n');
        return;
    }
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
        // Same tracked-changes policy as plain_text_table.
        for cell in row
            .cells
            .iter()
            .filter(|c| !c.properties.as_ref().is_some_and(|p| p.deleted))
        {
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
