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
    pub fn plain_text(&self) -> String {
        let mut out = String::new();
        plain_text_blocks(&self.body.elements, &mut out);
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

        // Headers (deduped on text content — headers may be repeated for
        // first-page / even / default variants but the text is usually the
        // same; we only want one copy in flat markdown).
        let mut seen: std::collections::HashSet<String> = std::collections::HashSet::new();
        for hf in &self.headers_footers {
            if !matches!(
                hf.hf_type,
                super::HeaderFooterType::Default
                    | super::HeaderFooterType::First
                    | super::HeaderFooterType::Even
            ) {
                continue;
            }
            let mut buf = String::new();
            markdown_blocks(&hf.content, &ctx, &mut buf, 0);
            let trimmed = buf.trim();
            // Skip empty headers/footers and duplicates.
            if trimmed.is_empty() || !seen.insert(trimmed.to_string()) {
                continue;
            }
            // We don't currently know which side (header vs footer) this
            // came from at this layer — `HeaderFooter` carries only the
            // type modifier (default/first/even). The body sits between
            // the headers and footers we emit, so we put all headers
            // before and all footers after the body.
        }

        // Decide header/footer split using each section's references.
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

        // Trim trailing newlines
        while out.ends_with('\n') {
            out.pop();
        }
        let _ = seen; // silence
        out
    }
}

/// Split parsed `HeaderFooter` entries into headers vs footers using the
/// section reference lists. Returns (headers, footers) as deduplicated
/// markdown-string vectors. We don't currently retain the relationship
/// IDs that map a section ref to a specific parsed `HeaderFooter`, so we
/// approximate: header_refs.len() entries from the front go to headers,
/// the rest go to footers. Correct for the common case (single section
/// with one of each); on multi-variant documents some misclassification
/// is possible but text is still preserved (just maybe in the wrong slot).
fn split_headers_footers(doc: &DocxDocument, ctx: &MarkdownCtx) -> (Vec<String>, Vec<String>) {
    let mut headers: Vec<String> = Vec::new();
    let mut footers: Vec<String> = Vec::new();
    let mut header_seen: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut footer_seen: std::collections::HashSet<String> = std::collections::HashSet::new();

    let n_header_refs: usize = doc.sections.iter().map(|s| s.header_refs.len()).sum();
    for (idx, hf) in doc.headers_footers.iter().enumerate() {
        let mut buf = String::new();
        markdown_blocks(&hf.content, ctx, &mut buf, 0);
        let t = buf.trim().to_string();
        if t.is_empty() {
            continue;
        }
        if idx < n_header_refs {
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
                    let hashes = "#".repeat(level.min(9));
                    out.push_str(&hashes);
                    out.push(' ');
                } else if let Some(ref prefix) = list_prefix {
                    out.push_str(prefix);
                }

                // Render paragraph content with inline formatting
                for content in &p.content {
                    match content {
                        ParagraphContent::Run(run) => markdown_run(run, out),
                        ParagraphContent::Hyperlink(hl) => {
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

fn markdown_run(run: &Run, out: &mut String) {
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

    // Collect text content
    let mut text = String::new();
    for content in &run.content {
        match content {
            RunContent::Text(t) => text.push_str(t),
            RunContent::Break(BreakType::Line) => text.push_str("  \n"),
            RunContent::Break(BreakType::Page | BreakType::Column) => {
                text.push_str("\n\n---\n\n");
            },
            RunContent::Tab => text.push('\t'),
            RunContent::Drawing(drawing) => {
                markdown_drawing(drawing, &mut text);
            },
        }
    }

    if text.is_empty() {
        return;
    }

    // Apply inline formatting wrappers
    if strike {
        out.push_str("~~");
    }
    if bold && italic {
        out.push_str("***");
    } else if bold {
        out.push_str("**");
    } else if italic {
        out.push('*');
    }

    out.push_str(&text);

    if bold && italic {
        out.push_str("***");
    } else if bold {
        out.push_str("**");
    } else if italic {
        out.push('*');
    }
    if strike {
        out.push_str("~~");
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
