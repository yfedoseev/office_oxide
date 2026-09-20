use crate::ir::*;

/// How `to_markdown_with` / `to_html_with` should represent embedded images.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ImageEmbed {
    /// Render the image's description, or nothing when it has none.
    /// This is what plain `to_markdown` and `to_html` do.
    #[default]
    None,
    /// Emit the image bytes inline at the image's position in the flow:
    /// `[image-base64:<data>]` in markdown, `<img src="data:…;base64,…">`
    /// in HTML.
    ///
    /// Keeps both the position and the content in one self-contained
    /// string, which is what a vision-capable model consuming the markdown
    /// — or a browser opening the HTML with no sidecar files — needs;
    /// images were otherwise dropped entirely.
    Base64,
}

/// Options for [`DocumentIR::to_markdown_with`].
#[derive(Debug, Clone, Copy, Default)]
pub struct MarkdownOptions {
    /// How to represent embedded images.
    pub image_embed: ImageEmbed,
}

/// Options for [`DocumentIR::to_html_with`].
#[derive(Debug, Clone, Copy, Default)]
pub struct HtmlOptions {
    /// How to represent embedded images.
    pub image_embed: ImageEmbed,
}

thread_local! {
    /// Rendering options for the current `to_markdown_with` call.
    ///
    /// The renderer is a tree of free functions taking only the node; a
    /// thread-local avoids threading an options parameter through every one
    /// of them purely to reach the single `Element::Image` arm.
    static MARKDOWN_OPTIONS: std::cell::Cell<MarkdownOptions> =
        const { std::cell::Cell::new(MarkdownOptions {
            image_embed: ImageEmbed::None,
        }) };

    /// Rendering options for the current `to_html_with` call; same
    /// reasoning as `MARKDOWN_OPTIONS`.
    static HTML_OPTIONS: std::cell::Cell<HtmlOptions> =
        const { std::cell::Cell::new(HtmlOptions {
            image_embed: ImageEmbed::None,
        }) };
}

/// The media type to put in an image's `data:` URI.
///
/// `Image::format` is authoritative when the converter set it; otherwise the
/// bytes are sniffed, because a data URI with the wrong (or a generic) type
/// does not render. `None` means "do not emit a data URI for this image" —
/// the caller falls back to describing it.
fn image_mime(img: &Image) -> Option<&'static str> {
    if let Some(ref fmt) = img.format {
        return Some(fmt.content_type());
    }
    let data = img.data.as_deref()?;
    Some(match data {
        [0x89, b'P', b'N', b'G', ..] => "image/png",
        [0xFF, 0xD8, 0xFF, ..] => "image/jpeg",
        [b'G', b'I', b'F', b'8', ..] => "image/gif",
        [b'B', b'M', ..] => "image/bmp",
        [0x49, 0x49, 0x2A, 0x00, ..] | [0x4D, 0x4D, 0x00, 0x2A, ..] => "image/tiff",
        _ => return None,
    })
}

/// Plain-text marker for a page, slide or thematic boundary.
///
/// A form feed is the conventional plain-text page separator (and what
/// this crate's own `.ppt` handling looks for). The previous marker was
/// markdown's `---`, which reached plain-text consumers as literal text
/// that appears nowhere in the source document.
pub const PLAIN_BREAK: &str = "\u{000C}";

mod block_default {
    //! Default flow-rendering for [`Element`] variants that don't
    //! carry a meaningful inline / paragraph / heading shape.
    //!
    //! Each `default_*` function is **exhaustive** over `Element`:
    //! the compiler forces a decision when a new variant is added
    //! ("is this variant invisible in flow output, or do specific
    //! renderers need to handle it?"). Renderers in the parent
    //! module keep arms only for variants where their output
    //! differs from these defaults; everything else falls through
    //! to the matching `default_*` here via `other => default_X(other)`.
    use super::*;
    use std::fmt::Write;

    /// Plain-text default. Most invisible variants → `""`;
    /// `ThematicBreak` → a form feed (`U+000C`), the conventional
    /// plain-text page/rule separator — `---` is markdown syntax and has
    /// no business in the plain-text renderer, where it arrives at the
    /// consumer as literal text that is not in the document. Container
    /// elements recursively render their children.
    pub fn default_plain(element: &Element) -> String {
        match element {
            Element::ThematicBreak => PLAIN_BREAK.to_string(),
            Element::TextBox(tb) => tb
                .content
                .iter()
                .map(super::render_element_plain)
                .collect::<Vec<_>>()
                .join("\n\n"),
            Element::Footnote(n) | Element::Endnote(n) => n
                .content
                .iter()
                .map(super::render_element_plain)
                .collect::<Vec<_>>()
                .join("\n\n"),
            // A page or column break is a real boundary in the source, so
            // it gets the same form-feed marker a thematic break does.
            Element::PageBreak | Element::ColumnBreak => PLAIN_BREAK.to_string(),
            // Invisible in flow: shapes are positioned, not flow content;
            // an unannotated image shows nothing in plain text.
            Element::Shape(_) | Element::Image(_) => String::new(),
            // The variants below have rich flow output and shouldn't
            // hit this default — `render_element_plain` handles them.
            // Reaching here means a renderer forgot a real arm; we
            // emit empty rather than panic so the document still
            // renders, but the explicit arms below let the compiler
            // catch added variants.
            Element::Heading(_)
            | Element::Paragraph(_)
            | Element::Table(_)
            | Element::List(_)
            | Element::CodeBlock(_) => String::new(),
        }
    }

    /// Markdown default. Same as plain except images get an alt-text
    /// `![alt]()` form.
    pub fn default_markdown(element: &Element) -> String {
        match element {
            Element::ThematicBreak => "---".to_string(),
            Element::TextBox(tb) => tb
                .content
                .iter()
                .map(super::render_element_markdown)
                .collect::<Vec<_>>()
                .join("\n\n"),
            Element::Footnote(n) | Element::Endnote(n) => n
                .content
                .iter()
                .map(super::render_element_markdown)
                .collect::<Vec<_>>()
                .join("\n\n"),
            Element::PageBreak | Element::ColumnBreak | Element::Shape(_) => String::new(),
            Element::Image(img) => {
                // With `ImageEmbed::Base64` the bytes go inline at the
                // image's position in the flow.
                if super::MARKDOWN_OPTIONS.with(|o| o.get().image_embed) == ImageEmbed::Base64 {
                    if let Some(ref data) = img.data {
                        return format!("[image-base64:{}]", crate::core::base64::encode(data));
                    }
                }
                // An `![alt]()` with an empty target renders as a broken
                // image. With no addressable source, emit the description
                // as ordinary italic text, and nothing when there is none.
                match img.alt_text.as_deref() {
                    Some(alt) if !alt.is_empty() => format!("*{}*", escape_markdown(alt)),
                    _ => String::new(),
                }
            },
            Element::Heading(_)
            | Element::Paragraph(_)
            | Element::Table(_)
            | Element::List(_)
            | Element::CodeBlock(_) => String::new(),
        }
    }

    /// HTML default. `ThematicBreak` → `<hr />`; images render an
    /// empty `<img alt="…"/>`; everything else mirrors `default_plain`
    /// behaviour with HTML escaping.
    pub fn default_html(element: &Element) -> String {
        match element {
            Element::ThematicBreak => "<hr />".to_string(),
            Element::TextBox(tb) => super::render_elements_html(&tb.content).join("\n"),
            Element::Footnote(n) | Element::Endnote(n) => {
                super::render_elements_html(&n.content).join("\n")
            },
            Element::PageBreak | Element::ColumnBreak | Element::Shape(_) => String::new(),
            Element::Image(img) => {
                // With `ImageEmbed::Base64` the bytes go inline as a data
                // URI, so the HTML is self-contained — `to_html` otherwise
                // has no way at all to show an image.
                if super::HTML_OPTIONS.with(|o| o.get().image_embed) == ImageEmbed::Base64 {
                    if let (Some(data), Some(mime)) = (img.data.as_ref(), super::image_mime(img)) {
                        let alt = img.alt_text.as_deref().unwrap_or("");
                        let mut out = String::new();
                        let _ = write!(
                            out,
                            "<img src=\"data:{mime};base64,{}\" alt=\"{}\" />",
                            crate::core::base64::encode(data),
                            super::escape_html(alt)
                        );
                        return out;
                    }
                }
                // `src` is required on `<img>`; an element without one is
                // invalid HTML. With no addressable source in the IR,
                // describe the image with its alt text instead.
                match img.alt_text.as_deref() {
                    Some(alt) if !alt.is_empty() => {
                        let mut out = String::new();
                        let _ = write!(
                            out,
                            "<figure><figcaption>{}</figcaption></figure>",
                            super::escape_html(alt)
                        );
                        out
                    },
                    _ => String::new(),
                }
            },
            Element::Heading(_)
            | Element::Paragraph(_)
            | Element::Table(_)
            | Element::List(_)
            | Element::CodeBlock(_) => String::new(),
        }
    }
}

impl DocumentIR {
    /// Render the IR as plain text.
    pub fn plain_text(&self) -> String {
        let section_texts: Vec<String> = self
            .sections
            .iter()
            .map(render_section_plain)
            .filter(|s| !s.is_empty())
            .collect();
        if section_texts.len() <= 1 {
            section_texts.into_iter().next().unwrap_or_default()
        } else {
            section_texts.join(&format!("\n\n{PLAIN_BREAK}\n\n"))
        }
    }

    /// Render the IR as an HTML fragment (no `<html>`/`<body>` wrapper).
    pub fn to_html(&self) -> String {
        self.to_html_with(HtmlOptions::default())
    }

    /// Render the IR as an HTML fragment with explicit options.
    ///
    /// With [`ImageEmbed::Base64`] each image whose bytes the IR carries is
    /// emitted as an `<img src="data:…;base64,…">`, giving a genuinely
    /// self-contained preview — the mirror of `to_markdown_with`'s existing
    /// image-embedding option.
    pub fn to_html_with(&self, options: HtmlOptions) -> String {
        HTML_OPTIONS.with(|o| o.set(options));
        let section_texts: Vec<String> = self
            .sections
            .iter()
            .map(render_section_html)
            .filter(|s| !s.is_empty())
            .collect();
        HTML_OPTIONS.with(|o| o.set(HtmlOptions::default()));
        section_texts.join("\n<hr />\n")
    }

    /// Render the IR as markdown.
    pub fn to_markdown(&self) -> String {
        self.to_markdown_with(MarkdownOptions::default())
    }

    /// Render the IR as markdown with explicit options.
    pub fn to_markdown_with(&self, options: MarkdownOptions) -> String {
        MARKDOWN_OPTIONS.with(|o| o.set(options));
        let section_texts: Vec<String> = self
            .sections
            .iter()
            .map(render_section_markdown)
            .filter(|s| !s.is_empty())
            .collect();
        MARKDOWN_OPTIONS.with(|o| o.set(MarkdownOptions::default()));
        section_texts.join("\n\n---\n\n")
    }
}

// ---------------------------------------------------------------------------
// Plain text rendering
// ---------------------------------------------------------------------------

/// The header/footer parts of a section, in the order they should be
/// rendered around the body: headers first, footers last.
///
/// All three renderers use this so they agree on what "the text of this
/// document" means. Previously only the DOCX markdown path emitted
/// headers and footers, so a consumer's word count changed depending on
/// which method they called.
/// Whether `Section::title` merely repeats the section's own first
/// heading.
///
/// The DOCX and PPTX converters set `Section.title` from the text of the
/// first `Element::Heading` and leave that heading in `elements`. Rendering
/// both printed every section's opening heading twice — once as a
/// synthesised `## {title}` at a fixed level, then again at its real level.
/// The title still exists for consumers that want a section label; it just
/// must not be rendered as body content when it is a copy.
fn section_title_is_redundant(section: &Section) -> bool {
    let Some(title) = section.title.as_deref().filter(|t| !t.is_empty()) else {
        return false;
    };
    match section.elements.first() {
        Some(Element::Heading(h)) => render_inline_plain(&h.content).trim() == title.trim(),
        _ => false,
    }
}

fn section_headers(section: &Section) -> impl Iterator<Item = &HeaderFooter> {
    [
        section.first_page_header.as_ref(),
        section.header.as_ref(),
        section.even_page_header.as_ref(),
    ]
    .into_iter()
    .flatten()
}

fn section_footers(section: &Section) -> impl Iterator<Item = &HeaderFooter> {
    [
        section.first_page_footer.as_ref(),
        section.footer.as_ref(),
        section.even_page_footer.as_ref(),
    ]
    .into_iter()
    .flatten()
}

fn render_section_plain(section: &Section) -> String {
    let mut parts = Vec::new();
    for hf in section_headers(section) {
        for elem in &hf.content {
            let text = render_element_plain(elem);
            if !text.is_empty() {
                parts.push(text);
            }
        }
    }
    if section_title_is_redundant(section) {
        // The title was lifted out of the section's own first heading; the
        // heading is still in `elements`, so emitting both prints it twice.
    } else if let Some(ref title) = section.title {
        if !title.is_empty() {
            parts.push(title.clone());
        }
    }
    for elem in &section.elements {
        let text = render_element_plain(elem);
        if !text.is_empty() {
            parts.push(text);
        }
    }
    // Speaker notes are not part of the visible surface; label them so a
    // consumer can tell them apart from slide body text.
    if let Some(ref notes) = section.speaker_notes {
        let text = notes
            .iter()
            .map(render_element_plain)
            .filter(|t| !t.is_empty())
            .collect::<Vec<_>>()
            .join("\n");
        if !text.is_empty() {
            parts.push(format!("[Notes]\n{text}"));
        }
    }
    for hf in section_footers(section) {
        for elem in &hf.content {
            let text = render_element_plain(elem);
            if !text.is_empty() {
                parts.push(text);
            }
        }
    }
    parts.join("\n\n")
}

fn render_element_plain(element: &Element) -> String {
    match element {
        Element::Heading(h) => render_inline_plain(&h.content),
        Element::Paragraph(p) => render_inline_plain(&p.content),
        Element::Table(t) => render_table_plain(t),
        Element::List(l) => render_list_plain(l, 0),
        Element::Image(img) => match &img.alt_text {
            Some(alt) => format!("[{alt}]"),
            None => String::new(),
        },
        Element::CodeBlock(cb) => cb.content.clone(),
        // Invisible-in-flow / container variants delegated to the
        // shared default. Adding a new `Element` variant forces a
        // compile error in `block_default::default_plain`, not here.
        other => block_default::default_plain(other),
    }
}

fn render_inline_plain(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => out.push_str(&span.text),
            InlineContent::LineBreak => out.push('\n'),
            InlineContent::FootnoteRef(_) | InlineContent::EndnoteRef(_) => {},
        }
    }
    out
}

fn render_table_plain(table: &Table) -> String {
    // Tab-separated output is column-aligned, so a spanned cell must leave
    // the positions it covers empty rather than shifting its neighbours.
    let mut rows = Vec::new();
    for row in table_grid(table) {
        let cells: Vec<String> = row
            .iter()
            .map(|slot| match slot {
                Some(cell) => cell
                    .content
                    .iter()
                    .map(render_element_plain)
                    .collect::<Vec<_>>()
                    .join(" "),
                None => String::new(),
            })
            .collect();
        rows.push(cells.join("\t"));
    }
    rows.join("\n")
}

fn render_list_plain(list: &List, indent: usize) -> String {
    let prefix_str = " ".repeat(indent * 2);
    let mut lines = Vec::new();
    for item in &list.items {
        let text = item
            .content
            .iter()
            .map(render_element_plain)
            .collect::<Vec<_>>()
            .join(" ");
        lines.push(format!("{prefix_str}- {text}"));
        if let Some(ref nested) = item.nested {
            lines.push(render_list_plain(nested, indent + 1));
        }
    }
    lines.join("\n")
}

// ---------------------------------------------------------------------------
// Markdown rendering
// ---------------------------------------------------------------------------

fn render_section_markdown(section: &Section) -> String {
    let mut parts = Vec::new();
    for hf in section_headers(section) {
        for elem in &hf.content {
            let text = render_element_markdown(elem);
            if !text.is_empty() {
                parts.push(text);
            }
        }
    }
    if section_title_is_redundant(section) {
        // See `section_title_is_redundant`.
    } else if let Some(ref title) = section.title {
        if !title.is_empty() {
            parts.push(format!("## {title}"));
        }
    }
    for elem in &section.elements {
        let text = render_element_markdown(elem);
        if !text.is_empty() {
            parts.push(text);
        }
    }
    if let Some(ref notes) = section.speaker_notes {
        let body = notes
            .iter()
            .map(render_element_markdown)
            .filter(|t| !t.is_empty())
            .collect::<Vec<_>>()
            .join("\n\n");
        if !body.is_empty() {
            // Every line must start with `>` to stay inside the
            // blockquote — a bare newline (e.g. between list items)
            // would otherwise end it partway through the notes.
            let quoted = body
                .lines()
                .map(|l| format!("> {l}"))
                .collect::<Vec<_>>()
                .join("\n");
            parts.push(format!("> **Notes:**\n{quoted}"));
        }
    }
    for hf in section_footers(section) {
        for elem in &hf.content {
            let text = render_element_markdown(elem);
            if !text.is_empty() {
                parts.push(text);
            }
        }
    }
    parts.join("\n\n")
}

fn render_element_markdown(element: &Element) -> String {
    match element {
        Element::Heading(h) => {
            let hashes = "#".repeat(h.clamped_level() as usize);
            let text = render_inline_markdown(&h.content);
            format!("{hashes} {text}")
        },
        Element::Paragraph(p) => render_inline_markdown(&p.content),
        Element::Table(t) => render_table_markdown(t),
        Element::List(l) => render_list_markdown(l, 0),
        Element::CodeBlock(cb) => {
            let lang = cb.language.as_deref().unwrap_or("");
            format!("```{lang}\n{}\n```", cb.content)
        },
        // Invisible-in-flow / container / image variants delegated
        // to the shared default — see `block_default::default_markdown`.
        other => block_default::default_markdown(other),
    }
}

/// The formatting that decides which markdown delimiters wrap a span.
///
/// Word splits a single visually-bold phrase into several runs constantly —
/// a spell-check boundary, a language attribute or a revision id is enough
/// — so adjacent runs must be merged before delimiters are emitted.
/// Wrapping each run separately produced `**BOLD_A****BOLD_B**`, and
/// CommonMark reads that `****` as four literal asterisks rather than as
/// the end of one emphasis span and the start of another.
#[derive(PartialEq)]
struct MarkdownStyle {
    bold: bool,
    italic: bool,
    strikethrough: bool,
    vertical_align: Option<VerticalAlign>,
    hyperlink: Option<String>,
}

impl MarkdownStyle {
    fn of(span: &TextSpan) -> Self {
        Self {
            bold: span.bold,
            italic: span.italic,
            strikethrough: span.strikethrough,
            vertical_align: span.vertical_align.clone(),
            hyperlink: span.hyperlink.as_deref().and_then(safe_url),
        }
    }

    /// Wrap already-escaped text in this style's delimiters.
    fn wrap(&self, text: &str) -> String {
        if text.is_empty() {
            return String::new();
        }
        // Leading/trailing spaces must sit outside the delimiters: CommonMark
        // does not open emphasis on `** text**`. An all-whitespace span has
        // no core to emphasise — and computing the two spans independently
        // makes them overlap, which inverts the slice range and panics, so
        // take the trailing span from what is left after the leading one.
        let core_str = text.trim();
        if core_str.is_empty() {
            return text.to_string();
        }
        let lead_len = text.len() - text.trim_start().len();
        let lead = &text[..lead_len];
        let trail = &text[lead_len + core_str.len()..];
        let core = core_str;

        let mut out = core.to_string();
        // Super/subscript have no markdown syntax; HTML is the conventional
        // fallback and is what every markdown flavour renders.
        match self.vertical_align {
            Some(VerticalAlign::Superscript) => out = format!("<sup>{out}</sup>"),
            Some(VerticalAlign::Subscript) => out = format!("<sub>{out}</sub>"),
            _ => {},
        }
        if self.strikethrough {
            out = format!("~~{out}~~");
        }
        if self.bold && self.italic {
            out = format!("***{out}***");
        } else if self.bold {
            out = format!("**{out}**");
        } else if self.italic {
            out = format!("*{out}*");
        }
        if let Some(ref url) = self.hyperlink {
            out = format!("[{out}]({})", escape_markdown_url(url));
        }
        format!("{lead}{out}{trail}")
    }
}

fn render_inline_markdown(content: &[InlineContent]) -> String {
    let mut out = String::new();
    // Accumulate consecutive spans that share formatting, and emit the run
    // once with a single pair of delimiters.
    let mut pending: Option<(MarkdownStyle, String)> = None;

    let flush = |pending: &mut Option<(MarkdownStyle, String)>, out: &mut String| {
        if let Some((style, text)) = pending.take() {
            out.push_str(&style.wrap(&text));
        }
    };

    for item in content {
        match item {
            InlineContent::Text(span) => {
                let style = MarkdownStyle::of(span);
                let text = escape_markdown(&span.text);
                match pending.as_mut() {
                    Some((cur, buf)) if *cur == style => buf.push_str(&text),
                    _ => {
                        flush(&mut pending, &mut out);
                        pending = Some((style, text));
                    },
                }
            },
            InlineContent::LineBreak => {
                flush(&mut pending, &mut out);
                out.push_str("  \n");
            },
            InlineContent::FootnoteRef(_) | InlineContent::EndnoteRef(_) => {},
        }
    }
    flush(&mut pending, &mut out);
    out
}

/// Lay a table out on a grid, resolving `col_span` and `row_span` into the
/// positions each cell actually occupies.
///
/// Markdown has no cell-spanning syntax, so a spanned cell's text goes in
/// its top-left position and the positions it covers render empty. Indexing
/// `row.cells` positionally instead — which is what this did — shifted every
/// cell to the right of a rowspan one column left, because the covered
/// position has no cell of its own in the IR.
fn table_grid(table: &Table) -> Vec<Vec<Option<&TableCell>>> {
    // Width is the widest row measured in grid columns, not cell count.
    // Defence in depth: the converters clamp spans, but an IR built by a
    // caller (it is Deserialize) can still carry anything, and this sizes an
    // allocation. Cap against the cells actually present — a span cannot
    // legitimately describe more columns than the table has content for.
    let cell_total: usize = table.rows.iter().map(|r| r.cells.len()).sum();
    let width = table
        .rows
        .iter()
        .map(|r| {
            r.cells
                .iter()
                .map(|c| c.col_span.max(1) as usize)
                .sum::<usize>()
        })
        .max()
        .unwrap_or(0)
        .min(cell_total.saturating_mul(1_000).max(1));
    let mut grid: Vec<Vec<Option<&TableCell>>> = vec![vec![None; width]; table.rows.len()];
    // Positions already claimed by a cell spanning down from an earlier row.
    let mut covered: Vec<Vec<bool>> = vec![vec![false; width]; table.rows.len()];

    for (r, row) in table.rows.iter().enumerate() {
        let mut c = 0usize;
        for cell in &row.cells {
            while c < width && covered[r][c] {
                c += 1;
            }
            if c >= width {
                break;
            }
            grid[r][c] = Some(cell);
            let cs = cell.col_span.max(1) as usize;
            let rs = cell.row_span.max(1) as usize;
            for dr in 0..rs {
                for dc in 0..cs {
                    if r + dr < covered.len() && c + dc < width {
                        covered[r + dr][c + dc] = true;
                    }
                }
            }
            c += cs;
        }
    }
    grid
}

fn render_table_markdown(table: &Table) -> String {
    if table.rows.is_empty() {
        return String::new();
    }

    let grid = table_grid(table);
    let col_count = grid.first().map(|r| r.len()).unwrap_or(0);
    if col_count == 0 {
        return String::new();
    }

    let mut result = String::new();

    let write_row = |cells: &[Option<&TableCell>], out: &mut String| {
        out.push('|');
        for slot in cells.iter().take(col_count) {
            out.push(' ');
            out.push_str(&slot.map(render_cell_markdown).unwrap_or_default());
            out.push_str(" |");
        }
        out.push('\n');
    };

    write_row(&grid[0], &mut result);

    // Separator
    result.push('|');
    for _ in 0..col_count {
        result.push_str(" --- |");
    }
    result.push('\n');

    for row in grid.iter().skip(1) {
        write_row(row, &mut result);
    }

    // Remove trailing newline
    if result.ends_with('\n') {
        result.pop();
    }

    result
}

fn render_cell_markdown(cell: &TableCell) -> String {
    cell.content
        .iter()
        .map(|e| match e {
            Element::Paragraph(p) => render_inline_markdown(&p.content),
            other => render_element_markdown(other),
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn render_list_markdown(list: &List, indent: usize) -> String {
    let prefix_str = "  ".repeat(indent);
    // A numbered list that starts at 3 in the source must start at 3 here:
    // `start_number` was parsed and then ignored, so every ordered list
    // rendered as 1, 2, 3 regardless of what the document said.
    let start = if list.ordered {
        list.start_number.unwrap_or(1)
    } else {
        1
    };
    let mut lines = Vec::new();
    for (i, item) in list.items.iter().enumerate() {
        let text = item
            .content
            .iter()
            .map(render_element_markdown)
            .collect::<Vec<_>>()
            .join(" ");
        let marker = if list.ordered {
            format!("{}. ", start.saturating_add(u32::try_from(i).unwrap_or(u32::MAX)))
        } else {
            "- ".to_string()
        };
        lines.push(format!("{prefix_str}{marker}{text}"));
        if let Some(ref nested) = item.nested {
            lines.push(render_list_markdown(nested, indent + 1));
        }
    }
    lines.join("\n")
}

// ---------------------------------------------------------------------------
// HTML rendering
// ---------------------------------------------------------------------------

/// Reject URL schemes that execute when a rendered document is opened.
///
/// A document is untrusted input: a `javascript:` or `data:text/html`
/// hyperlink copied verbatim into generated HTML or markdown becomes an
/// XSS vector in whatever viewer displays it. Relative URLs, fragments and
/// the ordinary network schemes pass through; anything with an unknown
/// scheme is dropped so the link text still renders as plain text.
fn safe_url(url: &str) -> Option<String> {
    let trimmed = url.trim();
    if trimmed.is_empty() {
        return None;
    }
    // A scheme is everything before the first ':' when that prefix contains
    // no '/', '?' or '#'. Control characters are stripped first: browsers
    // ignore them, so `java\0script:` would otherwise slip through.
    let cleaned: String = trimmed.chars().filter(|c| !c.is_control()).collect();
    let scheme_end = cleaned
        .find(':')
        .filter(|&i| !cleaned[..i].contains(['/', '?', '#']));
    match scheme_end {
        None => Some(cleaned),
        Some(i) => {
            let scheme = cleaned[..i].to_ascii_lowercase();
            const ALLOWED: &[&str] = &[
                "http", "https", "mailto", "tel", "ftp", "ftps", "sms", "callto", "file",
            ];
            ALLOWED.contains(&scheme.as_str()).then_some(cleaned)
        },
    }
}

/// Escape the markdown metacharacters that would otherwise let document
/// text inject structure into the rendered output — a cell containing
/// `|` splitting a table row, or a literal `[x](y)` becoming a link.
fn escape_markdown(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        // `_` is deliberately absent: CommonMark does not treat intra-word
        // `_` as emphasis, and escaping it turns ordinary identifiers like
        // `HEADER_TEXT` into unreadable `HEADER\_TEXT`.
        if matches!(c, '\\' | '`' | '*' | '[' | ']' | '<' | '>' | '|' | '~') {
            out.push('\\');
        }
        out.push(c);
    }
    out
}

/// Escape the characters that would terminate a markdown link target early.
fn escape_markdown_url(s: &str) -> String {
    s.replace('(', "%28")
        .replace(')', "%29")
        .replace(' ', "%20")
}

fn escape_html(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn render_section_html(section: &Section) -> String {
    let mut parts = Vec::new();
    for hf in section_headers(section) {
        for elem in &hf.content {
            let html = render_element_html(elem);
            if !html.is_empty() {
                parts.push(format!("<header>{html}</header>"));
            }
        }
    }
    if section_title_is_redundant(section) {
        // See `section_title_is_redundant`.
    } else if let Some(ref title) = section.title {
        if !title.is_empty() {
            parts.push(format!("<h2>{}</h2>", escape_html(title)));
        }
    }
    for html in render_elements_html(&section.elements) {
        if !html.is_empty() {
            parts.push(html);
        }
    }
    // Speaker notes are not slide-surface content, but dropping them from
    // HTML loses text the plain and markdown renderers both keep.
    if let Some(ref notes) = section.speaker_notes {
        let body = render_elements_html(notes)
            .into_iter()
            .filter(|h| !h.is_empty())
            .collect::<Vec<_>>()
            .join("");
        if !body.is_empty() {
            parts.push(format!("<aside class=\"speaker-notes\">{body}</aside>"));
        }
    }
    for hf in section_footers(section) {
        for elem in &hf.content {
            let html = render_element_html(elem);
            if !html.is_empty() {
                parts.push(format!("<footer>{html}</footer>"));
            }
        }
    }
    parts.join("\n")
}

fn render_element_html(element: &Element) -> String {
    match element {
        Element::Heading(h) => {
            let level = h.clamped_level();
            let content = render_inline_html(&h.content);
            format!("<h{level}>{content}</h{level}>")
        },
        Element::Paragraph(p) => {
            let content = render_inline_html(&p.content);
            format!("<p>{content}</p>")
        },
        Element::Table(t) => render_table_html(t),
        Element::List(l) => render_list_html(l),
        Element::CodeBlock(cb) => {
            let escaped = escape_html(&cb.content);
            format!("<pre><code>{escaped}</code></pre>")
        },
        // Invisible-in-flow / container / image variants delegated
        // to the shared default — see `block_default::default_html`.
        other => block_default::default_html(other),
    }
}

/// A CSS colour literal for an IR RGB triple.
fn css_rgb(rgb: [u8; 3]) -> String {
    format!("#{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2])
}

/// A font family name reduced to what is safe inside a quoted CSS value.
///
/// A font name comes from an untrusted document, so anything that could
/// terminate the quoted value or the declaration — quotes, `;`, `(`, `)` —
/// is dropped rather than escaped. Letters (including non-ASCII, for CJK
/// family names), digits, spaces and the few punctuation marks real font
/// names use survive.
fn css_font_family(name: &str) -> Option<String> {
    let cleaned: String = name
        .chars()
        .filter(|c| c.is_alphanumeric() || matches!(c, ' ' | '-' | '_' | '.'))
        .collect();
    let cleaned = cleaned.trim();
    (!cleaned.is_empty()).then(|| format!("font-family:'{cleaned}'"))
}

/// `text-decoration-style` for the underline variants a plain `<u>` cannot
/// distinguish. `<u>` already supplies `text-decoration-line: underline`.
fn underline_decoration_style(style: &UnderlineStyle) -> Option<&'static str> {
    match style {
        UnderlineStyle::Double => Some("double"),
        UnderlineStyle::Dotted => Some("dotted"),
        UnderlineStyle::Dash | UnderlineStyle::DotDash | UnderlineStyle::DotDotDash => {
            Some("dashed")
        },
        UnderlineStyle::Wave => Some("wavy"),
        UnderlineStyle::Single
        | UnderlineStyle::Thick
        | UnderlineStyle::Words
        | UnderlineStyle::None => None,
    }
}

/// The CSS declarations for the span formatting that has no dedicated HTML
/// element: colour, highlight, font and the caps variants.
///
/// Returns `None` when the span carries none of them, so an unformatted
/// span still renders as bare text with no wrapper.
fn span_style_css(span: &TextSpan) -> Option<String> {
    let mut decls: Vec<String> = Vec::new();
    if let Some(rgb) = span.color {
        decls.push(format!("color:{}", css_rgb(rgb)));
    }
    if let Some(rgb) = span.highlight {
        decls.push(format!("background-color:{}", css_rgb(rgb)));
    }
    if let Some(ref name) = span.font_name {
        if let Some(decl) = css_font_family(name) {
            decls.push(decl);
        }
    }
    if let Some(half_pt) = span.font_size_half_pt {
        if half_pt % 2 == 0 {
            decls.push(format!("font-size:{}pt", half_pt / 2));
        } else {
            decls.push(format!("font-size:{}.5pt", half_pt / 2));
        }
    }
    // `all_caps` wins over `small_caps` when a document sets both, which is
    // what Word renders.
    if span.all_caps {
        decls.push("text-transform:uppercase".to_string());
    } else if span.small_caps {
        decls.push("font-variant:small-caps".to_string());
    }
    (!decls.is_empty()).then(|| decls.join(";"))
}

fn render_inline_html(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => {
                let mut text = escape_html(&span.text);

                // Super/subscript sit innermost so the raised text still
                // picks up the emphasis and colour wrapped around it.
                match span.vertical_align {
                    Some(VerticalAlign::Superscript) => text = format!("<sup>{text}</sup>"),
                    Some(VerticalAlign::Subscript) => text = format!("<sub>{text}</sub>"),
                    Some(VerticalAlign::Baseline) | None => {},
                }
                if span.bold {
                    text = format!("<strong>{text}</strong>");
                }
                if span.italic {
                    text = format!("<em>{text}</em>");
                }
                if span.strikethrough {
                    text = format!("<del>{text}</del>");
                }
                // `UnderlineStyle::None` is an explicit "not underlined" in
                // the source, so it must not produce a `<u>`.
                if let Some(ref u) = span.underline {
                    if *u != UnderlineStyle::None {
                        text = match underline_decoration_style(u) {
                            Some(kind) => {
                                format!("<u style=\"text-decoration-style:{kind}\">{text}</u>")
                            },
                            None => format!("<u>{text}</u>"),
                        };
                    }
                }
                if let Some(css) = span_style_css(span) {
                    text = format!("<span style=\"{}\">{text}</span>", escape_html(&css));
                }
                if let Some(url) = span.hyperlink.as_deref().and_then(safe_url) {
                    text = format!("<a href=\"{}\">{text}</a>", escape_html(&url));
                }

                out.push_str(&text);
            },
            InlineContent::LineBreak => out.push_str("<br />"),
            InlineContent::FootnoteRef(_) | InlineContent::EndnoteRef(_) => {},
        }
    }
    out
}

fn render_table_html(table: &Table) -> String {
    let mut html = String::from("<table>\n");

    for row in &table.rows {
        html.push_str("<tr>");
        let tag = if row.is_header { "th" } else { "td" };
        for cell in &row.cells {
            let mut attrs = String::new();
            if cell.col_span > 1 {
                attrs.push_str(&format!(" colspan=\"{}\"", cell.col_span));
            }
            if cell.row_span > 1 {
                attrs.push_str(&format!(" rowspan=\"{}\"", cell.row_span));
            }
            let content = render_elements_html(&cell.content).join("");
            html.push_str(&format!("<{tag}{attrs}>{content}</{tag}>"));
        }
        html.push_str("</tr>\n");
    }

    html.push_str("</table>");
    html
}

fn render_list_html(list: &List) -> String {
    render_list_group_html(&[list])
}

/// Render one or more `List`s as a single `<ul>`/`<ol>` block.
///
/// Real DOCX generators hand visually-continuous bullets a fresh `w:numId`
/// per paragraph, which the converter faithfully turns into one
/// `Element::List` per fragment. Emitting a separate list block for each
/// produced a run of one-item `<ul>`s — extra margins in a browser and
/// "list, 1 item" announced repeatedly by a screen reader — where markdown's
/// line-based output incidentally showed one continuous list. Adjacent
/// fragments that agree on shape are re-joined here; see
/// [`merge_adjacent_lists`] for what counts as adjacent.
fn render_list_group_html(lists: &[&List]) -> String {
    let Some(first) = lists.first() else {
        return String::new();
    };
    let tag = if first.ordered { "ol" } else { "ul" };
    // `start` only exists on `<ol>`; a browser ignores it on `<ul>`. Omitted
    // for 1, which is the attribute's own default, so ordinary lists keep a
    // bare `<ol>`.
    let start_attr = match first.start_number {
        Some(n) if first.ordered && n != 1 => format!(" start=\"{n}\""),
        _ => String::new(),
    };
    let mut html = format!("<{tag}{start_attr}>\n");
    for list in lists {
        for item in &list.items {
            let content = render_elements_html(&item.content).join("");
            html.push_str(&format!("<li>{content}"));
            if let Some(ref nested) = item.nested {
                html.push('\n');
                html.push_str(&render_list_html(nested));
            }
            html.push_str("</li>\n");
        }
    }
    html.push_str(&format!("</{tag}>"));
    html
}

/// Whether `next` is a continuation of `prev` rather than a new list.
///
/// Conservative on purpose: the two must agree on ordered-ness, marker
/// style and nesting level, and an ordered list that carries its own
/// explicit `start_number` is a deliberate restart and stays separate.
fn lists_are_continuous(prev: &List, next: &List) -> bool {
    prev.ordered == next.ordered
        && prev.style == next.style
        && prev.level == next.level
        && !(next.ordered && next.start_number.is_some())
}

/// Group a block-element slice into runs, coalescing adjacent continuous
/// `Element::List` siblings so each run renders as one list block.
///
/// Borrows throughout — nothing is cloned, so this costs nothing on the
/// large documents where element counts matter.
fn merge_adjacent_lists(elements: &[Element]) -> Vec<Vec<&Element>> {
    let mut groups: Vec<Vec<&Element>> = Vec::with_capacity(elements.len());
    for element in elements {
        let continues = match (element, groups.last().and_then(|g| g.last())) {
            (Element::List(next), Some(Element::List(prev))) => lists_are_continuous(prev, next),
            _ => false,
        };
        if continues {
            // `continues` is only true when a last group exists.
            if let Some(group) = groups.last_mut() {
                group.push(element);
            }
        } else {
            groups.push(vec![element]);
        }
    }
    groups
}

/// Render a block-element slice to one HTML string per emitted block,
/// with adjacent continuous lists merged into single list blocks.
///
/// Every place that walks a `Vec<Element>` for HTML goes through this so
/// the merge applies uniformly — section bodies, table cells, text boxes
/// and note bodies alike.
fn render_elements_html(elements: &[Element]) -> Vec<String> {
    merge_adjacent_lists(elements)
        .into_iter()
        .map(|group| match group.as_slice() {
            [single] => render_element_html(single),
            many => {
                let lists: Vec<&List> = many
                    .iter()
                    .filter_map(|e| match e {
                        Element::List(l) => Some(l),
                        _ => None,
                    })
                    .collect();
                render_list_group_html(&lists)
            },
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::format::DocumentFormat;

    fn simple_ir(elements: Vec<Element>) -> DocumentIR {
        DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Docx,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                title: None,
                elements,
                ..Default::default()
            }],
            defined_names: Vec::new(),
        }
    }

    fn para(text: &str) -> Element {
        Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan::plain(text))],
            ..Default::default()
        })
    }

    fn span(text: &str) -> InlineContent {
        InlineContent::Text(TextSpan::plain(text))
    }

    /// A field added as a sibling of `Section::elements`
    /// (not inside it, like `speaker_notes`) is invisible to any renderer
    /// that was never explicitly taught about it, and nothing enforces
    /// that every renderer was. Adding `speaker_notes` broke two
    /// of four consumers (render_section_html and the CLI's `ir` JSON
    /// projection) silently — a corpus sweep found them, not the unit
    /// suite. This test populates every text-bearing sibling field at
    /// once and asserts each one reaches all three IR-level rendering
    /// surfaces, so a future field with the same shape fails a fast unit
    /// test instead of needing a multi-thousand-file corpus diff.
    #[test]
    fn test_maximal_section_reaches_every_rendering_surface() {
        let hf = |marker: &str| {
            Some(HeaderFooter {
                content: vec![para(marker)],
            })
        };
        let section = Section {
            title: Some("SECTION_TITLE_MARKER".to_string()),
            elements: vec![para("BODY_MARKER")],
            header: hf("HEADER_MARKER"),
            footer: hf("FOOTER_MARKER"),
            first_page_header: hf("FIRST_HEADER_MARKER"),
            first_page_footer: hf("FIRST_FOOTER_MARKER"),
            even_page_header: hf("EVEN_HEADER_MARKER"),
            even_page_footer: hf("EVEN_FOOTER_MARKER"),
            speaker_notes: Some(vec![para("SPEAKER_NOTES_MARKER")]),
            ..Default::default()
        };
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Pptx,
                title: None,
                ..Default::default()
            },
            sections: vec![section],
            defined_names: Vec::new(),
        };

        let markers = [
            "BODY_MARKER",
            "HEADER_MARKER",
            "FOOTER_MARKER",
            "FIRST_HEADER_MARKER",
            "FIRST_FOOTER_MARKER",
            "EVEN_HEADER_MARKER",
            "EVEN_FOOTER_MARKER",
            "SPEAKER_NOTES_MARKER",
        ];
        let plain = ir.plain_text();
        let markdown = ir.to_markdown();
        let html = ir.to_html();
        for marker in markers {
            assert!(plain.contains(marker), "plain_text() is missing {marker}: {plain:?}");
            assert!(markdown.contains(marker), "to_markdown() is missing {marker}: {markdown:?}");
            assert!(html.contains(marker), "to_html() is missing {marker}: {html:?}");
        }
    }

    #[test]
    fn test_plain_text_paragraph() {
        let ir = simple_ir(vec![para("Hello world")]);
        assert_eq!(ir.plain_text(), "Hello world");
    }

    #[test]
    fn test_markdown_heading() {
        let ir = simple_ir(vec![Element::Heading(Heading {
            level: 2,
            content: vec![span("Title")],
            ..Default::default()
        })]);
        assert_eq!(ir.to_markdown(), "## Title");
    }

    #[test]
    fn test_markdown_formatting() {
        let ir = simple_ir(vec![Element::Paragraph(Paragraph {
            content: vec![
                InlineContent::Text(TextSpan {
                    text: "bold".to_string(),
                    bold: true,
                    ..Default::default()
                }),
                InlineContent::Text(TextSpan::plain(" and ")),
                InlineContent::Text(TextSpan {
                    text: "italic".to_string(),
                    italic: true,
                    ..Default::default()
                }),
            ],
            ..Default::default()
        })]);
        assert_eq!(ir.to_markdown(), "**bold** and *italic*");
    }

    fn cell(text: &str) -> TableCell {
        TableCell {
            content: vec![Element::Paragraph(Paragraph {
                content: vec![span(text)],
                ..Default::default()
            })],
            col_span: 1,
            row_span: 1,
            ..Default::default()
        }
    }

    #[test]
    fn test_markdown_table() {
        let ir = simple_ir(vec![Element::Table(Table {
            rows: vec![
                TableRow {
                    cells: vec![cell("H1"), cell("H2")],
                    is_header: true,
                    ..Default::default()
                },
                TableRow {
                    cells: vec![cell("A"), cell("B")],
                    is_header: false,
                    ..Default::default()
                },
            ],
            ..Default::default()
        })]);
        let md = ir.to_markdown();
        assert!(md.contains("| H1 | H2 |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| A | B |"));
    }

    #[test]
    fn test_markdown_list() {
        let ir = simple_ir(vec![Element::List(List {
            ordered: false,
            items: vec![
                ListItem {
                    content: vec![para("First")],
                    nested: None,
                },
                ListItem {
                    content: vec![para("Second")],
                    nested: None,
                },
            ],
            ..Default::default()
        })]);
        assert_eq!(ir.to_markdown(), "- First\n- Second");
    }

    #[test]
    fn test_markdown_hyperlink() {
        let ir = simple_ir(vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan {
                text: "click".to_string(),
                hyperlink: Some("https://example.com".to_string()),
                ..Default::default()
            })],
            ..Default::default()
        })]);
        assert_eq!(ir.to_markdown(), "[click](https://example.com)");
    }

    #[test]
    fn test_multi_section_separator() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Xlsx,
                title: None,
                ..Default::default()
            },
            sections: vec![
                Section {
                    title: Some("Sheet1".to_string()),
                    elements: vec![para("Data A")],
                    ..Default::default()
                },
                Section {
                    title: Some("Sheet2".to_string()),
                    elements: vec![para("Data B")],
                    ..Default::default()
                },
            ],
            defined_names: Vec::new(),
        };
        let plain = ir.plain_text();
        assert!(plain.contains("Sheet1"));
        assert!(plain.contains("Data A"));
        // A form feed, not markdown's `---`: this is the plain-text renderer.
        assert!(plain.contains(PLAIN_BREAK));
        assert!(!plain.contains("---"));
        assert!(plain.contains("Data B"));
    }

    #[test]
    fn test_html_paragraph() {
        let ir = simple_ir(vec![para("Hello world")]);
        assert_eq!(ir.to_html(), "<p>Hello world</p>");
    }

    #[test]
    fn test_html_formatting() {
        let ir = simple_ir(vec![Element::Paragraph(Paragraph {
            content: vec![
                InlineContent::Text(TextSpan {
                    text: "bold".to_string(),
                    bold: true,
                    ..Default::default()
                }),
                InlineContent::Text(TextSpan::plain(" and ")),
                InlineContent::Text(TextSpan {
                    text: "link".to_string(),
                    hyperlink: Some("https://example.com".to_string()),
                    ..Default::default()
                }),
            ],
            ..Default::default()
        })]);
        assert_eq!(
            ir.to_html(),
            "<p><strong>bold</strong> and <a href=\"https://example.com\">link</a></p>"
        );
    }

    #[test]
    fn test_html_escaping() {
        let ir = simple_ir(vec![para("<script>alert('xss')</script>")]);
        assert!(ir.to_html().contains("&lt;script&gt;"));
        assert!(!ir.to_html().contains("<script>"));
    }

    #[test]
    fn test_html_table() {
        let ir = simple_ir(vec![Element::Table(Table {
            rows: vec![TableRow {
                cells: vec![cell("A")],
                is_header: true,
                ..Default::default()
            }],
            ..Default::default()
        })]);
        let html = ir.to_html();
        assert!(html.contains("<table>"));
        assert!(html.contains("<th>"));
        assert!(html.contains("A"));
    }

    #[test]
    fn test_html_list() {
        let ir = simple_ir(vec![Element::List(List {
            ordered: true,
            items: vec![
                ListItem {
                    content: vec![para("First")],
                    nested: None,
                },
                ListItem {
                    content: vec![para("Second")],
                    nested: None,
                },
            ],
            ..Default::default()
        })]);
        let html = ir.to_html();
        assert!(html.contains("<ol>"));
        assert!(html.contains("<li><p>First</p></li>"));
        assert!(html.contains("<li><p>Second</p></li>"));
    }

    /// `to_html()` had no image-embedding option at all, so an image
    /// never reached the HTML surface — only a `<figcaption>` when it
    /// happened to carry alt text.
    #[test]
    fn test_html_image_base64_embedding() {
        // A one-pixel PNG's magic bytes are enough: the renderer only needs
        // a media type and the bytes.
        let png = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0x01, 0x02];
        let ir = simple_ir(vec![Element::Image(Image {
            alt_text: Some("A <chart>".into()),
            data: Some(png.clone()),
            ..Default::default()
        })]);

        // Default behaviour is unchanged: a caption, no image data.
        let plain = ir.to_html();
        assert!(plain.contains("<figcaption>"), "{plain}");
        assert!(!plain.contains("<img"), "{plain}");

        let embedded = ir.to_html_with(HtmlOptions {
            image_embed: ImageEmbed::Base64,
        });
        assert!(embedded.contains("<img src=\"data:image/png;base64,"), "{embedded}");
        assert!(embedded.contains(&crate::core::base64::encode(&png)), "{embedded}");
        // Alt text is escaped, not injected.
        assert!(embedded.contains("alt=\"A &lt;chart&gt;\""), "{embedded}");

        // `format` is authoritative when the converter set it.
        let jpeg = simple_ir(vec![Element::Image(Image {
            data: Some(vec![0x00, 0x01, 0x02]),
            format: Some(ImageFormat::Jpeg),
            ..Default::default()
        })]);
        assert!(
            jpeg.to_html_with(HtmlOptions {
                image_embed: ImageEmbed::Base64,
            })
            .contains("data:image/jpeg;base64,"),
        );

        // Unknown bytes with no declared format fall back rather than
        // emitting a data URI a browser cannot render.
        let unknown = simple_ir(vec![Element::Image(Image {
            alt_text: Some("mystery".into()),
            data: Some(vec![0x00, 0x01, 0x02, 0x03]),
            ..Default::default()
        })]);
        let html = unknown.to_html_with(HtmlOptions {
            image_embed: ImageEmbed::Base64,
        });
        assert!(!html.contains("<img"), "{html}");
        assert!(html.contains("mystery"), "{html}");

        // The option does not leak into a later default render.
        assert!(!ir.to_html().contains("<img"));
    }

    /// numId fragmentation splits visually-continuous bullets into one
    /// `Element::List` each; HTML emitted a separate one-item `<ul>` per
    /// fragment where markdown showed one continuous list.
    #[test]
    fn test_adjacent_same_style_lists_merge_in_html() {
        let bullet = |text: &str| {
            Element::List(List {
                ordered: false,
                style: Some(ListStyle::Bullet),
                items: vec![ListItem {
                    content: vec![para(text)],
                    nested: None,
                }],
                ..Default::default()
            })
        };

        let ir = simple_ir(vec![bullet("one"), bullet("two"), bullet("three")]);
        let html = ir.to_html();
        assert_eq!(html.matches("<ul>").count(), 1, "html: {html}");
        assert_eq!(html.matches("</ul>").count(), 1, "html: {html}");
        assert_eq!(html.matches("<li>").count(), 3, "html: {html}");
        // Order is preserved.
        let pos = |needle: &str| html.find(needle).expect(needle);
        assert!(pos("one") < pos("two") && pos("two") < pos("three"), "html: {html}");

        // Intervening non-list content keeps the lists apart.
        let split = simple_ir(vec![bullet("one"), para("interruption"), bullet("two")]);
        assert_eq!(split.to_html().matches("<ul>").count(), 2, "{}", split.to_html());

        // A different marker style is a different list.
        let other_style = Element::List(List {
            ordered: false,
            style: Some(ListStyle::Square),
            items: vec![ListItem {
                content: vec![para("sq")],
                nested: None,
            }],
            ..Default::default()
        });
        let mixed = simple_ir(vec![bullet("one"), other_style]);
        assert_eq!(mixed.to_html().matches("<ul>").count(), 2, "{}", mixed.to_html());

        // Ordered vs unordered never merge.
        let numbered = Element::List(List {
            ordered: true,
            items: vec![ListItem {
                content: vec![para("n")],
                nested: None,
            }],
            ..Default::default()
        });
        let mixed = simple_ir(vec![bullet("one"), numbered]);
        let html = mixed.to_html();
        assert_eq!(html.matches("<ul>").count(), 1, "{html}");
        assert_eq!(html.matches("<ol>").count(), 1, "{html}");

        // An ordered list with its own explicit start is a deliberate
        // restart and keeps its own block (and its `start` attribute).
        let restart = |n: u32| {
            Element::List(List {
                ordered: true,
                start_number: Some(n),
                items: vec![ListItem {
                    content: vec![para("i")],
                    nested: None,
                }],
                ..Default::default()
            })
        };
        let restarted = simple_ir(vec![restart(1), restart(5)]);
        let html = restarted.to_html();
        assert_eq!(html.matches("<ol").count(), 2, "{html}");
        assert!(html.contains("<ol start=\"5\">"), "{html}");
    }

    /// `render_inline_html` read only bold/italic/strikethrough/
    /// hyperlink, so underline, super/subscript, highlight, colour, font and
    /// the caps variants vanished from `to_html()` while the IR carried them.
    #[test]
    fn test_html_span_formatting_is_not_dropped() {
        let styled = |f: fn(&mut TextSpan)| {
            let mut s = TextSpan::plain("X");
            f(&mut s);
            simple_ir(vec![Element::Paragraph(Paragraph {
                content: vec![InlineContent::Text(s)],
                ..Default::default()
            })])
            .to_html()
        };

        // Underline: `<u>`, and a distinguishable style for the variants a
        // bare `<u>` cannot express.
        let html = styled(|s| s.underline = Some(UnderlineStyle::Single));
        assert_eq!(html, "<p><u>X</u></p>");
        let html = styled(|s| s.underline = Some(UnderlineStyle::Double));
        assert!(html.contains("text-decoration-style:double"), "{html}");
        // An explicit "no underline" must not produce one.
        let html = styled(|s| s.underline = Some(UnderlineStyle::None));
        assert_eq!(html, "<p>X</p>");

        // Super/subscript, mirroring what the markdown renderer already did.
        let html = styled(|s| s.vertical_align = Some(VerticalAlign::Superscript));
        assert_eq!(html, "<p><sup>X</sup></p>");
        let html = styled(|s| s.vertical_align = Some(VerticalAlign::Subscript));
        assert_eq!(html, "<p><sub>X</sub></p>");
        let html = styled(|s| s.vertical_align = Some(VerticalAlign::Baseline));
        assert_eq!(html, "<p>X</p>");

        // Highlight, colour, font and caps all land in one style span.
        let html = styled(|s| s.highlight = Some([255, 255, 0]));
        assert!(html.contains("background-color:#FFFF00"), "{html}");
        let html = styled(|s| s.color = Some([17, 34, 51]));
        assert!(html.contains("color:#112233"), "{html}");
        let html = styled(|s| s.font_name = Some("Times New Roman".into()));
        assert!(html.contains("font-family:'Times New Roman'"), "{html}");
        let html = styled(|s| s.font_size_half_pt = Some(24));
        assert!(html.contains("font-size:12pt"), "{html}");
        let html = styled(|s| s.font_size_half_pt = Some(25));
        assert!(html.contains("font-size:12.5pt"), "{html}");
        let html = styled(|s| s.all_caps = true);
        assert!(html.contains("text-transform:uppercase"), "{html}");
        let html = styled(|s| s.small_caps = true);
        assert!(html.contains("font-variant:small-caps"), "{html}");

        // A plain span still renders bare — no empty wrapper.
        assert_eq!(styled(|_| {}), "<p>X</p>");
    }

    /// The new attribute values must not be able to break out of the
    /// `style="…"` they sit in, and text content stays escaped.
    #[test]
    fn test_html_span_style_values_are_escaped() {
        let ir = simple_ir(vec![Element::Paragraph(Paragraph {
            content: vec![InlineContent::Text(TextSpan {
                text: "<b>&hi</b>".into(),
                font_name: Some("Evil'; color:red; x:'".into()),
                underline: Some(UnderlineStyle::Single),
                ..Default::default()
            })],
            ..Default::default()
        })]);
        let html = ir.to_html();
        // Text content is still escaped.
        assert!(html.contains("&lt;b&gt;&amp;hi&lt;/b&gt;"), "{html}");
        assert!(!html.contains("<b>"), "{html}");
        // Nothing that could close the quoted value or start another
        // declaration survives into the attribute.
        let open = html.find("style=\"").expect("a style attribute") + "style=\"".len();
        let close = open + html[open..].find('"').expect("a closing quote");
        assert_eq!(&html[open..close], "font-family:'Evil colorred x'");
    }

    /// `List.start_number` was parsed and then ignored by both
    /// renderers, so a list the document starts at 3 rendered as 1, 2, 3.
    #[test]
    fn test_list_start_number_is_honoured() {
        let list = List {
            ordered: true,
            start_number: Some(3),
            items: vec![
                ListItem {
                    content: vec![para("A")],
                    nested: None,
                },
                ListItem {
                    content: vec![para("B")],
                    nested: None,
                },
                ListItem {
                    content: vec![para("C")],
                    nested: None,
                },
            ],
            ..Default::default()
        };
        let ir = simple_ir(vec![Element::List(list)]);

        let html = ir.to_html();
        assert!(html.contains("<ol start=\"3\">"), "html: {html}");

        let md = ir.to_markdown();
        assert!(md.contains("3. A"), "md: {md}");
        assert!(md.contains("4. B"), "md: {md}");
        assert!(md.contains("5. C"), "md: {md}");
        assert!(!md.contains("1. A"), "md still starts at 1: {md}");

        // An unordered list never gets a `start`, and a list starting at the
        // attribute's own default keeps a bare `<ol>`.
        let plain_ol = simple_ir(vec![Element::List(List {
            ordered: true,
            start_number: Some(1),
            items: vec![ListItem {
                content: vec![para("A")],
                nested: None,
            }],
            ..Default::default()
        })]);
        assert!(plain_ol.to_html().contains("<ol>"), "{}", plain_ol.to_html());

        let bullets = simple_ir(vec![Element::List(List {
            ordered: false,
            start_number: Some(7),
            items: vec![ListItem {
                content: vec![para("A")],
                nested: None,
            }],
            ..Default::default()
        })]);
        assert!(!bullets.to_html().contains("start="), "{}", bullets.to_html());
        assert!(bullets.to_markdown().contains("- A"), "{}", bullets.to_markdown());
    }

    // ── Defaults centralized in `block_default` ──────────────────────

    #[test]
    fn test_thematic_break_renders_as_a_form_feed_in_plain() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        assert_eq!(ir.plain_text(), PLAIN_BREAK);
    }

    #[test]
    fn test_thematic_break_renders_in_markdown() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        assert!(ir.to_markdown().contains("---"));
    }

    #[test]
    fn test_page_break_invisible_in_plain() {
        // PageBreak/ColumnBreak/Shape/Image have no plain-text counterpart
        // — they collapse to empty so plain_text shows only the surrounding
        // content.
        let ir = simple_ir(vec![para("before"), Element::PageBreak, para("after")]);
        let plain = ir.plain_text();
        assert!(plain.contains("before"));
        assert!(plain.contains("after"));
    }

    #[test]
    fn test_shape_invisible_in_plain() {
        let ir = simple_ir(vec![
            para("before"),
            Element::Shape(Shape::default()),
            para("after"),
        ]);
        let plain = ir.plain_text();
        assert!(plain.contains("before"));
        assert!(plain.contains("after"));
    }

    #[test]
    fn test_text_box_recursively_renders_children() {
        let ir = simple_ir(vec![Element::TextBox(TextBox {
            content: vec![para("inside")],
            ..Default::default()
        })]);
        let plain = ir.plain_text();
        assert!(plain.contains("inside"), "plain: {plain}");
    }

    #[test]
    fn test_html_thematic_break() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        let html = ir.to_html();
        assert!(html.contains("<hr"), "html: {html}");
    }
}

#[cfg(test)]
mod speaker_notes_render_tests {
    use super::*;

    /// All three renderers must surface speaker notes. Moving notes out of
    /// `elements` fixed the leak on the write side but dropped them from
    /// HTML, which a corpus sweep against v0.1.10 caught.
    #[test]
    fn test_every_renderer_surfaces_speaker_notes() {
        let ir = DocumentIR {
            sections: vec![Section {
                elements: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan {
                        text: "VisibleBody".into(),
                        ..Default::default()
                    })],
                    ..Default::default()
                })],
                speaker_notes: Some(vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan {
                        text: "NoteText".into(),
                        ..Default::default()
                    })],
                    ..Default::default()
                })]),
                ..Default::default()
            }],
            ..Default::default()
        };
        for (name, out) in [
            ("plain", ir.plain_text()),
            ("markdown", ir.to_markdown()),
            ("html", ir.to_html()),
        ] {
            assert!(out.contains("VisibleBody"), "{name}: body text lost");
            assert!(out.contains("NoteText"), "{name}: speaker notes lost:\n{out}");
        }
    }
}
