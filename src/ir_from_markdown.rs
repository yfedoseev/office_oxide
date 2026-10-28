//! Parse Markdown text into a `DocumentIR`.
//!
//! Handles the subset of Markdown that document extraction pipelines and
//! office_oxide's own `to_markdown()` produce: ATX headings, pipe tables,
//! bullet/numbered lists, thematic breaks, and paragraphs with bold/italic
//! inline spans.  This is not a full CommonMark implementation — it is
//! intentionally minimal so it carries no extra dependencies.

use crate::format::DocumentFormat;
use crate::ir::{
    BorderLine, BorderStyle, DocumentIR, Element, Heading, InlineContent, List, ListItem, Metadata,
    Paragraph, Section, Table, TableBorder, TableCell, TableRow, TextSpan,
};

impl DocumentIR {
    /// Parse Markdown text into a `DocumentIR`.
    ///
    /// Sections are separated by `---` horizontal rules (common for page
    /// boundaries in extracted documents).  Each ATX heading that is not immediately inside
    /// a list or table also acts as a natural section boundary.
    ///
    /// # Example
    ///
    /// ```rust
    /// use office_oxide::ir::DocumentIR;
    /// use office_oxide::format::DocumentFormat;
    ///
    /// let md = "# Title\n\nHello **world**.\n\n- item one\n- item two\n";
    /// let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
    /// assert!(!ir.sections.is_empty());
    /// ```
    ///
    /// At most [`crate::limits::max_markdown_list_items`] list items are
    /// kept (each is about 250 bytes of IR, from as little as four bytes of
    /// Markdown); the rest are dropped, a notice paragraph ends the last
    /// section, and `Metadata::text_truncated` and a warning say so.
    /// The document as a whole is bounded the same way by
    /// [`crate::limits::max_markdown_elements`]: past it the rest of the
    /// Markdown is not read.
    pub fn from_markdown(markdown: &str, format: DocumentFormat) -> Self {
        Self::from_markdown_with_limits(
            markdown,
            format,
            crate::limits::max_markdown_list_items(),
            crate::limits::max_markdown_elements(),
        )
    }

    #[cfg(test)]
    pub(crate) fn from_markdown_with_list_limit(
        markdown: &str,
        format: DocumentFormat,
        max_list_items: usize,
    ) -> Self {
        Self::from_markdown_with_limits(
            markdown,
            format,
            max_list_items,
            crate::limits::max_markdown_elements(),
        )
    }

    pub(crate) fn from_markdown_with_limits(
        markdown: &str,
        format: DocumentFormat,
        max_list_items: usize,
        max_elements: usize,
    ) -> Self {
        let mut parser = MarkdownParser::new(markdown, max_list_items, max_elements);
        let mut sections = parser.parse_sections();
        let mut metadata = Metadata {
            format,
            title: None,
            ..Default::default()
        };
        let mut notices: Vec<String> = Vec::new();
        if parser.dropped_list_items > 0 {
            notices.push(format!(
                "[{} list item(s) dropped: the document exceeds the {} list item limit \
                 (office_oxide::limits::set_max_markdown_list_items)]",
                parser.dropped_list_items, max_list_items
            ));
        }
        if parser.elements_exhausted {
            notices.push(format!(
                "[the rest of the document was dropped: it exceeds the {max_elements} block \
                 limit (office_oxide::limits::set_max_markdown_elements)]"
            ));
        }
        for notice in notices {
            metadata.text_truncated = true;
            metadata.warnings.push(notice.clone());
            if sections.is_empty() {
                sections.push(Section::default());
            }
            if let Some(last) = sections.last_mut() {
                last.elements.push(Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan::plain(notice))],
                    ..Default::default()
                }));
            }
        }
        DocumentIR { metadata, sections }
    }
}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

struct MarkdownParser<'a> {
    lines: Vec<&'a str>,
    pos: usize,
    /// List items still allowed, over the whole document.
    list_items_left: usize,
    /// List items read and dropped for want of allowance.
    dropped_list_items: usize,
    /// Block allowance still left, over the whole document.
    elements_left: usize,
    /// A block did not fit; the rest of the document is not read.
    elements_exhausted: bool,
}

impl<'a> MarkdownParser<'a> {
    fn new(src: &'a str, max_list_items: usize, max_elements: usize) -> Self {
        Self {
            elements_left: max_elements,
            elements_exhausted: false,
            lines: src.lines().collect(),
            pos: 0,
            list_items_left: max_list_items,
            dropped_list_items: 0,
        }
    }

    /// Spend `cost` blocks of the allowance before a block is built. When
    /// they do not fit, the allowance is closed for good: the document is
    /// cut here rather than continued with later, smaller blocks.
    fn charge(&mut self, cost: usize) -> bool {
        if self.elements_exhausted {
            return false;
        }
        if cost > self.elements_left {
            self.elements_exhausted = true;
            return false;
        }
        self.elements_left -= cost;
        true
    }

    fn peek(&self) -> Option<&'a str> {
        self.lines.get(self.pos).copied()
    }

    fn advance(&mut self) -> Option<&'a str> {
        let line = self.lines.get(self.pos).copied();
        self.pos += 1;
        line
    }

    /// Parse the full document into sections, splitting on thematic `---`
    /// breaks and top-level H1/H2 headings.
    fn parse_sections(&mut self) -> Vec<Section> {
        let mut sections: Vec<Section> = Vec::new();
        let mut current = Section {
            title: None,
            elements: Vec::new(),
            ..Default::default()
        };

        while self.pos < self.lines.len() {
            if self.elements_exhausted {
                break;
            }
            let line = match self.peek() {
                Some(l) => l,
                None => break,
            };

            // Blank line
            if line.trim().is_empty() {
                self.advance();
                continue;
            }

            // Fenced code block. Without this the fence's language leaked
            // into the text as an ordinary paragraph ("rust let x = 1;") and
            // no Element::CodeBlock was ever produced from markdown.
            if let Some(lang) = fence_language(line) {
                if !self.charge(1) {
                    break;
                }
                self.advance();
                let mut body: Vec<String> = Vec::new();
                while let Some(l) = self.peek() {
                    if fence_language(l).is_some() {
                        self.advance();
                        break;
                    }
                    body.push(l.to_string());
                    self.advance();
                }
                current
                    .elements
                    .push(Element::CodeBlock(crate::ir::CodeBlock {
                        content: body.join("\n"),
                        language: lang,
                    }));
                continue;
            }

            // Thematic break `---` / `***` / `___`. It both records itself as
            // an `Element::ThematicBreak` (so it round-trips instead of being
            // silently swallowed) and, matching this parser's existing page-
            // boundary convention, starts a new section.
            if is_thematic_break(line) {
                if !self.charge(1) {
                    break;
                }
                self.advance();
                current.elements.push(Element::ThematicBreak);
                if !current.elements.is_empty() || current.title.is_some() {
                    if !self.charge(1) {
                        break;
                    }
                    sections.push(current);
                    current = Section {
                        title: None,
                        elements: Vec::new(),
                        ..Default::default()
                    };
                }
                continue;
            }

            // ATX heading
            if let Some((level, text)) = parse_atx_heading(line) {
                if !self.charge(1) {
                    break;
                }
                self.advance();
                // H1 headings start a new section
                if level == 1 {
                    if !current.elements.is_empty() || current.title.is_some() {
                        if !self.charge(1) {
                            break;
                        }
                        sections.push(current);
                    }
                    current = Section {
                        title: Some(text.clone()),
                        elements: Vec::new(),
                        ..Default::default()
                    };
                } else {
                    current.elements.push(Element::Heading(Heading {
                        level,
                        content: parse_inline(&text),
                        ..Default::default()
                    }));
                }
                continue;
            }

            // Pipe table
            if line.trim_start().starts_with('|') {
                if let Some(table) = self.parse_table() {
                    current.elements.push(Element::Table(table));
                    continue;
                }
                if self.elements_exhausted {
                    break;
                }
            }

            // Unordered list
            if is_unordered_list_marker(line) {
                if !self.charge(1) {
                    break;
                }
                let list = self.parse_list(false);
                current.elements.push(Element::List(list));
                continue;
            }

            // Ordered list
            if is_ordered_list_marker(line) {
                if !self.charge(1) {
                    break;
                }
                let list = self.parse_list(true);
                current.elements.push(Element::List(list));
                continue;
            }

            // Regular paragraph (accumulate until blank line or block element)
            if !self.charge(1) {
                break;
            }
            let para = self.parse_paragraph();
            if !para.content.is_empty() {
                current.elements.push(Element::Paragraph(para));
            }
        }

        if !current.elements.is_empty() || current.title.is_some() {
            sections.push(current);
        }

        // Ensure at least one section
        if sections.is_empty() {
            sections.push(Section {
                title: None,
                elements: Vec::new(),
                ..Default::default()
            });
        }

        sections
    }

    // -----------------------------------------------------------------------
    // Block parsers
    // -----------------------------------------------------------------------

    fn parse_paragraph(&mut self) -> Paragraph {
        let mut lines: Vec<&str> = Vec::new();
        loop {
            match self.peek() {
                None => break,
                Some(line) => {
                    if line.trim().is_empty()
                        || parse_atx_heading(line).is_some()
                        || is_thematic_break(line)
                        || line.trim_start().starts_with('|')
                        || is_unordered_list_marker(line)
                        || is_ordered_list_marker(line)
                    {
                        break;
                    }
                    lines.push(line);
                    self.advance();
                },
            }
        }
        let text = lines.join(" ");
        Paragraph {
            content: parse_inline(&text),
            ..Default::default()
        }
    }

    fn parse_table(&mut self) -> Option<Table> {
        // Collect all consecutive pipe lines
        let mut raw: Vec<&'a str> = Vec::new();
        while let Some(line) = self.peek() {
            if line.trim_start().starts_with('|') {
                raw.push(line);
                self.advance();
            } else {
                break;
            }
        }

        if raw.is_empty() {
            return None;
        }

        // Filter out alignment rows (cells that look like `---`, `:---`, `---:`)
        let data_rows: Vec<&str> = raw
            .iter()
            .copied()
            .filter(|line| !is_table_separator_row(line))
            .collect();

        if data_rows.is_empty() {
            return None;
        }

        // The table and each row's cells are charged before they are built.
        if !self.charge(1) {
            return None;
        }
        let mut rows: Vec<TableRow> = Vec::new();
        for (i, row_line) in data_rows.iter().enumerate() {
            let is_header = i == 0;
            let cell_texts = split_pipe_row(row_line);
            if !self.charge(cell_texts.len().saturating_mul(TABLE_CELL_COST)) {
                break;
            }
            let cells = cell_texts
                .into_iter()
                .map(|cell_text| {
                    // The header is marked by its rule and `repeat_as_header`,
                    // not by bold text: bolding would make the cell render as
                    // `**Col A**` and break the Markdown round trip.
                    let content = parse_inline(cell_text.trim());
                    TableCell {
                        content: vec![Element::Paragraph(Paragraph {
                            content,
                            ..Default::default()
                        })],
                        col_span: 1,
                        row_span: 1,
                        border: is_header.then(|| TableBorder {
                            top: None,
                            bottom: Some(border_line(HEADER_RULE_EIGHTHS_PT)),
                            left: None,
                            right: None,
                            inside_h: None,
                            inside_v: None,
                        }),
                        ..Default::default()
                    }
                })
                .collect();
            rows.push(TableRow {
                cells,
                is_header,
                repeat_as_header: is_header,
                ..Default::default()
            });
        }

        if rows.is_empty() {
            return None;
        }

        // A Markdown table has no styling of its own; without borders and a
        // distinct header row, a written table is unreadable.
        let thin = Some(border_line(CELL_RULE_EIGHTHS_PT));
        Some(Table {
            rows,
            border: Some(TableBorder {
                top: thin.clone(),
                bottom: thin.clone(),
                left: thin.clone(),
                right: thin.clone(),
                inside_h: thin.clone(),
                inside_v: thin,
            }),
            ..Default::default()
        })
    }

    fn parse_list(&mut self, ordered: bool) -> List {
        self.parse_list_at(ordered, indent_width(self.peek().unwrap_or("")))
    }

    /// Parse a list whose markers sit at `base_indent` columns.
    ///
    /// A more-indented marker starts a sub-list hanging off the previous item.
    /// `nested` used to be hardcoded to `None`, which flattened every level
    /// into one — and meant the markdown path could never exercise the
    /// nested-list handling in the writers.
    fn parse_list_at(&mut self, ordered: bool, base_indent: usize) -> List {
        let mut items: Vec<ListItem> = Vec::new();
        let mut start_number = None;
        while let Some(line) = self.peek() {
            let is_marker = if ordered {
                is_ordered_list_marker(line)
            } else {
                is_unordered_list_marker(line)
            };
            let nested_ordered = is_ordered_list_marker(line);
            let nested_unordered = is_unordered_list_marker(line);
            if !is_marker && !nested_ordered && !nested_unordered {
                break;
            }
            let indent = indent_width(line);
            if indent > base_indent {
                // Deeper marker: attach a sub-list to the item just parsed.
                let sub = self.parse_list_at(nested_ordered, indent);
                if let Some(last) = items.last_mut() {
                    last.nested = Some(sub);
                    continue;
                }
                break;
            }
            if indent < base_indent || !is_marker {
                break;
            }
            self.advance();
            if self.list_items_left == 0 {
                self.dropped_list_items += 1;
                continue;
            }
            self.list_items_left -= 1;
            // CommonMark §5.2: an ordered list starts at its first item's
            // number.
            if ordered && items.is_empty() {
                start_number = ordered_marker_number(line).filter(|&n| n != 1);
            }
            let content_str = strip_list_marker(line);
            items.push(ListItem {
                content: vec![Element::Paragraph(Paragraph {
                    content: parse_inline(content_str),
                    ..Default::default()
                })],
                nested: None,
            });
        }
        List {
            ordered,
            items,
            start_number,
            ..Default::default()
        }
    }
}

/// The number of an ordered-list marker line (`3. a` → 3).
fn ordered_marker_number(line: &str) -> Option<u32> {
    let t = line.trim_start();
    let digits = t.bytes().take_while(u8::is_ascii_digit).count();
    t[..digits].parse().ok()
}

// ---------------------------------------------------------------------------
// Inline parser (bold, italic, plain)
// ---------------------------------------------------------------------------

/// What a table cell costs against [`crate::limits::max_markdown_elements`]:
/// about four times the IR of a paragraph.
const TABLE_CELL_COST: usize = 4;

/// Cell rule width, in eighths of a point (0.5 pt).
const CELL_RULE_EIGHTHS_PT: u32 = 4;
/// Rule under the header row, in eighths of a point (1.5 pt).
const HEADER_RULE_EIGHTHS_PT: u32 = 12;

fn border_line(size_eighths_pt: u32) -> BorderLine {
    BorderLine {
        style: BorderStyle::Single,
        color: None,
        size: Some(size_eighths_pt),
        space: None,
    }
}

/// Finds the next occurrence of one marker, remembering the last answer.
///
/// The inline parser looks for a closing marker from every opener, so an
/// input of many openers and no closer rescanned the rest of the line each
/// time. The first occurrence at or after `from` is `hit` (or there is none),
/// which is also the first occurrence at or after any later start up to
/// `hit`, so a search that starts inside the remembered range is answered
/// without scanning. Starts only move forward, so the scans of one marker
/// cover disjoint stretches of the text and the whole parse is linear.
struct NextMarker {
    needle: &'static str,
    from: usize,
    hit: Option<usize>,
    primed: bool,
}

impl NextMarker {
    const fn new(needle: &'static str) -> Self {
        Self {
            needle,
            from: 0,
            hit: None,
            primed: false,
        }
    }

    /// The byte offset in `text` of the first `needle` at or after `start`.
    fn find(&mut self, text: &str, start: usize) -> Option<usize> {
        if self.primed && start >= self.from && self.hit.is_none_or(|h| h >= start) {
            return self.hit;
        }
        self.primed = true;
        self.from = start;
        self.hit = text[start..].find(self.needle).map(|p| start + p);
        self.hit
    }
}

fn parse_inline(text: &str) -> Vec<InlineContent> {
    let mut out: Vec<InlineContent> = Vec::new();
    let bytes = text.as_bytes();
    let len = text.len();
    let mut plain_start = 0usize;
    let mut star2 = NextMarker::new("**");
    let mut under2 = NextMarker::new("__");
    let mut star1 = NextMarker::new("*");
    let mut under1 = NextMarker::new("_");
    let mut tilde2 = NextMarker::new("~~");
    let mut tick = NextMarker::new("`");
    let mut close_bracket = NextMarker::new("]");
    let mut close_paren = NextMarker::new(")");

    macro_rules! flush_plain {
        ($end:expr) => {
            if plain_start < $end {
                let t = &text[plain_start..$end];
                if !t.is_empty() {
                    out.push(InlineContent::Text(TextSpan::plain(t)));
                }
            }
        };
    }

    let mut i = 0usize;
    while i < len {
        // Bold: **text** or __text__
        if i + 1 < len
            && ((bytes[i] == b'*' && bytes[i + 1] == b'*')
                || (bytes[i] == b'_' && bytes[i + 1] == b'_'))
        {
            let found = if bytes[i] == b'*' {
                star2.find(text, i + 2)
            } else {
                under2.find(text, i + 2)
            };
            if let Some(close) = found {
                let end = close - (i + 2);
                flush_plain!(i);
                let inner = &text[i + 2..i + 2 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    bold: true,
                    ..Default::default()
                }));
                i += 2 + end + 2;
                plain_start = i;
                continue;
            }
        }

        // Italic: *text* or _text_
        if (bytes[i] == b'*' || bytes[i] == b'_') && i + 1 < len && bytes[i + 1] != bytes[i] {
            let found = if bytes[i] == b'*' {
                star1.find(text, i + 1)
            } else {
                under1.find(text, i + 1)
            };
            if let Some(close) = found {
                let end = close - (i + 1);
                flush_plain!(i);
                let inner = &text[i + 1..i + 1 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    italic: true,
                    ..Default::default()
                }));
                i += 1 + end + 1;
                plain_start = i;
                continue;
            }
        }

        // Strikethrough: ~~text~~
        if i + 1 < len && bytes[i] == b'~' && bytes[i + 1] == b'~' {
            if let Some(close) = tilde2.find(text, i + 2) {
                let end = close - (i + 2);
                flush_plain!(i);
                let inner = &text[i + 2..i + 2 + end];
                out.push(InlineContent::Text(TextSpan {
                    text: inner.to_string(),
                    strikethrough: true,
                    ..Default::default()
                }));
                i += 2 + end + 2;
                plain_start = i;
                continue;
            }
        }

        // Inline code: `code` — strip backticks, treat as plain
        if bytes[i] == b'`' {
            if let Some(close) = tick.find(text, i + 1) {
                let end = close - (i + 1);
                flush_plain!(i);
                let inner = &text[i + 1..i + 1 + end];
                out.push(InlineContent::Text(TextSpan::plain(inner)));
                i += 1 + end + 1;
                plain_start = i;
                continue;
            }
        }

        // Markdown link: [text](url)
        if bytes[i] == b'[' {
            if let Some(close) = close_bracket.find(text, i + 1) {
                let bracket_end = close - (i + 1);
                let after_bracket = i + 1 + bracket_end + 1;
                if after_bracket < len && bytes[after_bracket] == b'(' {
                    if let Some(close) = close_paren.find(text, after_bracket + 1) {
                        let paren_end = close - (after_bracket + 1);
                        flush_plain!(i);
                        let link_text = &text[i + 1..i + 1 + bracket_end];
                        let url = &text[after_bracket + 1..after_bracket + 1 + paren_end];
                        out.push(InlineContent::Text(TextSpan {
                            text: link_text.to_string(),
                            hyperlink: Some(url.to_string()),
                            ..Default::default()
                        }));
                        i = after_bracket + 1 + paren_end + 1;
                        plain_start = i;
                        continue;
                    }
                }
            }
        }

        i += text[i..].chars().next().map(|c| c.len_utf8()).unwrap_or(1);
    }

    flush_plain!(len);

    out
}

// ---------------------------------------------------------------------------
// Line classifiers
// ---------------------------------------------------------------------------

fn parse_atx_heading(line: &str) -> Option<(u8, String)> {
    let trimmed = line.trim_start();
    let hashes = trimmed.bytes().take_while(|&b| b == b'#').count();
    if hashes == 0 || hashes > 6 {
        return None;
    }
    let rest = &trimmed[hashes..];
    if rest.is_empty() || rest.starts_with(' ') || rest.starts_with('\t') {
        // CommonMark §4.2: an optional closing run of `#` counts only when
        // it is the whole content or a space precedes it — the `#` of
        // `C#` is text.
        let content = rest.trim();
        let open = content.trim_end_matches('#');
        let text = if open.is_empty() {
            ""
        } else if open.ends_with([' ', '\t']) {
            open.trim_end()
        } else {
            content
        };
        Some((hashes as u8, text.to_string()))
    } else {
        None
    }
}

/// Leading-whitespace width of `line`, with a tab counted as four columns.
fn indent_width(line: &str) -> usize {
    line.chars()
        .take_while(|c| c.is_whitespace())
        .map(|c| if c == '\t' { 4 } else { 1 })
        .sum()
}

/// `Some(language)` when `line` opens or closes a fenced code block.
/// The language is `None` for a bare fence.
fn fence_language(line: &str) -> Option<Option<String>> {
    let t = line.trim_start();
    let rest = t.strip_prefix("```").or_else(|| t.strip_prefix("~~~"))?;
    let lang = rest.trim();
    Some(if lang.is_empty() {
        None
    } else {
        Some(lang.to_string())
    })
}

fn is_thematic_break(line: &str) -> bool {
    let t = line.trim();
    if t.len() < 3 {
        return false;
    }
    let Some(ch) = t.chars().next() else {
        return false;
    };
    if !matches!(ch, '-' | '*' | '_') {
        return false;
    }
    t.chars().all(|c| c == ch || c == ' ') && t.chars().filter(|&c| c == ch).count() >= 3
}

fn is_table_separator_row(line: &str) -> bool {
    let trimmed = line.trim().trim_matches('|');
    trimmed.split('|').all(|cell| {
        let c = cell.trim().trim_start_matches(':').trim_end_matches(':');
        !c.is_empty() && c.bytes().all(|b| b == b'-')
    })
}

fn split_pipe_row(line: &str) -> Vec<&str> {
    let inner = line.trim().trim_start_matches('|').trim_end_matches('|');
    inner.split('|').collect()
}

fn is_unordered_list_marker(line: &str) -> bool {
    let t = line.trim_start();
    (t.starts_with("- ") || t.starts_with("* ") || t.starts_with("+ ")) && !is_thematic_break(line)
}

fn is_ordered_list_marker(line: &str) -> bool {
    let t = line.trim_start();
    // e.g. "1. " "12. " "1) "
    let num_end = t.bytes().take_while(|b| b.is_ascii_digit()).count();
    if num_end == 0 {
        return false;
    }
    let after = &t[num_end..];
    after.starts_with(". ") || after.starts_with(") ")
}

fn strip_list_marker(line: &str) -> &str {
    let t = line.trim_start();
    if t.starts_with("- ") || t.starts_with("* ") || t.starts_with("+ ") {
        t[2..].trim_start()
    } else {
        // ordered: skip digits + ". " or ") "
        let num_end = t.bytes().take_while(|b| b.is_ascii_digit()).count();
        t[num_end + 2..].trim_start()
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::format::DocumentFormat;

    /// The quadratic parser `parse_inline` replaced, kept as the reference
    /// its output must equal.
    fn parse_inline_reference(text: &str) -> Vec<InlineContent> {
        let mut out: Vec<InlineContent> = Vec::new();
        let bytes = text.as_bytes();
        let len = text.len();
        let mut plain_start = 0usize;

        macro_rules! flush_plain {
            ($end:expr) => {
                if plain_start < $end {
                    let t = &text[plain_start..$end];
                    if !t.is_empty() {
                        out.push(InlineContent::Text(TextSpan::plain(t)));
                    }
                }
            };
        }

        let mut i = 0usize;
        while i < len {
            // Bold: **text** or __text__
            if i + 1 < len
                && ((bytes[i] == b'*' && bytes[i + 1] == b'*')
                    || (bytes[i] == b'_' && bytes[i + 1] == b'_'))
            {
                let marker = &text[i..i + 2];
                if let Some(end) = text[i + 2..].find(marker) {
                    flush_plain!(i);
                    let inner = &text[i + 2..i + 2 + end];
                    out.push(InlineContent::Text(TextSpan {
                        text: inner.to_string(),
                        bold: true,
                        ..Default::default()
                    }));
                    i += 2 + end + 2;
                    plain_start = i;
                    continue;
                }
            }

            // Italic: *text* or _text_
            if (bytes[i] == b'*' || bytes[i] == b'_') && i + 1 < len && bytes[i + 1] != bytes[i] {
                let marker = &text[i..i + 1];
                if let Some(end) = text[i + 1..].find(marker) {
                    flush_plain!(i);
                    let inner = &text[i + 1..i + 1 + end];
                    out.push(InlineContent::Text(TextSpan {
                        text: inner.to_string(),
                        italic: true,
                        ..Default::default()
                    }));
                    i += 1 + end + 1;
                    plain_start = i;
                    continue;
                }
            }

            // Strikethrough: ~~text~~
            if i + 1 < len && bytes[i] == b'~' && bytes[i + 1] == b'~' {
                if let Some(end) = text[i + 2..].find("~~") {
                    flush_plain!(i);
                    let inner = &text[i + 2..i + 2 + end];
                    out.push(InlineContent::Text(TextSpan {
                        text: inner.to_string(),
                        strikethrough: true,
                        ..Default::default()
                    }));
                    i += 2 + end + 2;
                    plain_start = i;
                    continue;
                }
            }

            // Inline code: `code` — strip backticks, treat as plain
            if bytes[i] == b'`' {
                if let Some(end) = text[i + 1..].find('`') {
                    flush_plain!(i);
                    let inner = &text[i + 1..i + 1 + end];
                    out.push(InlineContent::Text(TextSpan::plain(inner)));
                    i += 1 + end + 1;
                    plain_start = i;
                    continue;
                }
            }

            // Markdown link: [text](url)
            if bytes[i] == b'[' {
                if let Some(bracket_end) = text[i + 1..].find(']') {
                    let after_bracket = i + 1 + bracket_end + 1;
                    if after_bracket < len && bytes[after_bracket] == b'(' {
                        if let Some(paren_end) = text[after_bracket + 1..].find(')') {
                            flush_plain!(i);
                            let link_text = &text[i + 1..i + 1 + bracket_end];
                            let url = &text[after_bracket + 1..after_bracket + 1 + paren_end];
                            out.push(InlineContent::Text(TextSpan {
                                text: link_text.to_string(),
                                hyperlink: Some(url.to_string()),
                                ..Default::default()
                            }));
                            i = after_bracket + 1 + paren_end + 1;
                            plain_start = i;
                            continue;
                        }
                    }
                }
            }

            i += text[i..].chars().next().map(|c| c.len_utf8()).unwrap_or(1);
        }

        flush_plain!(len);

        out
    }

    /// Every string over the characters the inline parser reacts to, up to
    /// length 8, parses exactly as it did when each opener rescanned the line.
    #[test]
    fn test_inline_parse_equals_the_reference_on_all_short_strings() {
        let alphabet = ['*', '_', '~', '`', '[', ']', '(', ')', 'a', 'é'];
        let mut buf: Vec<usize> = Vec::new();
        fn walk(alphabet: &[char], buf: &mut Vec<usize>, max: usize, check: &mut dyn FnMut(&str)) {
            let s: String = buf.iter().map(|&k| alphabet[k]).collect();
            check(&s);
            if buf.len() == max {
                return;
            }
            for k in 0..alphabet.len() {
                buf.push(k);
                walk(alphabet, buf, max, check);
                buf.pop();
            }
        }
        let mut checked = 0usize;
        walk(&alphabet, &mut buf, 5, &mut |s| {
            assert_eq!(parse_inline(s), parse_inline_reference(s), "input {s:?}");
            checked += 1;
        });
        // Longer strings from a small alphabet that exercises every marker.
        let small = ['*', '_', '~', '`', '[', ']', '(', ')'];
        let mut state = 7u32;
        for _ in 0..20_000 {
            let mut s = String::new();
            state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
            for _ in 0..(state >> 28) as usize + 4 {
                state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
                s.push(small[(state >> 16) as usize % small.len()]);
                if state & 0x100 != 0 {
                    s.push('x');
                }
            }
            assert_eq!(parse_inline(&s), parse_inline_reference(&s), "input {s:?}");
        }
        assert!(checked > 100_000);
    }

    /// An opener with no closer, repeated, used to rescan the rest of the
    /// line for each one.
    #[test]
    fn test_inline_parse_of_unclosed_openers_is_linear() {
        let start = std::time::Instant::now();
        for unit in ["[", "[](", "**", "__", "~~", "`", "*a", "_a", "[a]", "[a]("] {
            let text = unit.repeat(60_000 / unit.len());
            let out = parse_inline(&text);
            assert!(!out.is_empty(), "{unit:?}");
        }
        assert!(
            start.elapsed() < std::time::Duration::from_secs(5),
            "took {:?}",
            start.elapsed()
        );
    }

    #[test]
    fn test_parse_heading_paragraph() {
        let md = "# Hello\n\nSome text here.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Hello"));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn test_parse_page_break_into_sections() {
        let md = "# Page 1\n\nText one.\n\n---\n\n# Page 2\n\nText two.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 2);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Page 1"));
        assert_eq!(ir.sections[1].title.as_deref(), Some("Page 2"));
    }

    #[test]
    fn test_parse_unordered_list() {
        let md = "- apple\n- banana\n- cherry\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert!(!list.ordered);
        assert_eq!(list.items.len(), 3);
    }

    #[test]
    fn test_parse_ordered_list() {
        let md = "1. first\n2. second\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert!(list.ordered);
    }

    #[test]
    fn test_parse_pipe_table() {
        let md = "| Name | Age |\n|------|-----|\n| Alice | 30 |\n| Bob | 25 |\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let table = match &ir.sections[0].elements[0] {
            Element::Table(t) => t,
            other => panic!("expected Table, got {other:?}"),
        };
        assert_eq!(table.rows.len(), 3); // header + 2 data rows (separator stripped)
        assert!(table.rows[0].is_header);
    }

    #[test]
    fn test_parse_bold_italic_inline() {
        let md = "Hello **world** and *rust*.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let para = match &ir.sections[0].elements[0] {
            Element::Paragraph(p) => p,
            other => panic!("expected Paragraph, got {other:?}"),
        };
        let spans: Vec<_> = para
            .content
            .iter()
            .filter_map(|c| match c {
                InlineContent::Text(s) => Some(s),
                _ => None,
            })
            .collect();
        assert!(spans.iter().any(|s| s.bold && s.text == "world"));
        assert!(spans.iter().any(|s| s.italic && s.text == "rust"));
    }

    #[test]
    fn test_fenced_code_block_becomes_code_block_element() {
        // gap 1: the fence language used to leak into the text
        // as a plain paragraph ("rust let x = 1;") with no CodeBlock at all.
        let md = "```rust\nlet x = 1;\n```\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let block = match &ir.sections[0].elements[0] {
            Element::CodeBlock(c) => c,
            other => panic!("expected CodeBlock, got {other:?}"),
        };
        assert_eq!(block.language.as_deref(), Some("rust"));
        assert_eq!(block.content, "let x = 1;");
    }

    #[test]
    fn test_thematic_break_becomes_thematic_break_element() {
        // gap 2: `---` only ever split sections; no
        // Element::ThematicBreak was ever produced, so the mark itself
        // was silently dropped from the round trip.
        let md = "```rust\nlet x = 1;\n```\n\n---\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert!(
            ir.sections[0]
                .elements
                .iter()
                .any(|e| matches!(e, Element::ThematicBreak)),
            "expected an Element::ThematicBreak, got {:#?}",
            ir.sections[0].elements
        );
    }

    #[test]
    fn test_nested_bullet_list_preserves_nesting_depth() {
        // gap 3: src/ir_from_markdown.rs:281 used to hardcode
        // `nested: None`, flattening every item to ilvl=0.
        let md = "- item1\n  - sub1\n  - sub2\n- item2\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert_eq!(list.items.len(), 2);
        let nested = list.items[0]
            .nested
            .as_ref()
            .expect("first item should have a nested sub-list");
        assert_eq!(nested.items.len(), 2);
        assert!(list.items[1].nested.is_none());
    }

    /// CommonMark §4.2: a closing `#` sequence is stripped only when a
    /// space precedes it. `trim_end_matches('#')` also ate the `#` of
    /// `C#`.
    #[test]
    fn test_atx_heading_keeps_a_trailing_hash_that_is_part_of_the_text() {
        for (line, want) in [
            ("# Learn C#", "Learn C#"),
            ("## F# and C#", "F# and C#"),
            ("# Title ##", "Title"),
            ("# Title #   ", "Title"),
            ("### ###", ""),
            ("#", ""),
            ("# #hashtag", "#hashtag"),
        ] {
            assert_eq!(parse_atx_heading(line).map(|(_, t)| t).as_deref(), Some(want), "{line:?}");
        }
    }

    /// An ordered list's first number is its start (CommonMark §5.2);
    /// `3. a` was written as a list starting at 1.
    #[test]
    fn test_ordered_list_start_number_is_kept() {
        let ir =
            DocumentIR::from_markdown("3. a\n4. b\n\n- x\n  7) nested\n", DocumentFormat::Docx);
        let lists: Vec<&List> = ir.sections[0]
            .elements
            .iter()
            .filter_map(|e| match e {
                Element::List(l) => Some(l),
                _ => None,
            })
            .collect();
        assert_eq!(lists[0].start_number, Some(3));
        let nested = lists[1].items[0].nested.as_ref().expect("nested list");
        assert_eq!(nested.start_number, Some(7));
        // A list starting at 1 keeps the default.
        let ir = DocumentIR::from_markdown("1. a\n", DocumentFormat::Docx);
        let Element::List(ref l) = ir.sections[0].elements[0] else {
            panic!()
        };
        assert_eq!(l.start_number, None);
    }

    #[test]
    fn test_parse_empty_markdown() {
        let ir = DocumentIR::from_markdown("", DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert!(ir.sections[0].elements.is_empty());
    }

    fn item_count(l: &List) -> usize {
        l.items
            .iter()
            .map(|i| 1 + i.nested.as_ref().map_or(0, item_count))
            .sum()
    }

    #[test]
    fn test_list_items_past_the_limit_are_dropped_with_a_notice() {
        let md = "- x\n".repeat(100) + "\nafter\n";
        let ir = DocumentIR::from_markdown_with_list_limit(&md, DocumentFormat::Docx, 10);
        let kept: usize = ir.sections[0]
            .elements
            .iter()
            .map(|e| {
                if let Element::List(l) = e {
                    item_count(l)
                } else {
                    0
                }
            })
            .sum();
        assert_eq!(kept, 10);
        assert!(ir.metadata.text_truncated);
        assert!(ir.metadata.warnings[0].contains("90 list item(s) dropped"));
        let text = ir.plain_text();
        assert!(text.contains("after"), "text after the list is still read");
        assert!(text.contains("list item limit"), "the notice reaches every renderer");
    }

    #[test]
    fn test_nested_and_ordered_items_count_toward_the_limit() {
        let md = "1. a\n   - b\n   - c\n2. d\n";
        let ir = DocumentIR::from_markdown_with_list_limit(md, DocumentFormat::Docx, 3);
        let kept: usize = ir.sections[0]
            .elements
            .iter()
            .map(|e| {
                if let Element::List(l) = e {
                    item_count(l)
                } else {
                    0
                }
            })
            .sum();
        assert_eq!(kept, 3);
        assert!(ir.metadata.text_truncated);
    }

    #[test]
    fn test_a_list_within_the_limit_is_untouched() {
        let ir = DocumentIR::from_markdown_with_list_limit("- a\n- b\n", DocumentFormat::Docx, 2);
        assert!(!ir.metadata.text_truncated);
        assert!(ir.metadata.warnings.is_empty());
    }

    fn block_count(ir: &DocumentIR) -> usize {
        ir.sections.iter().map(|s| s.elements.len()).sum()
    }

    fn with_blocks(md: &str, max: usize) -> DocumentIR {
        DocumentIR::from_markdown_with_limits(md, DocumentFormat::Docx, usize::MAX, max)
    }

    #[test]
    fn test_paragraphs_past_the_block_limit_are_dropped_with_a_notice() {
        let md = "para\n\n".repeat(100) + "# late\n";
        let ir = with_blocks(&md, 10);
        // Ten paragraphs, then the notice.
        assert_eq!(block_count(&ir), 11);
        assert!(ir.metadata.text_truncated);
        assert!(ir.metadata.warnings[0].contains("set_max_markdown_elements"));
        assert!(ir.plain_text().contains("block limit"), "the notice reaches every renderer");
        assert!(!ir.plain_text().contains("late"), "nothing after the cut is read");
    }

    #[test]
    fn test_sections_past_the_block_limit_are_dropped() {
        let md = "---\n".repeat(1_000);
        let ir = with_blocks(&md, 10);
        // Each break is a block and a section: five of each fit.
        assert!(ir.sections.len() <= 7, "{} sections", ir.sections.len());
        assert!(ir.metadata.text_truncated);
    }

    #[test]
    fn test_table_cells_are_charged_before_the_row_is_built() {
        let wide = format!("|{}\n", "a|".repeat(50_000));
        let ir = with_blocks(&wide, 100);
        let cells: usize = ir
            .sections
            .iter()
            .flat_map(|s| &s.elements)
            .map(|e| match e {
                Element::Table(t) => t.rows.iter().map(|r| r.cells.len()).sum(),
                _ => 0,
            })
            .sum();
        assert_eq!(cells, 0, "a header row of 50,000 cells does not fit 100 blocks");
        assert!(ir.metadata.text_truncated);
        // A small table fits: 1 + 2 rows of 2 cells at 4 each = 17.
        let ir = with_blocks("|a|b|\n|-|-|\n|c|d|\n", 100);
        assert!(!ir.metadata.text_truncated);
    }

    #[test]
    fn test_a_document_within_the_block_limit_is_unchanged() {
        let md = "# T\n\ntext **b** [l](u)\n\n- a\n- b\n\n|x|y|\n|-|-|\n|1|2|\n\n```\ncode\n```\n\n---\n\nmore\n";
        let unlimited = with_blocks(md, usize::MAX);
        // Counted blocks: 2 sections, heading-less paragraphs, list, table
        // (1 + 4 cells x 2 rows), code, break.
        assert_eq!(with_blocks(md, 30), unlimited);
        assert!(!unlimited.metadata.text_truncated);
        assert!(with_blocks(md, 3).metadata.text_truncated);
    }
}
