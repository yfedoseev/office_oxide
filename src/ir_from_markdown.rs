//! Parse Markdown text into a `DocumentIR`.
//!
//! Handles the subset of Markdown that document extraction pipelines and
//! office_oxide's own `to_markdown()` produce: ATX headings, pipe tables,
//! bullet/numbered lists, thematic breaks, and paragraphs with bold/italic
//! inline spans.  This is not a full CommonMark implementation — it is
//! intentionally minimal so it carries no extra dependencies.

use crate::format::DocumentFormat;
use crate::ir::{
    DocumentIR, Element, Heading, InlineContent, List, ListItem, Metadata, Paragraph, Section,
    Table, TableCell, TableRow, TextSpan,
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
    pub fn from_markdown(markdown: &str, format: DocumentFormat) -> Self {
        let mut parser = MarkdownParser::new(markdown);
        let sections = parser.parse_sections();
        DocumentIR {
            metadata: Metadata {
                format,
                title: None,
                ..Default::default()
            },
            sections,
        }
    }
}

// ---------------------------------------------------------------------------
// Parser state
// ---------------------------------------------------------------------------

struct MarkdownParser<'a> {
    lines: Vec<&'a str>,
    pos: usize,
}

impl<'a> MarkdownParser<'a> {
    fn new(src: &'a str) -> Self {
        Self {
            lines: src.lines().collect(),
            pos: 0,
        }
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
            let line = match self.peek() {
                Some(l) => l,
                None => break,
            };

            // Blank line
            if line.trim().is_empty() {
                self.advance();
                continue;
            }

            // Thematic break `---` / `***` / `___` starts a new section
            if is_thematic_break(line) {
                self.advance();
                if !current.elements.is_empty() || current.title.is_some() {
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
                self.advance();
                // H1 headings start a new section
                if level == 1 {
                    if !current.elements.is_empty() || current.title.is_some() {
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
            }

            // Unordered list
            if is_unordered_list_marker(line) {
                let list = self.parse_list(false);
                current.elements.push(Element::List(list));
                continue;
            }

            // Ordered list
            if is_ordered_list_marker(line) {
                let list = self.parse_list(true);
                current.elements.push(Element::List(list));
                continue;
            }

            // Regular paragraph (accumulate until blank line or block element)
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

        let mut rows: Vec<TableRow> = Vec::new();
        for (i, row_line) in data_rows.iter().enumerate() {
            let cells = split_pipe_row(row_line)
                .into_iter()
                .map(|cell_text| TableCell {
                    content: vec![Element::Paragraph(Paragraph {
                        content: parse_inline(cell_text.trim()),
                        ..Default::default()
                    })],
                    col_span: 1,
                    row_span: 1,
                    ..Default::default()
                })
                .collect();
            rows.push(TableRow {
                cells,
                is_header: i == 0,
                ..Default::default()
            });
        }

        Some(Table { rows, ..Default::default() })
    }

    fn parse_list(&mut self, ordered: bool) -> List {
        let mut items: Vec<ListItem> = Vec::new();
        loop {
            match self.peek() {
                None => break,
                Some(line) => {
                    if ordered && !is_ordered_list_marker(line) {
                        break;
                    }
                    if !ordered && !is_unordered_list_marker(line) {
                        break;
                    }
                    self.advance();
                    let content_str = strip_list_marker(line);
                    items.push(ListItem {
                        content: vec![Element::Paragraph(Paragraph {
                            content: parse_inline(content_str),
                            ..Default::default()
                        })],
                        nested: None,
                    });
                },
            }
        }
        List { ordered, items, ..Default::default() }
    }
}

// ---------------------------------------------------------------------------
// Inline parser (bold, italic, plain)
// ---------------------------------------------------------------------------

fn parse_inline(text: &str) -> Vec<InlineContent> {
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
        let text = rest.trim().trim_end_matches('#').trim().to_string();
        Some((hashes as u8, text))
    } else {
        None
    }
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

    #[test]
    fn parse_heading_paragraph() {
        let md = "# Hello\n\nSome text here.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Hello"));
        assert!(matches!(ir.sections[0].elements[0], Element::Paragraph(_)));
    }

    #[test]
    fn parse_page_break_into_sections() {
        let md = "# Page 1\n\nText one.\n\n---\n\n# Page 2\n\nText two.\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 2);
        assert_eq!(ir.sections[0].title.as_deref(), Some("Page 1"));
        assert_eq!(ir.sections[1].title.as_deref(), Some("Page 2"));
    }

    #[test]
    fn parse_unordered_list() {
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
    fn parse_ordered_list() {
        let md = "1. first\n2. second\n";
        let ir = DocumentIR::from_markdown(md, DocumentFormat::Docx);
        let list = match &ir.sections[0].elements[0] {
            Element::List(l) => l,
            other => panic!("expected List, got {other:?}"),
        };
        assert!(list.ordered);
    }

    #[test]
    fn parse_pipe_table() {
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
    fn parse_bold_italic_inline() {
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
    fn parse_empty_markdown() {
        let ir = DocumentIR::from_markdown("", DocumentFormat::Docx);
        assert_eq!(ir.sections.len(), 1);
        assert!(ir.sections[0].elements.is_empty());
    }
}
