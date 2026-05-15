use crate::ir::*;

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
    /// `ThematicBreak` → `"---"` (matches markdown); container
    /// elements recursively render their children.
    pub fn default_plain(element: &Element) -> String {
        match element {
            Element::ThematicBreak => "---".to_string(),
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
            // Invisible in flow: shapes are positioned, not flow content;
            // page/column breaks have no plain-text counterpart; an
            // unannotated image shows nothing in plain text.
            Element::PageBreak | Element::ColumnBreak | Element::Shape(_) | Element::Image(_) => {
                String::new()
            },
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
                let alt = img.alt_text.as_deref().unwrap_or("");
                format!("![{alt}]()")
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
            Element::TextBox(tb) => tb
                .content
                .iter()
                .map(super::render_element_html)
                .collect::<Vec<_>>()
                .join("\n"),
            Element::Footnote(n) | Element::Endnote(n) => n
                .content
                .iter()
                .map(super::render_element_html)
                .collect::<Vec<_>>()
                .join("\n"),
            Element::PageBreak | Element::ColumnBreak | Element::Shape(_) => String::new(),
            Element::Image(img) => {
                let alt = img
                    .alt_text
                    .as_deref()
                    .map(super::escape_html)
                    .unwrap_or_default();
                let mut out = String::with_capacity(20 + alt.len());
                let _ = write!(out, "<img alt=\"{alt}\" />");
                out
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
            section_texts.join("\n\n---\n\n")
        }
    }

    /// Render the IR as an HTML fragment (no `<html>`/`<body>` wrapper).
    pub fn to_html(&self) -> String {
        let section_texts: Vec<String> = self
            .sections
            .iter()
            .map(render_section_html)
            .filter(|s| !s.is_empty())
            .collect();
        section_texts.join("\n<hr />\n")
    }

    /// Render the IR as markdown.
    pub fn to_markdown(&self) -> String {
        let section_texts: Vec<String> = self
            .sections
            .iter()
            .map(render_section_markdown)
            .filter(|s| !s.is_empty())
            .collect();
        section_texts.join("\n\n---\n\n")
    }
}

// ---------------------------------------------------------------------------
// Plain text rendering
// ---------------------------------------------------------------------------

fn render_section_plain(section: &Section) -> String {
    let mut parts = Vec::new();
    if let Some(ref title) = section.title {
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
    let mut rows = Vec::new();
    for row in &table.rows {
        let cells: Vec<String> = row
            .cells
            .iter()
            .map(|cell| {
                cell.content
                    .iter()
                    .map(render_element_plain)
                    .collect::<Vec<_>>()
                    .join(" ")
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
    if let Some(ref title) = section.title {
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
    parts.join("\n\n")
}

fn render_element_markdown(element: &Element) -> String {
    match element {
        Element::Heading(h) => {
            let hashes = "#".repeat(h.level.min(6) as usize);
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

fn render_inline_markdown(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => {
                let mut text = span.text.clone();

                if span.strikethrough {
                    text = format!("~~{text}~~");
                }
                if span.bold && span.italic {
                    text = format!("***{text}***");
                } else if span.bold {
                    text = format!("**{text}**");
                } else if span.italic {
                    text = format!("*{text}*");
                }

                if let Some(ref url) = span.hyperlink {
                    text = format!("[{text}]({url})");
                }

                out.push_str(&text);
            },
            InlineContent::LineBreak => out.push_str("  \n"),
            InlineContent::FootnoteRef(_) | InlineContent::EndnoteRef(_) => {},
        }
    }
    out
}

fn render_table_markdown(table: &Table) -> String {
    if table.rows.is_empty() {
        return String::new();
    }

    let col_count = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
    if col_count == 0 {
        return String::new();
    }

    let mut result = String::new();

    let first_row = &table.rows[0];
    result.push('|');
    for i in 0..col_count {
        let text = first_row
            .cells
            .get(i)
            .map(render_cell_markdown)
            .unwrap_or_default();
        result.push(' ');
        result.push_str(&text);
        result.push_str(" |");
    }
    result.push('\n');

    // Separator
    result.push('|');
    for _ in 0..col_count {
        result.push_str(" --- |");
    }
    result.push('\n');

    // Remaining rows
    for row in table.rows.iter().skip(1) {
        result.push('|');
        for i in 0..col_count {
            let text = row
                .cells
                .get(i)
                .map(render_cell_markdown)
                .unwrap_or_default();
            result.push(' ');
            result.push_str(&text);
            result.push_str(" |");
        }
        result.push('\n');
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
    let mut lines = Vec::new();
    for (i, item) in list.items.iter().enumerate() {
        let text = item
            .content
            .iter()
            .map(render_element_markdown)
            .collect::<Vec<_>>()
            .join(" ");
        let marker = if list.ordered {
            format!("{}. ", i + 1)
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

fn escape_html(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn render_section_html(section: &Section) -> String {
    let mut parts = Vec::new();
    if let Some(ref title) = section.title {
        if !title.is_empty() {
            parts.push(format!("<h2>{}</h2>", escape_html(title)));
        }
    }
    for elem in &section.elements {
        let html = render_element_html(elem);
        if !html.is_empty() {
            parts.push(html);
        }
    }
    parts.join("\n")
}

fn render_element_html(element: &Element) -> String {
    match element {
        Element::Heading(h) => {
            let level = h.level.clamp(1, 6);
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

fn render_inline_html(content: &[InlineContent]) -> String {
    let mut out = String::new();
    for item in content {
        match item {
            InlineContent::Text(span) => {
                let mut text = escape_html(&span.text);

                if span.bold {
                    text = format!("<strong>{text}</strong>");
                }
                if span.italic {
                    text = format!("<em>{text}</em>");
                }
                if span.strikethrough {
                    text = format!("<del>{text}</del>");
                }
                if let Some(ref url) = span.hyperlink {
                    text = format!("<a href=\"{}\">{text}</a>", escape_html(url));
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
            let content: Vec<String> = cell.content.iter().map(render_element_html).collect();
            html.push_str(&format!("<{tag}{attrs}>{}</{tag}>", content.join("")));
        }
        html.push_str("</tr>\n");
    }

    html.push_str("</table>");
    html
}

fn render_list_html(list: &List) -> String {
    let tag = if list.ordered { "ol" } else { "ul" };
    let mut html = format!("<{tag}>\n");
    for item in &list.items {
        let content = item
            .content
            .iter()
            .map(render_element_html)
            .collect::<Vec<_>>()
            .join("");
        html.push_str(&format!("<li>{content}"));
        if let Some(ref nested) = item.nested {
            html.push('\n');
            html.push_str(&render_list_html(nested));
        }
        html.push_str("</li>\n");
    }
    html.push_str(&format!("</{tag}>"));
    html
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

    #[test]
    fn plain_text_paragraph() {
        let ir = simple_ir(vec![para("Hello world")]);
        assert_eq!(ir.plain_text(), "Hello world");
    }

    #[test]
    fn markdown_heading() {
        let ir = simple_ir(vec![Element::Heading(Heading {
            level: 2,
            content: vec![span("Title")],
            ..Default::default()
        })]);
        assert_eq!(ir.to_markdown(), "## Title");
    }

    #[test]
    fn markdown_formatting() {
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
    fn markdown_table() {
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
    fn markdown_list() {
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
    fn markdown_hyperlink() {
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
    fn multi_section_separator() {
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
        };
        let plain = ir.plain_text();
        assert!(plain.contains("Sheet1"));
        assert!(plain.contains("Data A"));
        assert!(plain.contains("---"));
        assert!(plain.contains("Data B"));
    }

    #[test]
    fn html_paragraph() {
        let ir = simple_ir(vec![para("Hello world")]);
        assert_eq!(ir.to_html(), "<p>Hello world</p>");
    }

    #[test]
    fn html_formatting() {
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
    fn html_escaping() {
        let ir = simple_ir(vec![para("<script>alert('xss')</script>")]);
        assert!(ir.to_html().contains("&lt;script&gt;"));
        assert!(!ir.to_html().contains("<script>"));
    }

    #[test]
    fn html_table() {
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
    fn html_list() {
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

    // ── Defaults centralized in `block_default` ──────────────────────

    #[test]
    fn thematic_break_renders_as_hr_in_plain() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        assert_eq!(ir.plain_text(), "---");
    }

    #[test]
    fn thematic_break_renders_in_markdown() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        assert!(ir.to_markdown().contains("---"));
    }

    #[test]
    fn page_break_invisible_in_plain() {
        // PageBreak/ColumnBreak/Shape/Image have no plain-text counterpart
        // — they collapse to empty so plain_text shows only the surrounding
        // content.
        let ir = simple_ir(vec![para("before"), Element::PageBreak, para("after")]);
        let plain = ir.plain_text();
        assert!(plain.contains("before"));
        assert!(plain.contains("after"));
    }

    #[test]
    fn shape_invisible_in_plain() {
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
    fn text_box_recursively_renders_children() {
        let ir = simple_ir(vec![Element::TextBox(TextBox {
            content: vec![para("inside")],
            ..Default::default()
        })]);
        let plain = ir.plain_text();
        assert!(plain.contains("inside"), "plain: {plain}");
    }

    #[test]
    fn html_thematic_break() {
        let ir = simple_ir(vec![Element::ThematicBreak]);
        let html = ir.to_html();
        assert!(html.contains("<hr"), "html: {html}");
    }
}
