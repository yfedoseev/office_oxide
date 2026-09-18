use super::PptxDocument;
use super::shape::{
    GraphicContent, HyperlinkTarget, Shape, ShapePosition, Table, TableCell, TableRow, TextBody,
    TextContent,
};

impl PptxDocument {
    /// Extract plain text from the entire presentation.
    ///
    /// Shapes are spatially sorted (top-to-bottom, left-to-right) per slide.
    /// Slides are separated by `\n\n---\n\n`.
    pub fn plain_text(&self) -> String {
        let mut parts = Vec::new();
        for (i, _) in self.slides.iter().enumerate() {
            if let Some(text) = self.slide_plain_text(i) {
                if !text.is_empty() {
                    parts.push(text);
                }
            }
        }
        parts.join("\n\n---\n\n")
    }

    /// Extract plain text from a single slide by index.
    pub fn slide_plain_text(&self, index: usize) -> Option<String> {
        let slide = self.slides.get(index)?;
        let mut entries = Vec::new();
        collect_text_entries(&slide.shapes, &mut entries);
        entries.sort_by(|a, b| spatial_cmp(&a.0, &b.0));

        let mut parts = Vec::new();
        for (_, text) in &entries {
            if !text.is_empty() {
                parts.push(text.as_str());
            }
        }

        let mut result = parts.join("\n\n");

        if let Some(ref notes) = slide.notes {
            if !notes.is_empty() {
                result.push_str("\n\n[Notes]\n");
                result.push_str(notes);
            }
        }

        Some(result)
    }

    /// Convert the entire presentation to markdown.
    pub fn to_markdown(&self) -> String {
        let mut parts = Vec::new();
        for (i, _) in self.slides.iter().enumerate() {
            if let Some(md) = self.slide_to_markdown(i) {
                parts.push(md);
            }
        }
        parts.join("\n\n")
    }

    /// Convert a single slide to markdown by index.
    pub fn slide_to_markdown(&self, index: usize) -> Option<String> {
        let slide = self.slides.get(index)?;
        let mut result = String::new();

        // Slide heading: use title placeholder text or "Slide N"
        let title = find_title_text(&slide.shapes);
        if let Some(ref title) = title {
            result.push_str(&format!("## {title}\n\n"));
        } else {
            result.push_str(&format!("## Slide {}\n\n", index + 1));
        }

        let mut entries = Vec::new();
        collect_markdown_entries(&slide.shapes, &mut entries);
        entries.sort_by(|a, b| spatial_cmp(&a.0, &b.0));

        for (_, md) in &entries {
            if !md.is_empty() {
                result.push_str(md);
                result.push_str("\n\n");
            }
        }

        if let Some(ref notes) = slide.notes {
            if !notes.is_empty() {
                for line in notes.lines() {
                    result.push_str("> ");
                    result.push_str(line);
                    result.push('\n');
                }
                result.push('\n');
            }
        }

        // Trim trailing whitespace
        let trimmed = result.trim_end().to_string();
        Some(trimmed)
    }
}

// ---------------------------------------------------------------------------
// Spatial sorting
// ---------------------------------------------------------------------------

fn spatial_cmp(a: &Option<ShapePosition>, b: &Option<ShapePosition>) -> std::cmp::Ordering {
    match (a, b) {
        (Some(a), Some(b)) => a.y.cmp(&b.y).then(a.x.cmp(&b.x)),
        (Some(_), None) => std::cmp::Ordering::Less,
        (None, Some(_)) => std::cmp::Ordering::Greater,
        (None, None) => std::cmp::Ordering::Equal,
    }
}

// ---------------------------------------------------------------------------
// Plain text collection
// ---------------------------------------------------------------------------

/// Collect (position, plain_text) entries from shapes, recursing into groups.
fn collect_text_entries(shapes: &[Shape], entries: &mut Vec<(Option<ShapePosition>, String)>) {
    for shape in shapes {
        match shape {
            Shape::AutoShape(auto) => {
                if let Some(ref tb) = auto.text_body {
                    let text = plain_text_from_body(tb);
                    if !text.is_empty() {
                        entries.push((auto.position.clone(), text));
                    }
                }
            },
            Shape::Picture(pic) => {
                if let Some(ref alt) = pic.alt_text {
                    if !alt.is_empty() {
                        entries.push((pic.position.clone(), alt.clone()));
                    }
                }
            },
            Shape::Group(grp) => {
                collect_text_entries(&grp.children, entries);
            },
            Shape::GraphicFrame(gf) => {
                if let GraphicContent::Table(ref tbl) = gf.content {
                    let text = plain_text_from_table(tbl);
                    if !text.is_empty() {
                        entries.push((gf.position.clone(), text));
                    }
                }
            },
            Shape::Connector(_) => {},
        }
    }
}

fn plain_text_from_body(body: &TextBody) -> String {
    let mut parts = Vec::new();
    for para in &body.paragraphs {
        let mut para_text = String::new();
        for content in &para.content {
            match content {
                TextContent::Run(run) => para_text.push_str(&run.text),
                TextContent::LineBreak => para_text.push('\n'),
                TextContent::Field(field) => para_text.push_str(&field.text),
            }
        }
        parts.push(para_text);
    }
    parts.join("\n")
}

/// Resolve `grid_span`/`row_span` into a proper 2-D grid, so a merged
/// cell's neighbours land at their true column position instead of
/// shifting left to fill the gap. Mirrors `ir_render.rs::table_grid`
/// exactly — this crate hit the identical "flat cell list, not a grid"
/// bug shape a third time here, on PPTX's own default text/markdown
/// renderers, after DOCX's read side (#137, closed) and the write path
/// (#265) (issue #289).
fn pptx_table_grid(table: &Table) -> Vec<Vec<Option<&TableCell>>> {
    // A raw PPTX row's `cells` includes h_merge/v_merge *continuation*
    // placeholders alongside the real, span-owning cell — unlike the IR's
    // `Table`, which never carries them at all (convert_pptx_table drops
    // them before building `TableRow`). Summing every cell's grid_span,
    // continuation placeholders included, double-counted the columns a
    // merge already covers.
    fn real_cells(r: &TableRow) -> impl Iterator<Item = &TableCell> {
        r.cells.iter().filter(|c| !c.h_merge && !c.v_merge)
    }
    let cell_total: usize = table.rows.iter().map(|r| real_cells(r).count()).sum();
    let width = table
        .rows
        .iter()
        .map(|r| real_cells(r).map(|c| c.grid_span.max(1) as usize).sum::<usize>())
        .max()
        .unwrap_or(0)
        .min(cell_total.saturating_mul(1_000).max(1));
    let mut grid: Vec<Vec<Option<&TableCell>>> = vec![vec![None; width]; table.rows.len()];
    let mut covered: Vec<Vec<bool>> = vec![vec![false; width]; table.rows.len()];

    for (r, row) in table.rows.iter().enumerate() {
        let mut c = 0usize;
        for cell in &row.cells {
            if cell.h_merge || cell.v_merge {
                continue;
            }
            while c < width && covered[r][c] {
                c += 1;
            }
            if c >= width {
                break;
            }
            grid[r][c] = Some(cell);
            let cs = cell.grid_span.max(1) as usize;
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

fn cell_plain_text(cell: &TableCell) -> String {
    cell.text_body.as_ref().map(plain_text_from_body).unwrap_or_default()
}

fn plain_text_from_table(table: &Table) -> String {
    let grid = pptx_table_grid(table);
    let rows: Vec<String> = grid
        .iter()
        .map(|row| {
            row.iter()
                .map(|slot| slot.map(cell_plain_text).unwrap_or_default())
                .collect::<Vec<_>>()
                .join("\t")
        })
        .collect();
    rows.join("\n")
}

// ---------------------------------------------------------------------------
// Markdown collection
// ---------------------------------------------------------------------------

fn collect_markdown_entries(shapes: &[Shape], entries: &mut Vec<(Option<ShapePosition>, String)>) {
    for shape in shapes {
        match shape {
            Shape::AutoShape(auto) => {
                // Skip title placeholder — already used as heading
                if auto
                    .placeholder
                    .as_ref()
                    .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref()))
                {
                    continue;
                }
                if let Some(ref tb) = auto.text_body {
                    let md = markdown_from_body(tb);
                    if !md.is_empty() {
                        entries.push((auto.position.clone(), md));
                    }
                }
            },
            Shape::Picture(pic) => {
                if let Some(ref alt) = pic.alt_text {
                    if !alt.is_empty() {
                        entries.push((pic.position.clone(), format!("![{alt}]()")));
                    }
                }
            },
            Shape::Group(grp) => {
                collect_markdown_entries(&grp.children, entries);
            },
            Shape::GraphicFrame(gf) => {
                if let GraphicContent::Table(ref tbl) = gf.content {
                    let md = markdown_table(tbl);
                    if !md.is_empty() {
                        entries.push((gf.position.clone(), md));
                    }
                }
            },
            Shape::Connector(_) => {},
        }
    }
}

fn markdown_from_body(body: &TextBody) -> String {
    let mut parts = Vec::new();
    for para in &body.paragraphs {
        let text = markdown_paragraph(para);
        parts.push(text);
    }
    parts.join("\n")
}

fn markdown_paragraph(para: &super::shape::TextParagraph) -> String {
    let mut text = String::new();
    for content in &para.content {
        match content {
            TextContent::Run(run) => {
                text.push_str(&markdown_run(run));
            },
            TextContent::LineBreak => {
                text.push_str("  \n");
            },
            TextContent::Field(field) => {
                text.push_str(&field.text);
            },
        }
    }
    // Add bullet indent for outline levels > 0
    if para.level > 0 {
        let indent = "  ".repeat(para.level as usize);
        format!("{indent}- {text}")
    } else {
        text
    }
}

fn markdown_run(run: &super::shape::TextRun) -> String {
    if run.text.is_empty() {
        return String::new();
    }

    let mut text = run.text.clone();

    // Apply inline formatting
    if run.strikethrough {
        text = format!("~~{text}~~");
    }
    if run.bold == Some(true) && run.italic == Some(true) {
        text = format!("***{text}***");
    } else if run.bold == Some(true) {
        text = format!("**{text}**");
    } else if run.italic == Some(true) {
        text = format!("*{text}*");
    }

    // Apply hyperlink
    if let Some(ref link) = run.hyperlink {
        match &link.target {
            HyperlinkTarget::External(url) => {
                text = format!("[{text}]({url})");
            },
            HyperlinkTarget::Internal(_) => {
                // Internal links — just keep the text
            },
        }
    }

    text
}

fn markdown_table(table: &Table) -> String {
    if table.rows.is_empty() {
        return String::new();
    }

    let grid = pptx_table_grid(table);
    let col_count = grid.first().map(Vec::len).unwrap_or(0);
    if col_count == 0 {
        return String::new();
    }
    let md_rows: Vec<Vec<String>> = grid
        .iter()
        .map(|row| {
            row.iter()
                .map(|slot| {
                    slot.map(|cell| {
                        cell.text_body
                            .as_ref()
                            .map(|tb| {
                                // Flatten paragraphs for table cells — replace newlines with spaces
                                plain_text_from_body(tb).replace('\n', " ")
                            })
                            .unwrap_or_default()
                    })
                    .unwrap_or_default()
                })
                .collect()
        })
        .collect();

    let mut result = String::new();

    // Header row
    if let Some(header) = md_rows.first() {
        result.push('|');
        for cell in header {
            result.push_str(&format!(" {cell} |"));
        }
        result.push('\n');

        // Separator
        result.push('|');
        for _ in 0..col_count {
            result.push_str(" --- |");
        }
        result.push('\n');
    }

    // Data rows
    for row in md_rows.iter().skip(1) {
        result.push('|');
        for cell in row {
            result.push_str(&format!(" {cell} |"));
        }
        result.push('\n');
    }

    // Remove trailing newline
    if result.ends_with('\n') {
        result.pop();
    }

    result
}

// ---------------------------------------------------------------------------
// Title extraction
// ---------------------------------------------------------------------------

fn is_title_placeholder(ph_type: Option<&str>) -> bool {
    matches!(ph_type, Some("title" | "ctrTitle"))
}

fn find_title_text(shapes: &[Shape]) -> Option<String> {
    for shape in shapes {
        match shape {
            Shape::AutoShape(auto)
                if auto
                    .placeholder
                    .as_ref()
                    .is_some_and(|ph| is_title_placeholder(ph.ph_type.as_deref())) =>
            {
                if let Some(ref tb) = auto.text_body {
                    let text = plain_text_from_body(tb);
                    if !text.is_empty() {
                        return Some(text);
                    }
                }
            },
            Shape::Group(grp) => {
                if let Some(title) = find_title_text(&grp.children) {
                    return Some(title);
                }
            },
            _ => {},
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use crate::pptx::PptxDocument;
    use crate::pptx::presentation::PresentationInfo;
    use crate::pptx::shape::*;
    use crate::pptx::slide::Slide;

    fn make_doc(slides: Vec<Slide>) -> PptxDocument {
        PptxDocument {
            core_properties: None,
            app_properties: None,
            has_macros: false,
            presentation: PresentationInfo {
                slides: Vec::new(),
                slide_size: None,
            },
            slides,
            theme: None,
            embedded_fonts: Vec::new(),
        }
    }

    fn text_shape(name: &str, text: &str, x: i64, y: i64) -> Shape {
        Shape::AutoShape(AutoShape {
            id: 1,
            name: name.to_string(),
            alt_text: None,
            position: Some(ShapePosition {
                x,
                y,
                cx: 1000,
                cy: 500,
            }),
            text_body: Some(TextBody {
                paragraphs: vec![TextParagraph {
                    level: 0,
                    alignment: None,
                    space_before_hundredths_pt: None,
                    content: vec![TextContent::Run(TextRun {
                        text: text.to_string(),
                        bold: None,
                        italic: None,
                        strikethrough: false,
                        hyperlink: None,
                        font_size_hundredths_pt: None,
                        color_rgb: None,
                        ..Default::default()
                    })],
                    ..Default::default()
                }],
            }),
            placeholder: None,
            hyperlink: None,
        })
    }

    fn title_shape(text: &str) -> Shape {
        Shape::AutoShape(AutoShape {
            id: 1,
            name: "Title".to_string(),
            alt_text: None,
            position: Some(ShapePosition {
                x: 0,
                y: 0,
                cx: 9000,
                cy: 1000,
            }),
            text_body: Some(TextBody {
                paragraphs: vec![TextParagraph {
                    level: 0,
                    alignment: None,
                    space_before_hundredths_pt: None,
                    content: vec![TextContent::Run(TextRun {
                        text: text.to_string(),
                        bold: None,
                        italic: None,
                        strikethrough: false,
                        hyperlink: None,
                        font_size_hundredths_pt: None,
                        color_rgb: None,
                        ..Default::default()
                    })],
                    ..Default::default()
                }],
            }),
            placeholder: Some(PlaceholderInfo {
                ph_type: Some("title".to_string()),
                idx: Some(0),
            }),
            hyperlink: None,
        })
    }

    #[test]
    fn spatial_sort_order() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![
                text_shape("Bottom", "bottom text", 100, 5000),
                text_shape("Top", "top text", 200, 100),
                text_shape("Middle", "middle text", 50, 2500),
            ],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let text = doc.slide_plain_text(0).unwrap();
        assert_eq!(text, "top text\n\nmiddle text\n\nbottom text");
    }

    #[test]
    fn plain_text_with_notes() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![text_shape("Text", "Hello", 0, 0)],
            notes: Some("Speaker notes".to_string()),
            background_rgb: None,
            ..Default::default()
        }]);

        let text = doc.slide_plain_text(0).unwrap();
        assert_eq!(text, "Hello\n\n[Notes]\nSpeaker notes");
    }

    #[test]
    fn plain_text_multi_slide() {
        let doc = make_doc(vec![
            Slide {
                name: String::new(),
                shapes: vec![text_shape("A", "Slide one", 0, 0)],
                notes: None,
                background_rgb: None,
                ..Default::default()
            },
            Slide {
                name: String::new(),
                shapes: vec![text_shape("B", "Slide two", 0, 0)],
                notes: None,
                background_rgb: None,
                ..Default::default()
            },
        ]);

        let text = doc.plain_text();
        assert_eq!(text, "Slide one\n\n---\n\nSlide two");
    }

    #[test]
    fn markdown_with_title() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![
                title_shape("My Title"),
                text_shape("Body", "Body text", 0, 2000),
            ],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.starts_with("## My Title\n\n"));
        assert!(md.contains("Body text"));
        // Title should not be duplicated in body
        assert_eq!(md.matches("My Title").count(), 1);
    }

    #[test]
    fn markdown_formatting() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![Shape::AutoShape(AutoShape {
                id: 1,
                name: "Text".to_string(),
                alt_text: None,
                position: Some(ShapePosition {
                    x: 0,
                    y: 0,
                    cx: 1000,
                    cy: 500,
                }),
                text_body: Some(TextBody {
                    paragraphs: vec![TextParagraph {
                        level: 0,
                        alignment: None,
                        space_before_hundredths_pt: None,
                        content: vec![
                            TextContent::Run(TextRun {
                                text: "bold".to_string(),
                                bold: Some(true),
                                italic: None,
                                strikethrough: false,
                                hyperlink: None,
                                font_size_hundredths_pt: None,
                                color_rgb: None,
                                ..Default::default()
                            }),
                            TextContent::Run(TextRun {
                                text: " and ".to_string(),
                                bold: None,
                                italic: None,
                                strikethrough: false,
                                hyperlink: None,
                                font_size_hundredths_pt: None,
                                color_rgb: None,
                                ..Default::default()
                            }),
                            TextContent::Run(TextRun {
                                text: "italic".to_string(),
                                bold: None,
                                italic: Some(true),
                                strikethrough: false,
                                hyperlink: None,
                                font_size_hundredths_pt: None,
                                color_rgb: None,
                                ..Default::default()
                            }),
                        ],
                        ..Default::default()
                    }],
                }),
                placeholder: None,
                hyperlink: None,
            })],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.contains("**bold** and *italic*"));
    }

    #[test]
    fn markdown_notes_blockquote() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![text_shape("Text", "Content", 0, 0)],
            notes: Some("Note line 1\nNote line 2".to_string()),
            background_rgb: None,
            ..Default::default()
        }]);

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.contains("> Note line 1\n> Note line 2"));
    }

    #[test]
    fn markdown_table() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![Shape::GraphicFrame(GraphicFrame {
                id: 1,
                name: "Table".to_string(),
                position: Some(ShapePosition {
                    x: 0,
                    y: 0,
                    cx: 9000,
                    cy: 3000,
                }),
                content: GraphicContent::Table(Table {
                    first_row_header: true,
                    last_row_header: false,
                    rows: vec![
                        TableRow {
                            cells: vec![
                                TableCell {
                                    text_body: Some(TextBody {
                                        paragraphs: vec![TextParagraph {
                                            level: 0,
                                            alignment: None,
                                            space_before_hundredths_pt: None,
                                            content: vec![TextContent::Run(TextRun {
                                                text: "H1".to_string(),
                                                bold: None,
                                                italic: None,
                                                strikethrough: false,
                                                hyperlink: None,
                                                font_size_hundredths_pt: None,
                                                color_rgb: None,
                                                ..Default::default()
                                            })],
                                            ..Default::default()
                                        }],
                                    }),
                                    grid_span: 1,
                                    row_span: 1,
                                    h_merge: false,
                                    v_merge: false,
                                },
                                TableCell {
                                    text_body: Some(TextBody {
                                        paragraphs: vec![TextParagraph {
                                            level: 0,
                                            alignment: None,
                                            space_before_hundredths_pt: None,
                                            content: vec![TextContent::Run(TextRun {
                                                text: "H2".to_string(),
                                                bold: None,
                                                italic: None,
                                                strikethrough: false,
                                                hyperlink: None,
                                                font_size_hundredths_pt: None,
                                                color_rgb: None,
                                                ..Default::default()
                                            })],
                                            ..Default::default()
                                        }],
                                    }),
                                    grid_span: 1,
                                    row_span: 1,
                                    h_merge: false,
                                    v_merge: false,
                                },
                            ],
                        },
                        TableRow {
                            cells: vec![
                                TableCell {
                                    text_body: Some(TextBody {
                                        paragraphs: vec![TextParagraph {
                                            level: 0,
                                            alignment: None,
                                            space_before_hundredths_pt: None,
                                            content: vec![TextContent::Run(TextRun {
                                                text: "A".to_string(),
                                                bold: None,
                                                italic: None,
                                                strikethrough: false,
                                                hyperlink: None,
                                                font_size_hundredths_pt: None,
                                                color_rgb: None,
                                                ..Default::default()
                                            })],
                                            ..Default::default()
                                        }],
                                    }),
                                    grid_span: 1,
                                    row_span: 1,
                                    h_merge: false,
                                    v_merge: false,
                                },
                                TableCell {
                                    text_body: Some(TextBody {
                                        paragraphs: vec![TextParagraph {
                                            level: 0,
                                            alignment: None,
                                            space_before_hundredths_pt: None,
                                            content: vec![TextContent::Run(TextRun {
                                                text: "B".to_string(),
                                                bold: None,
                                                italic: None,
                                                strikethrough: false,
                                                hyperlink: None,
                                                font_size_hundredths_pt: None,
                                                color_rgb: None,
                                                ..Default::default()
                                            })],
                                            ..Default::default()
                                        }],
                                    }),
                                    grid_span: 1,
                                    row_span: 1,
                                    h_merge: false,
                                    v_merge: false,
                                },
                            ],
                        },
                    ],
                }),
            })],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.contains("| H1 | H2 |"));
        assert!(md.contains("| --- | --- |"));
        assert!(md.contains("| A | B |"));
    }

    fn text_cell(text: &str, grid_span: u32, row_span: u32, h_merge: bool, v_merge: bool) -> TableCell {
        let text_body = if h_merge || v_merge {
            None
        } else {
            Some(TextBody {
                paragraphs: vec![TextParagraph {
                    level: 0,
                    alignment: None,
                    space_before_hundredths_pt: None,
                    content: vec![TextContent::Run(TextRun {
                        text: text.to_string(),
                        bold: None,
                        italic: None,
                        strikethrough: false,
                        hyperlink: None,
                        font_size_hundredths_pt: None,
                        color_rgb: None,
                        ..Default::default()
                    })],
                    ..Default::default()
                }],
            })
        };
        TableCell {
            text_body,
            grid_span,
            row_span,
            h_merge,
            v_merge,
        }
    }

    /// issue #289 — merged-cell tables rendered with the wrong column
    /// count and shifted/misaligned content: both plain_text_from_table
    /// and markdown_table filtered out h_merge/v_merge continuation
    /// cells and then just pushed the *remaining* cells into a flat
    /// list (`col_count = cells.len()`), with no grid/covered-position
    /// tracking — the same "flat cell list, not a grid" bug shape
    /// already found and fixed for DOCX's read side (#137) and write
    /// path (#265). A 2x2 grid where row 0's first cell spans both
    /// columns must still show row 1's two cells at their true
    /// positions, with row 0's merged cell's content followed by one
    /// blank slot, not shifted left.
    #[test]
    fn markdown_and_plain_text_tables_handle_merged_cells() {
        let table = Table {
            first_row_header: false,
            last_row_header: false,
            rows: vec![
                TableRow {
                    cells: vec![
                        text_cell("Merged", 2, 1, false, false),
                        text_cell("", 1, 1, true, false),
                    ],
                },
                TableRow {
                    cells: vec![
                        text_cell("A", 1, 1, false, false),
                        text_cell("B", 1, 1, false, false),
                    ],
                },
            ],
        };
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![Shape::GraphicFrame(super::super::shape::GraphicFrame {
                id: 1,
                name: "Table".to_string(),
                position: Some(ShapePosition {
                    x: 0,
                    y: 0,
                    cx: 9000,
                    cy: 3000,
                }),
                content: GraphicContent::Table(table),
            })],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let plain = doc.slide_plain_text(0).unwrap();
        assert!(
            plain.contains("A\tB"),
            "row 1's cells must stay at their true columns: {plain:?}"
        );

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.contains("| Merged |  |"), "row 0 must show 2 columns: {md:?}");
        assert!(md.contains("| A | B |"), "row 1 must not shift left: {md:?}");
    }

    #[test]
    fn markdown_hyperlink() {
        let doc = make_doc(vec![Slide {
            name: String::new(),
            shapes: vec![Shape::AutoShape(AutoShape {
                id: 1,
                name: "Text".to_string(),
                alt_text: None,
                position: Some(ShapePosition {
                    x: 0,
                    y: 0,
                    cx: 1000,
                    cy: 500,
                }),
                text_body: Some(TextBody {
                    paragraphs: vec![TextParagraph {
                        level: 0,
                        alignment: None,
                        space_before_hundredths_pt: None,
                        content: vec![TextContent::Run(TextRun {
                            text: "Click here".to_string(),
                            bold: None,
                            italic: None,
                            strikethrough: false,
                            hyperlink: Some(HyperlinkInfo {
                                target: HyperlinkTarget::External(
                                    "https://example.com".to_string(),
                                ),
                                tooltip: None,
                            }),
                            font_size_hundredths_pt: None,
                            color_rgb: None,
                            ..Default::default()
                        })],
                        ..Default::default()
                    }],
                }),
                placeholder: None,
                hyperlink: None,
            })],
            notes: None,
            background_rgb: None,
            ..Default::default()
        }]);

        let md = doc.slide_to_markdown(0).unwrap();
        assert!(md.contains("[Click here](https://example.com)"));
    }
}
