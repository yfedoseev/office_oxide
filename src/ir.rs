use crate::format::DocumentFormat;

/// A format-agnostic intermediate representation of a document.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct DocumentIR {
    /// Document-level metadata (format, title, etc.).
    pub metadata: Metadata,
    /// Ordered list of sections (pages, worksheets, slides, etc.).
    pub sections: Vec<Section>,
}

/// Document-level metadata extracted from the source file.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Metadata {
    /// The source format this document was parsed from.
    pub format: DocumentFormat,
    /// Optional document title from core properties.
    pub title: Option<String>,
}

/// A logical section (DOCX: section break, XLSX: worksheet, PPTX: slide).
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Section {
    /// Optional section title (e.g. slide title or worksheet name).
    pub title: Option<String>,
    /// Content elements within this section.
    pub elements: Vec<Element>,
}

/// A block-level content element.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Element {
    /// A heading with a numeric level (1–6).
    Heading(Heading),
    /// A paragraph of inline content.
    Paragraph(Paragraph),
    /// A table.
    Table(Table),
    /// An ordered or unordered list.
    List(List),
    /// An embedded image.
    Image(Image),
    /// A horizontal rule / thematic break.
    ThematicBreak,
}

/// A heading element with a nesting level.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Heading {
    /// Heading level 1–6 (1 = largest).
    pub level: u8,
    /// Inline content of the heading.
    pub content: Vec<InlineContent>,
}

/// A paragraph of inline content.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Paragraph {
    /// Inline runs making up this paragraph.
    pub content: Vec<InlineContent>,
}

/// Inline content within a paragraph or heading.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum InlineContent {
    /// A styled text span.
    Text(TextSpan),
    /// A line break within a paragraph.
    LineBreak,
}

/// A styled run of text.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct TextSpan {
    /// The text content.
    pub text: String,
    /// Whether the text is bold.
    pub bold: bool,
    /// Whether the text is italic.
    pub italic: bool,
    /// Whether the text has strikethrough.
    pub strikethrough: bool,
    /// Optional hyperlink URL.
    pub hyperlink: Option<String>,
}

impl TextSpan {
    /// Create a plain (unformatted) text span.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            bold: false,
            italic: false,
            strikethrough: false,
            hyperlink: None,
        }
    }
}

/// A table with rows and cells.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Table {
    /// Rows in the table (first row is header when `is_header = true`).
    pub rows: Vec<TableRow>,
}

/// A single row within a table.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct TableRow {
    /// Cells within this row.
    pub cells: Vec<TableCell>,
    /// Whether this row is a header row.
    pub is_header: bool,
}

/// A single cell within a table row.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct TableCell {
    /// Block elements inside the cell.
    pub content: Vec<Element>,
    /// Number of columns this cell spans.
    pub col_span: u32,
    /// Number of rows this cell spans.
    pub row_span: u32,
}

/// An ordered or unordered list.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct List {
    /// `true` = numbered list, `false` = bullet list.
    pub ordered: bool,
    /// Items in the list.
    pub items: Vec<ListItem>,
}

/// A single item within a list.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct ListItem {
    /// Inline content of this item.
    pub content: Vec<InlineContent>,
    /// Optional nested sub-list.
    pub nested: Option<List>,
}

/// An embedded image reference.
#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct Image {
    /// Optional alt-text description of the image.
    pub alt_text: Option<String>,
}
