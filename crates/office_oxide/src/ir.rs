use crate::format::DocumentFormat;

/// A format-agnostic intermediate representation of a document.
#[derive(Debug, Clone)]
pub struct DocumentIR {
    pub metadata: Metadata,
    pub sections: Vec<Section>,
}

#[derive(Debug, Clone)]
pub struct Metadata {
    pub format: DocumentFormat,
    pub title: Option<String>,
}

/// A logical section (DOCX: section break, XLSX: worksheet, PPTX: slide).
#[derive(Debug, Clone)]
pub struct Section {
    pub title: Option<String>,
    pub elements: Vec<Element>,
}

#[derive(Debug, Clone)]
pub enum Element {
    Heading(Heading),
    Paragraph(Paragraph),
    Table(Table),
    List(List),
    Image(Image),
    ThematicBreak,
}

#[derive(Debug, Clone)]
pub struct Heading {
    pub level: u8,
    pub content: Vec<InlineContent>,
}

#[derive(Debug, Clone)]
pub struct Paragraph {
    pub content: Vec<InlineContent>,
}

#[derive(Debug, Clone)]
pub enum InlineContent {
    Text(TextSpan),
    LineBreak,
}

#[derive(Debug, Clone)]
pub struct TextSpan {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub strikethrough: bool,
    pub hyperlink: Option<String>,
}

#[derive(Debug, Clone)]
pub struct Table {
    pub rows: Vec<TableRow>,
}

#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub is_header: bool,
}

#[derive(Debug, Clone)]
pub struct TableCell {
    pub content: Vec<Element>,
    pub col_span: u32,
    pub row_span: u32,
}

#[derive(Debug, Clone)]
pub struct List {
    pub ordered: bool,
    pub items: Vec<ListItem>,
}

#[derive(Debug, Clone)]
pub struct ListItem {
    pub content: Vec<InlineContent>,
    pub nested: Option<List>,
}

#[derive(Debug, Clone)]
pub struct Image {
    pub alt_text: Option<String>,
}
