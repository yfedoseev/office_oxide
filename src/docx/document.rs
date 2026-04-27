use super::paragraph::Paragraph;
use super::table::Table;

/// The document body, containing all block-level elements.
#[derive(Debug, Clone, Default)]
pub struct Body {
    /// Ordered list of block elements (paragraphs and tables).
    pub elements: Vec<BlockElement>,
}

/// A block-level element in the document body (or in a table cell).
#[derive(Debug, Clone)]
pub enum BlockElement {
    /// A paragraph (`w:p`).
    Paragraph(Paragraph),
    /// A table (`w:tbl`).
    Table(Table),
}
