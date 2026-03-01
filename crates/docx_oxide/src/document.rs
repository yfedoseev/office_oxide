use crate::paragraph::Paragraph;
use crate::table::Table;

/// The document body, containing all block-level elements.
#[derive(Debug, Clone, Default)]
pub struct Body {
    pub elements: Vec<BlockElement>,
}

/// A block-level element in the document body (or in a table cell).
#[derive(Debug, Clone)]
pub enum BlockElement {
    Paragraph(Paragraph),
    Table(Table),
}
