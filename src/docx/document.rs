use super::paragraph::Paragraph;
use super::table::Table;

/// The document body, containing all block-level elements.
#[derive(Debug, Clone, Default)]
pub struct Body {
    /// Ordered list of block elements (paragraphs and tables).
    pub elements: Vec<BlockElement>,
    /// Indices into `elements` where each `<w:sectPr>` boundary falls.
    /// `section_breaks[i]` is the **count of elements covered by the
    /// i-th section** — i.e. elements `[prev_break, section_breaks[i])`
    /// belong to section `i`. The final section runs from the last
    /// break to `elements.len()` and uses the document-level `sectPr`.
    /// Empty for documents with only one section.
    pub section_breaks: Vec<usize>,
}

/// A block-level element in the document body (or in a table cell).
// `Table` is boxed: it is ~2.7× the size of `Paragraph`, and an enum is as
// large as its largest variant, so an unboxed `Table` made every
// paragraph slot — and the capacity-4 `Vec` behind every one-paragraph
// table cell — pay for a table that was not there. A table is rare next
// to a paragraph, so the one allocation per table is the cheap side.
#[derive(Debug, Clone)]
pub enum BlockElement {
    /// A paragraph (`w:p`).
    Paragraph(Paragraph),
    /// A table (`w:tbl`).
    Table(Box<Table>),
}
