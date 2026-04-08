use crate::core::units::Twip;

use super::document::BlockElement;

/// Section properties (`w:sectPr`).
#[derive(Debug, Clone, Default)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub margins: Option<PageMargins>,
    pub header_refs: Vec<HeaderFooterRef>,
    pub footer_refs: Vec<HeaderFooterRef>,
    pub columns: Option<u32>,
}

/// Page dimensions.
#[derive(Debug, Clone)]
pub struct PageSize {
    pub width: Twip,
    pub height: Twip,
    pub orient: Option<PageOrientation>,
}

/// Page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

/// Page margins.
#[derive(Debug, Clone)]
pub struct PageMargins {
    pub top: Twip,
    pub bottom: Twip,
    pub left: Twip,
    pub right: Twip,
    pub header: Option<Twip>,
    pub footer: Option<Twip>,
    pub gutter: Option<Twip>,
}

/// Reference to a header or footer part.
#[derive(Debug, Clone)]
pub struct HeaderFooterRef {
    pub hf_type: HeaderFooterType,
    pub relationship_id: String,
}

/// Type of header/footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    Default,
    First,
    Even,
}

/// A parsed header or footer.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    pub hf_type: HeaderFooterType,
    pub content: Vec<BlockElement>,
}
