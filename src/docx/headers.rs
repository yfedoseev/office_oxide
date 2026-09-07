use crate::core::units::Twip;

use super::document::BlockElement;

/// Section properties (`w:sectPr`).
#[derive(Debug, Clone, Default)]
pub struct SectionProperties {
    /// Page size and orientation.
    pub page_size: Option<PageSize>,
    /// Page margin settings.
    pub margins: Option<PageMargins>,
    /// References to header parts.
    pub header_refs: Vec<HeaderFooterRef>,
    /// References to footer parts.
    pub footer_refs: Vec<HeaderFooterRef>,
    /// Number of text columns (if multi-column layout).
    pub columns: Option<u32>,
    /// Full `<w:cols>` layout: spacing, separator line and per-column widths.
    pub column_layout: Option<ColumnDefs>,
    /// Section break kind from `<w:type w:val="...">`. `None` means the
    /// element was absent, which ECMA-376 defines as `nextPage`.
    pub break_type: Option<SectionBreakKind>,
    /// `<w:titlePg/>` — this section uses a distinct first-page header/footer.
    pub title_page: bool,
}

/// `<w:cols>` layout details beyond the bare column count.
#[derive(Debug, Clone, Default)]
pub struct ColumnDefs {
    /// Default spacing between columns, in twips (`w:space`).
    pub space: Option<u32>,
    /// Whether a separator line is drawn between columns (`w:sep`).
    pub separator: bool,
    /// Per-column widths in twips from `<w:col w:w="..."/>`. Empty when the
    /// columns are equal-width (`w:equalWidth="1"` or no `<w:col>` children).
    pub widths: Vec<u32>,
}

/// Section break kind (`<w:type>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionBreakKind {
    /// No page break — the section continues on the same page.
    Continuous,
    /// Section starts on the next page.
    NextPage,
    /// Section starts on the next even-numbered page.
    EvenPage,
    /// Section starts on the next odd-numbered page.
    OddPage,
    /// Section starts in the next column.
    NextColumn,
}

/// Page dimensions.
#[derive(Debug, Clone)]
pub struct PageSize {
    /// Page width in twips.
    pub width: Twip,
    /// Page height in twips.
    pub height: Twip,
    /// Page orientation (portrait or landscape).
    pub orient: Option<PageOrientation>,
}

/// Page orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    /// Portrait orientation (height > width).
    Portrait,
    /// Landscape orientation (width > height).
    Landscape,
}

/// Page margins.
#[derive(Debug, Clone)]
pub struct PageMargins {
    /// Top margin in twips.
    pub top: Twip,
    /// Bottom margin in twips.
    pub bottom: Twip,
    /// Left margin in twips.
    pub left: Twip,
    /// Right margin in twips.
    pub right: Twip,
    /// Header distance from top edge in twips.
    pub header: Option<Twip>,
    /// Footer distance from bottom edge in twips.
    pub footer: Option<Twip>,
    /// Gutter width in twips.
    pub gutter: Option<Twip>,
}

/// Reference to a header or footer part.
#[derive(Debug, Clone)]
pub struct HeaderFooterRef {
    /// Which pages this header/footer applies to.
    pub hf_type: HeaderFooterType,
    /// Relationship ID referencing the header/footer part.
    pub relationship_id: String,
}

/// Type of header/footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    /// Default header/footer (odd pages or all pages).
    Default,
    /// First-page header/footer.
    First,
    /// Even-page header/footer.
    Even,
}

/// A parsed header or footer.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    /// Which pages this applies to.
    pub hf_type: HeaderFooterType,
    /// Block content within the header or footer.
    pub content: Vec<BlockElement>,
    /// `true` if this came from a `<w:headerReference>`, `false` if
    /// from a `<w:footerReference>`. Lets downstream consumers (e.g.
    /// the markdown renderer) sort entries into headers vs footers
    /// without trying to back-derive from cumulative ref counts.
    pub is_header: bool,
}
