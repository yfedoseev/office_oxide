use crate::core::units::Emu;

/// Information about a drawing/image reference within a run.
///
/// Carries enough data for both bitmap pictures (`<a:blip r:embed=…/>`)
/// and DrawingML preset shapes (`<wps:wsp>` with `<a:prstGeom>`). Only
/// one of `relationship_id` or `shape` is populated for any given
/// drawing — the consumer (`convert_docx`) uses whichever is set to
/// decide what kind of IR `Element` to emit.
#[derive(Debug, Clone)]
pub struct DrawingInfo {
    /// Relationship ID pointing to the image part. Empty when the
    /// drawing is a vector shape rather than a raster picture.
    pub relationship_id: String,
    /// Alt-text description from `wp:docPr/@descr`.
    pub description: Option<String>,
    /// Image / shape width in EMUs.
    pub width: Emu,
    /// Image / shape height in EMUs.
    pub height: Emu,
    /// `true` = inline, `false` = anchor (floating).
    pub inline: bool,
    /// Floating-anchor position (only set when `inline == false`).
    pub anchor_position: Option<AnchorPosition>,
    /// Vector shape data when the drawing is a `<wps:wsp>` rather
    /// than an embedded picture.
    pub shape: Option<ShapeInfo>,
}

/// Absolute coordinates extracted from a `<wp:anchor>` wrapper.
#[derive(Debug, Clone, Copy, Default)]
pub struct AnchorPosition {
    /// Horizontal offset in EMUs.
    pub x_emu: i64,
    /// Vertical offset in EMUs.
    pub y_emu: i64,
    /// What the horizontal offset is anchored to (page / margin / column).
    pub h_relative_from: AnchorFrame,
    /// What the vertical offset is anchored to (page / margin / paragraph).
    pub v_relative_from: AnchorFrame,
}

/// Reference frame for a floating-object anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AnchorFrame {
    /// Position relative to the page.
    #[default]
    Page,
    /// Position relative to the page margin.
    Margin,
    /// Position relative to the column.
    Column,
    /// Position relative to the paragraph (for vertical anchor).
    Paragraph,
    /// Position relative to the page line (for vertical anchor).
    Line,
    /// Position relative to the character (for horizontal anchor).
    Character,
}

/// Vector-shape data parsed from `<wps:wsp>`.
#[derive(Debug, Clone)]
pub struct ShapeInfo {
    /// Geometry preset from `<a:prstGeom prst="…">`.
    pub kind: ShapeKind,
    /// Stroke colour (`<a:ln><a:solidFill><a:srgbClr val="…"/>`).
    pub stroke_rgb: Option<(u8, u8, u8)>,
    /// Fill colour (`<wps:spPr><a:solidFill><a:srgbClr val="…"/>`).
    pub fill_rgb: Option<(u8, u8, u8)>,
    /// Stroke width in EMUs (`<a:ln w="…">`).
    pub stroke_w_emu: Option<i64>,
}

/// Subset of DrawingML preset shape kinds we currently round-trip.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeKind {
    /// Straight line (`prst="line"`).
    Line,
    /// Rectangle (`prst="rect"`).
    Rect,
}
