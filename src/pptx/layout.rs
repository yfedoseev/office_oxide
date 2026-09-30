//! Slide layouts and slide masters.
//!
//! Text placed directly on a layout or master — outside any placeholder,
//! e.g. a company-name footer or boilerplate — is drawn on every slide
//! that uses it, but it is not part of any slide. Like python-pptx,
//! [`PptxDocument::plain_text`], [`PptxDocument::to_markdown`] and the IR
//! include only each slide's own shapes; the layout and master shapes are
//! available here, and [`PptxDocument::static_text_for_slide`] returns the
//! boilerplate text a given slide displays.

use super::shape::{Shape, TextBody, TextContent};
use super::{PptxDocument, slide::extract_plain_text_from_body};

/// A slide layout (`ppt/slideLayouts/slideLayoutN.xml`).
#[derive(Debug, Clone)]
pub struct SlideLayout {
    /// The layout's part name.
    pub part_name: String,
    /// `<p:cSld name>`.
    pub name: String,
    /// Every shape on the layout, placeholders included.
    pub shapes: Vec<Shape>,
    /// Index into [`PptxDocument::masters`] of the master this layout
    /// belongs to.
    pub master_index: Option<usize>,
    /// `<p:sldLayout showMasterSp>` — whether the master's shapes are drawn
    /// on slides using this layout (default `true`).
    pub show_master_shapes: bool,
}

/// A slide master (`ppt/slideMasters/slideMasterN.xml`).
#[derive(Debug, Clone)]
pub struct SlideMaster {
    /// The master's part name.
    pub part_name: String,
    /// `<p:cSld name>`.
    pub name: String,
    /// Every shape on the master, placeholders included.
    pub shapes: Vec<Shape>,
}

impl PptxDocument {
    /// The text of the non-placeholder shapes the slide at `index` shows
    /// from its layout and master (layout first), honouring
    /// `showMasterSp`. Placeholders are excluded: on a layout or master
    /// they hold prompt text ("Click to edit…"), and the slide's own
    /// placeholders replace them.
    ///
    /// Not part of [`Self::plain_text`] or the IR, matching python-pptx.
    pub fn static_text_for_slide(&self, index: usize) -> Vec<String> {
        let mut out = Vec::new();
        let Some(slide) = self.slides.get(index) else {
            return out;
        };
        if slide.hide_master_shapes {
            return out;
        }
        let Some(layout) = slide.layout_index.and_then(|i| self.layouts.get(i)) else {
            return out;
        };
        collect_static_text(&layout.shapes, &mut out);
        if layout.show_master_shapes {
            if let Some(master) = layout.master_index.and_then(|i| self.masters.get(i)) {
                collect_static_text(&master.shapes, &mut out);
            }
        }
        out
    }
}

fn collect_static_text(shapes: &[Shape], out: &mut Vec<String>) {
    for shape in shapes {
        match shape {
            Shape::AutoShape(auto) if auto.placeholder.is_none() => {
                if let Some(ref tb) = auto.text_body {
                    push_text(tb, out);
                }
            },
            Shape::Group(g) => collect_static_text(&g.children, out),
            _ => {},
        }
    }
}

fn push_text(tb: &TextBody, out: &mut Vec<String>) {
    // Fields (slide number, date) on a layout are placeholders for the
    // slide's own value; their cached text is not the slide's.
    let has_text = tb
        .paragraphs
        .iter()
        .flat_map(|p| &p.content)
        .any(|c| matches!(c, TextContent::Run(r) if !r.text.trim().is_empty()));
    if has_text {
        let text = extract_plain_text_from_body(tb);
        let text = text.trim();
        if !text.is_empty() {
            out.push(text.to_string());
        }
    }
}
