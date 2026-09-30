//! Slide layouts and slide masters.
//!
//! Text placed directly on a layout or master — outside any placeholder,
//! e.g. a company-name footer or boilerplate — is drawn on every slide
//! that uses it (unless the slide or layout hides the master's shapes,
//! `showMasterSp="0"`), but it belongs to no single slide.
//!
//! It is surfaced **once per document**, not once per slide: after the
//! slides, [`PptxDocument::plain_text`], [`PptxDocument::to_markdown`] and
//! the IR (and so HTML) carry a trailing section titled
//! [`MASTER_TEXT_SECTION_TITLE`] holding the static text of every layout
//! and master at least one slide shows, each distinct text once
//! ([`PptxDocument::master_static_text`]). Repeating it on every slide
//! would multiply one footer by the slide count; leaving it out (the
//! python-pptx default) loses text that is visibly on the slides and that
//! Apache POI/Tika extract. Legacy `.ppt` decks get the same section.
//!
//! Placeholders are never static text: on a layout or master they hold
//! PowerPoint's prompts ("Click to edit Master title style"), which no
//! slide shows. [`PptxDocument::static_text_for_slide`] returns what one
//! slide displays.

use super::shape::{Shape, TextBody, TextContent};
use super::{PptxDocument, slide::extract_plain_text_from_body};

/// Title of the trailing section that carries a deck's master and layout
/// static text, once per document — on every surface (`plain_text`,
/// `to_markdown`, the IR and everything rendered from it) and for both
/// `.pptx` and `.ppt`.
pub const MASTER_TEXT_SECTION_TITLE: &str = "Slide Master";

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
    pub fn static_text_for_slide(&self, index: usize) -> Vec<String> {
        self.static_bodies_for_slide(index)
            .into_iter()
            .map(|(_, text)| text)
            .collect()
    }

    /// The static layout and master text of the whole deck: every text
    /// [`Self::static_text_for_slide`] returns for any slide, each distinct
    /// text once, in first-use order. This is what the trailing
    /// [`MASTER_TEXT_SECTION_TITLE`] section of every rendering holds.
    pub fn master_static_text(&self) -> Vec<String> {
        self.master_static_bodies()
            .into_iter()
            .map(|(_, text)| text)
            .collect()
    }

    /// [`Self::master_static_text`] with each text's body, for renderers
    /// that keep its paragraphs and formatting.
    pub(crate) fn master_static_bodies(&self) -> Vec<(&TextBody, String)> {
        let mut seen = std::collections::HashSet::new();
        let mut out = Vec::new();
        for index in 0..self.slides.len() {
            for (body, text) in self.static_bodies_for_slide(index) {
                if seen.insert(text.clone()) {
                    out.push((body, text));
                }
            }
        }
        out
    }

    fn static_bodies_for_slide(&self, index: usize) -> Vec<(&TextBody, String)> {
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
        collect_static_bodies(&layout.shapes, &mut out);
        if layout.show_master_shapes {
            if let Some(master) = layout.master_index.and_then(|i| self.masters.get(i)) {
                collect_static_bodies(&master.shapes, &mut out);
            }
        }
        out
    }
}

fn collect_static_bodies<'a>(shapes: &'a [Shape], out: &mut Vec<(&'a TextBody, String)>) {
    for shape in shapes {
        match shape {
            Shape::AutoShape(auto) if auto.placeholder.is_none() => {
                if let Some(ref tb) = auto.text_body {
                    if let Some(text) = static_text(tb) {
                        out.push((tb, text));
                    }
                }
            },
            Shape::Group(g) => collect_static_bodies(&g.children, out),
            _ => {},
        }
    }
}

fn static_text(tb: &TextBody) -> Option<String> {
    // Fields (slide number, date) on a layout are placeholders for the
    // slide's own value; their cached text is not the slide's.
    let has_text = tb
        .paragraphs
        .iter()
        .flat_map(|p| &p.content)
        .any(|c| matches!(c, TextContent::Run(r) if !r.text.trim().is_empty()));
    if !has_text {
        return None;
    }
    let text = extract_plain_text_from_body(tb);
    let text = text.trim();
    (!text.is_empty()).then(|| text.to_string())
}
