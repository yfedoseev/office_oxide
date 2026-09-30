//! Inherited text formatting: list styles on shapes, slide layouts and
//! slide masters, the master's `<p:txStyles>`, and the presentation's
//! `<p:defaultTextStyle>`.
//!
//! Per ECMA-376 Part 1 §19.3.1 / §21.1.2, a run's character properties and
//! its paragraph's alignment are resolved, most specific first, from:
//!
//! 1. the run's own `<a:rPr>` and the paragraph's `<a:pPr>`;
//! 2. the shape's own `<a:lstStyle>`;
//! 3. the matching placeholder on the slide layout (its `<a:lstStyle>`);
//! 4. the matching placeholder on the slide master;
//! 5. the master's `<p:txStyles>` — `titleStyle` for title placeholders,
//!    `bodyStyle` for body/object placeholders, `otherStyle` for every
//!    other placeholder and for text in non-placeholder shapes;
//! 6. the presentation's `<p:defaultTextStyle>`.
//!
//! Each list style holds one `<a:lvlNpPr>` per outline level (1–9) and a
//! level-agnostic `<a:defPPr>`. Only the properties this crate's model
//! carries are resolved (bold, italic, underline, size, sRGB colour,
//! alignment); theme colour/font references are left unset rather than
//! guessed. The theme's own `objectDefaults` are not consulted.

use quick_xml::events::Event;

use crate::core::Result as CoreResult;
use crate::core::xml;
use crate::ir::ParagraphAlignment;

/// Default formatting from one `<a:lvlNpPr>`/`<a:defPPr>`. Every field is
/// `None` when the style doesn't specify it — the same "unset, don't
/// guess" contract every other formatting field in this crate's PPTX
/// reader follows.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct MasterRunDefaults {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub font_size_hundredths_pt: Option<u32>,
    pub color_rgb: Option<[u8; 3]>,
    pub alignment: Option<ParagraphAlignment>,
    /// The level's bullet (`<a:buNone>`/`<a:buChar>`/`<a:buAutoNum>`,
    /// ECMA-376 Part 1 §21.1.2.4): a body placeholder's paragraphs are
    /// bulleted because the master's `bodyStyle` says so, not their own
    /// `<a:pPr>`.
    pub bullet: Option<super::shape::BulletStyle>,
}

impl MasterRunDefaults {
    /// Fill every unset field from `lower` (a less specific style).
    fn fill_from(&mut self, lower: &MasterRunDefaults) {
        fill(&mut self.bold, &lower.bold);
        fill(&mut self.italic, &lower.italic);
        fill(&mut self.underline, &lower.underline);
        fill(&mut self.font_size_hundredths_pt, &lower.font_size_hundredths_pt);
        fill(&mut self.color_rgb, &lower.color_rgb);
        fill(&mut self.alignment, &lower.alignment);
        fill(&mut self.bullet, &lower.bullet);
    }
}

fn fill<T: Clone>(slot: &mut Option<T>, from: &Option<T>) {
    if slot.is_none() {
        slot.clone_from(from);
    }
}

/// Number of outline levels a list style defines (`lvl1pPr`…`lvl9pPr`).
pub(crate) const LEVELS: usize = 9;

/// A list style: per-outline-level defaults, with `<a:defPPr>` already
/// folded into each level.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct LevelStyles {
    levels: [Option<MasterRunDefaults>; LEVELS],
}

impl LevelStyles {
    /// The defaults for 0-based outline `level` (clamped to the last).
    pub(crate) fn level(&self, level: u32) -> Option<&MasterRunDefaults> {
        let i = usize::try_from(level).unwrap_or(LEVELS - 1).min(LEVELS - 1);
        self.levels[i].as_ref()
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.levels.iter().all(Option::is_none)
    }
}

/// A slide master's `<p:txStyles>`.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct MasterTextStyles {
    pub title: LevelStyles,
    pub body: LevelStyles,
    pub other: LevelStyles,
}

/// The list style of one placeholder on a slide layout or master.
#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct PlaceholderStyle {
    /// `<p:ph type>`; `None` means the schema default, `obj`.
    pub ph_type: Option<String>,
    /// `<p:ph idx>`; `None` means the schema default, 0.
    pub idx: Option<u32>,
    pub levels: LevelStyles,
}

/// Which `<p:txStyles>` entry a placeholder type falls back to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StyleCategory {
    Title,
    Body,
    Other,
}

/// The category of a slide shape: `None` is a non-placeholder shape;
/// `Some(None)` a placeholder without a `type` (the default, `obj`).
pub(crate) fn category(ph_type: Option<Option<&str>>) -> StyleCategory {
    match ph_type {
        None => StyleCategory::Other,
        Some(Some("title" | "ctrTitle")) => StyleCategory::Title,
        // ST_PlaceholderType (§19.7.10): content placeholders use the body
        // style; date, footer, header and slide number use the other style.
        Some(None | Some("body" | "obj" | "subTitle" | "chart" | "tbl" | "clipArt" | "dgm")) => {
            StyleCategory::Body
        },
        Some(Some("media" | "pic")) => StyleCategory::Body,
        Some(Some(_)) => StyleCategory::Other,
    }
}

/// Parse a `ppt/slideMasters/slideMasterN.xml` document's `<p:txStyles>`.
/// Returns empty styles (not an error) on malformed/missing input —
/// style inheritance is a best-effort enhancement, never a hard
/// requirement for reading the rest of the file.
pub(crate) fn parse_master_text_styles(xml_data: &[u8]) -> MasterTextStyles {
    let mut reader = make_reader(xml_data);
    let mut styles = MasterTextStyles::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let target = match e.local_name().as_ref() {
                    "titleStyle" => &mut styles.title,
                    "bodyStyle" => &mut styles.body,
                    "otherStyle" => &mut styles.other,
                    _ => continue,
                };
                let local = e.local_name().as_ref().to_string();
                match parse_level_styles(&mut reader, &local) {
                    Ok(s) => *target = s,
                    Err(_) => break,
                }
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {},
        }
    }
    styles
}

/// Parse `ppt/presentation.xml`'s `<p:defaultTextStyle>`.
pub(crate) fn parse_default_text_style(pres_xml: &[u8]) -> LevelStyles {
    let mut reader = make_reader(pres_xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == "defaultTextStyle" => {
                return parse_level_styles(&mut reader, "defaultTextStyle").unwrap_or_default();
            },
            Ok(Event::Eof) | Err(_) => return LevelStyles::default(),
            _ => {},
        }
    }
}

/// The list styles of every placeholder on a slide layout or master
/// (`<p:sp>` with `<p:ph>` whose `<p:txBody>` has an `<a:lstStyle>`).
/// Placeholders without a list style are listed too, with empty styles,
/// so a match by `idx` still finds them.
pub(crate) fn parse_placeholder_styles(xml_data: &[u8]) -> Vec<PlaceholderStyle> {
    let mut reader = make_reader(xml_data);
    let mut out = Vec::new();
    // The `<p:sp>` being read: its placeholder (if seen yet) and styles.
    let mut current: Option<(Option<PlaceholderStyle>, LevelStyles)> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                "sp" => current = Some((None, LevelStyles::default())),
                "lstStyle" => {
                    let styles = match parse_level_styles(&mut reader, "lstStyle") {
                        Ok(s) => s,
                        Err(_) => break,
                    };
                    if let Some((_, ref mut levels)) = current {
                        *levels = styles;
                    }
                },
                "ph" => set_placeholder(&mut current, e),
                _ => {},
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == "ph" => {
                set_placeholder(&mut current, e);
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == "sp" => {
                if let Some((Some(mut ph), levels)) = current.take() {
                    ph.levels = levels;
                    out.push(ph);
                }
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {},
        }
    }
    out
}

fn set_placeholder(
    current: &mut Option<(Option<PlaceholderStyle>, LevelStyles)>,
    e: &quick_xml::events::BytesStart,
) {
    if let Some((ph, _)) = current {
        *ph = Some(PlaceholderStyle {
            ph_type: xml::optional_attr_str(e, "type")
                .ok()
                .flatten()
                .map(|v| v.into_owned()),
            idx: xml::optional_attr_str(e, "idx")
                .ok()
                .flatten()
                .and_then(|v| v.parse().ok()),
            levels: LevelStyles::default(),
        });
    }
}

fn make_reader(xml_data: &[u8]) -> quick_xml::Reader<&[u8]> {
    let mut reader = quick_xml::Reader::from_reader(xml_data);
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    reader
}

/// Parse a list-style container (`a:lstStyle`, `p:titleStyle`,
/// `p:defaultTextStyle`, …) whose Start tag was just read, through its end
/// tag `end_local`: `<a:defPPr>` and `<a:lvl1pPr>`…`<a:lvl9pPr>`
/// (`CT_TextListStyle`, ECMA-376 Part 1 §21.1.2.4.12).
pub(crate) fn parse_level_styles(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &str,
) -> CoreResult<LevelStyles> {
    let mut styles = LevelStyles::default();
    let mut def: Option<MasterRunDefaults> = None;
    loop {
        let event = reader.read_event()?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                let slot = match local.as_ref() {
                    "defPPr" => None,
                    name => match lvl_index(name) {
                        Some(i) => Some(i),
                        None => {
                            if is_start {
                                xml::skip_element_fast(reader)?;
                            }
                            continue;
                        },
                    },
                };
                let parsed = if is_start {
                    parse_lvl_pr(reader, e)?
                } else {
                    MasterRunDefaults {
                        alignment: parse_algn(e)?,
                        ..Default::default()
                    }
                };
                match slot {
                    Some(i) => styles.levels[i] = Some(parsed),
                    None => def = Some(parsed),
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => break,
            _ => {},
        }
    }
    if let Some(def) = def {
        for level in &mut styles.levels {
            match level {
                Some(l) => l.fill_from(&def),
                None => *level = Some(def.clone()),
            }
        }
    }
    Ok(styles)
}

/// `lvl1pPr` → 0 … `lvl9pPr` → 8.
fn lvl_index(name: &str) -> Option<usize> {
    let digits = name.strip_prefix("lvl")?.strip_suffix("pPr")?.as_bytes();
    match digits {
        [d @ b'1'..=b'9'] => Some(usize::from(d - b'1')),
        _ => None,
    }
}

fn parse_algn(e: &quick_xml::events::BytesStart) -> CoreResult<Option<ParagraphAlignment>> {
    Ok(xml::optional_attr_str(e, "algn")?.and_then(|v| match v.as_ref() {
        "l" => Some(ParagraphAlignment::Left),
        "ctr" => Some(ParagraphAlignment::Center),
        "r" => Some(ParagraphAlignment::Right),
        "just" | "justLow" => Some(ParagraphAlignment::Justify),
        "dist" | "thaiDist" => Some(ParagraphAlignment::Distribute),
        _ => None,
    }))
}

/// Parse a non-self-closing `<a:lvlNpPr algn="…">…<a:defRPr …/>…</a:lvlNpPr>`:
/// the paragraph's own `algn` attribute plus its child `<a:defRPr>`'s
/// character formatting, reading through the element's own end tag.
fn parse_lvl_pr(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> CoreResult<MasterRunDefaults> {
    let mut defaults = MasterRunDefaults {
        alignment: parse_algn(start)?,
        ..Default::default()
    };
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == "defRPr" => {
                parse_def_rpr(reader, e, &mut defaults)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == "defRPr" => {
                apply_rpr_attrs(e, &mut defaults)?;
            },
            Event::Empty(ref e) if depth == 1 => {
                if let Some(b) = super::slide::parse_bullet(e)? {
                    defaults.bullet = Some(b);
                }
            },
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(defaults)
}

/// `<a:defRPr b="1" i="0" u="sng" sz="4400">…<a:solidFill><a:srgbClr
/// val="…"/></a:solidFill>…</a:defRPr>` — same attribute/child shape as
/// an ordinary run's `<a:rPr>`, deliberately parsed fresh here (rather
/// than reused from `slide.rs`'s `parse_run_properties`) to avoid
/// disturbing that function's more complex hyperlink/nested-field
/// handling, which `defRPr` never carries.
fn parse_def_rpr(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    defaults: &mut MasterRunDefaults,
) -> CoreResult<()> {
    apply_rpr_attrs(start, defaults)?;
    let mut in_solid_fill = false;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == "solidFill" => {
                in_solid_fill = true;
            },
            Event::End(ref e) if e.local_name().as_ref() == "solidFill" => {
                in_solid_fill = false;
            },
            Event::Empty(ref e) if in_solid_fill && e.local_name().as_ref() == "srgbClr" => {
                if defaults.color_rgb.is_none() {
                    defaults.color_rgb = parse_srgb_clr(e);
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == "defRPr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn apply_rpr_attrs(
    e: &quick_xml::events::BytesStart,
    defaults: &mut MasterRunDefaults,
) -> CoreResult<()> {
    if let Some(v) = xml::optional_attr_str(e, "b")? {
        defaults.bold = Some(v.as_ref() != "0");
    }
    if let Some(v) = xml::optional_attr_str(e, "i")? {
        defaults.italic = Some(v.as_ref() != "0");
    }
    if let Some(v) = xml::optional_attr_str(e, "u")? {
        defaults.underline = Some(v.into_owned());
    }
    if let Some(v) = xml::optional_attr_str(e, "sz")? {
        defaults.font_size_hundredths_pt = v.parse::<u32>().ok();
    }
    Ok(())
}

fn parse_srgb_clr(e: &quick_xml::events::BytesStart) -> Option<[u8; 3]> {
    let val = xml::optional_attr_str(e, "val").ok().flatten()?;
    let s = val.as_ref();
    if s.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some([r, g, b])
}

/// Everything a slide's shapes inherit from, resolved once per layout.
#[derive(Debug, Clone, Default)]
pub(crate) struct StyleChain {
    pub layout_placeholders: std::sync::Arc<Vec<PlaceholderStyle>>,
    pub master_placeholders: std::sync::Arc<Vec<PlaceholderStyle>>,
    pub master_text_styles: std::sync::Arc<MasterTextStyles>,
    pub presentation_default: std::sync::Arc<LevelStyles>,
}

impl StyleChain {
    pub(crate) fn is_empty(&self) -> bool {
        self.layout_placeholders.iter().all(|p| p.levels.is_empty())
            && self.master_placeholders.iter().all(|p| p.levels.is_empty())
            && self.master_text_styles.title.is_empty()
            && self.master_text_styles.body.is_empty()
            && self.master_text_styles.other.is_empty()
            && self.presentation_default.is_empty()
    }

    /// The resolved defaults at outline `level` for a shape whose
    /// placeholder is `ph` (`None` for a non-placeholder shape).
    pub(crate) fn resolve(
        &self,
        ph: Option<(Option<&str>, Option<u32>)>,
        level: u32,
    ) -> MasterRunDefaults {
        let mut out = MasterRunDefaults::default();
        if let Some((ph_type, idx)) = ph {
            if let Some(l) = match_layout(&self.layout_placeholders, ph_type, idx) {
                if let Some(d) = l.levels.level(level) {
                    out.fill_from(d);
                }
            }
            if let Some(m) = match_master(&self.master_placeholders, ph_type) {
                if let Some(d) = m.levels.level(level) {
                    out.fill_from(d);
                }
            }
        }
        let tx = &self.master_text_styles;
        let styles = match category(ph.map(|(t, _)| t)) {
            StyleCategory::Title => &tx.title,
            StyleCategory::Body => &tx.body,
            StyleCategory::Other => &tx.other,
        };
        if let Some(d) = styles.level(level) {
            out.fill_from(d);
        }
        if let Some(d) = self.presentation_default.level(level) {
            out.fill_from(d);
        }
        out
    }
}

/// The layout placeholder a slide placeholder inherits from: the one with
/// the same `idx` (default 0), else the first with the same `type`
/// (python-pptx and PowerPoint match on idx).
fn match_layout<'a>(
    phs: &'a [PlaceholderStyle],
    ph_type: Option<&str>,
    idx: Option<u32>,
) -> Option<&'a PlaceholderStyle> {
    let idx = idx.unwrap_or(0);
    phs.iter()
        .find(|p| p.idx.unwrap_or(0) == idx)
        .or_else(|| phs.iter().find(|p| p.ph_type.as_deref() == ph_type))
}

/// The master placeholder a layout placeholder inherits from, by type:
/// title types map to the master's title, content types to its body.
fn match_master<'a>(
    phs: &'a [PlaceholderStyle],
    ph_type: Option<&str>,
) -> Option<&'a PlaceholderStyle> {
    let want = match category(Some(ph_type)) {
        StyleCategory::Title => "title",
        StyleCategory::Body => "body",
        StyleCategory::Other => ph_type.unwrap_or("obj"),
    };
    phs.iter().find(|p| {
        let t = p.ph_type.as_deref().unwrap_or("obj");
        t == want || (want == "title" && t == "ctrTitle")
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_title_and_body_level_defaults_parsed() {
        let xml = br#"<?xml version="1.0"?>
<p:txStyles xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:titleStyle>
    <a:lvl1pPr algn="ctr">
      <a:defRPr sz="4400" b="1">
        <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
      </a:defRPr>
    </a:lvl1pPr>
  </p:titleStyle>
  <p:bodyStyle>
    <a:lvl1pPr algn="l">
      <a:defRPr sz="3200"/>
    </a:lvl1pPr>
    <a:lvl2pPr>
      <a:defRPr sz="2800"/>
    </a:lvl2pPr>
  </p:bodyStyle>
</p:txStyles>"#;

        let styles = parse_master_text_styles(xml);
        let title = styles.title.level(0).expect("title level 1");
        assert_eq!(title.alignment, Some(ParagraphAlignment::Center));
        assert_eq!(title.bold, Some(true));
        assert_eq!(title.font_size_hundredths_pt, Some(4400));
        assert_eq!(title.color_rgb, Some([0x11, 0x22, 0x33]));

        let body = styles.body.level(0).expect("body level 1");
        assert_eq!(body.alignment, Some(ParagraphAlignment::Left));
        assert_eq!(body.font_size_hundredths_pt, Some(3200));
        assert_eq!(body.bold, None);
        // Level 2 is its own entry — only level 1 used to be read.
        assert_eq!(styles.body.level(1).unwrap().font_size_hundredths_pt, Some(2800));
        assert!(styles.body.level(2).is_none());
    }

    #[test]
    fn test_self_closing_lvl1_ppr_with_only_algn_is_not_an_error() {
        let xml = br#"<?xml version="1.0"?>
<p:txStyles xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:titleStyle><a:lvl1pPr algn="r"/></p:titleStyle>
</p:txStyles>"#;
        let styles = parse_master_text_styles(xml);
        assert_eq!(styles.title.level(0).unwrap().alignment, Some(ParagraphAlignment::Right));
        assert!(styles.body.is_empty());
    }

    #[test]
    fn test_missing_tx_styles_yields_empty() {
        let xml = br#"<?xml version="1.0"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let styles = parse_master_text_styles(xml);
        assert!(styles.title.is_empty() && styles.body.is_empty() && styles.other.is_empty());
    }

    #[test]
    fn test_truncated_xml_does_not_panic() {
        let styles = parse_master_text_styles(b"<p:txStyles><p:titleStyle><a:lvl1pPr");
        assert!(styles.title.is_empty());
    }

    /// `<a:defPPr>` applies to every level that does not override it.
    #[test]
    fn test_def_ppr_fills_every_level() {
        let xml = br#"<p:presentation xmlns:a="a" xmlns:p="p"><p:defaultTextStyle>
            <a:defPPr><a:defRPr sz="1800" b="0"/></a:defPPr>
            <a:lvl3pPr><a:defRPr sz="1400"/></a:lvl3pPr>
          </p:defaultTextStyle></p:presentation>"#;
        let d = parse_default_text_style(xml);
        assert_eq!(d.level(0).unwrap().font_size_hundredths_pt, Some(1800));
        assert_eq!(d.level(2).unwrap().font_size_hundredths_pt, Some(1400));
        assert_eq!(d.level(2).unwrap().bold, Some(false), "filled from defPPr");
        assert_eq!(d.level(8).unwrap().font_size_hundredths_pt, Some(1800));
    }

    #[test]
    fn test_lvl_index() {
        assert_eq!(lvl_index("lvl1pPr"), Some(0));
        assert_eq!(lvl_index("lvl9pPr"), Some(8));
        assert_eq!(lvl_index("lvl0pPr"), None);
        assert_eq!(lvl_index("lvl10pPr"), None);
        assert_eq!(lvl_index("defPPr"), None);
    }
}
