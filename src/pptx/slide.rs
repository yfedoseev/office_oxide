use quick_xml::events::Event;

use crate::core::relationships::{Relationships, TargetMode};
use crate::core::xml;

use super::shape::{
    AutoShape, BulletStyle, ConnectorShape, GraphicContent, GraphicFrame, GroupShape,
    HyperlinkInfo, HyperlinkTarget, PictureShape, PlaceholderInfo, Shape, ShapePosition, Table,
    TableCell, TableRow, TextBody, TextContent, TextField, TextParagraph, TextRun,
};

type CoreResult<T> = crate::core::Result<T>;

/// Parsed run properties: (bold, italic, strikethrough, hyperlink, font_size_hundredths_pt).
///
/// PPTX `<a:rPr sz="..."/>` carries font size in hundredths of a point
/// (e.g. `sz="1800"` = 18 pt). Carrying it through the parser is what
/// keeps PDF→PPTX→PDF round-trips from defaulting every paragraph to
/// the writer's 12 pt fallback (which inflated 8-page A4 sources to
/// ~30 pages).
#[derive(Default)]
struct RunProps {
    bold: Option<bool>,
    italic: Option<bool>,
    strikethrough: bool,
    hyperlink: Option<HyperlinkInfo>,
    font_size_hundredths_pt: Option<u32>,
    color_rgb: Option<[u8; 3]>,
    underline: Option<String>,
    font_name: Option<String>,
    baseline: Option<i32>,
    caps: Option<String>,
    char_spacing_hundredths_pt: Option<i32>,
}

/// Read the `<a:rPr>` attributes shared by the Start and Empty forms.
fn run_props_from_attrs(e: &quick_xml::events::BytesStart) -> CoreResult<RunProps> {
    let strike = xml::optional_attr_str(e, b"strike")?;
    Ok(RunProps {
        bold: parse_bool_attr(e, b"b")?,
        italic: parse_bool_attr(e, b"i")?,
        strikethrough: strike.as_deref().is_some_and(|v| v != "noStrike"),
        font_size_hundredths_pt: parse_u32_attr(e, b"sz")?,
        // `u`, `baseline`, `cap` and `spc` were parsed by no one, so
        // underline in particular — the third most common piece of direct
        // formatting — never reached the IR from PPTX even though it did
        // from DOCX.
        underline: xml::optional_attr_str(e, b"u")?.map(|v| v.into_owned()),
        baseline: xml::optional_attr_str(e, b"baseline")?.and_then(|v| v.parse().ok()),
        caps: xml::optional_attr_str(e, b"cap")?.map(|v| v.into_owned()),
        char_spacing_hundredths_pt: xml::optional_attr_str(e, b"spc")?.and_then(|v| v.parse().ok()),
        ..Default::default()
    })
}

/// A parsed PPTX slide.
#[derive(Debug, Clone, Default)]
pub struct Slide {
    /// Slide name from the `<p:cSld name="...">` attribute.
    pub name: String,
    /// All top-level shapes on this slide.
    pub shapes: Vec<Shape>,
    /// Speaker notes body, if a notes slide is present. Kept as the
    /// structured `TextBody` (same model ordinary slide body text uses)
    /// rather than flattened text, so bold/italic/bullets/numbering in
    /// notes survive through to the IR.
    pub notes: Option<TextBody>,
    /// Solid background colour (RGB) extracted from the slide's
    /// `<p:cSld><p:bg><p:bgPr><a:solidFill>` element. Only the solid
    /// case is parsed; gradient / image / theme-reference fills are
    /// dropped silently and surface as `None`.
    pub background_rgb: Option<[u8; 3]>,
    /// `<p:sld show="0">` — the slide is hidden from the slideshow. Its
    /// content is still extracted, but a consumer can now tell that the
    /// author excluded it, which was impossible before.
    pub hidden: bool,
    /// Comments attached to this slide, from `ppt/comments/*.xml`.
    pub comments: Vec<SlideComment>,
}

/// A comment attached to a slide (`ppt/comments/modernComment*.xml` or the
/// legacy `ppt/comments/comment*.xml`).
#[derive(Debug, Clone)]
pub struct SlideComment {
    /// Author name, when the deck's author list resolves the id.
    pub author: Option<String>,
    /// Comment body text.
    pub text: String,
}

/// Create a fast reader that does NOT trim text content.
fn make_content_reader(xml_data: &[u8]) -> quick_xml::Reader<&[u8]> {
    let mut reader = quick_xml::Reader::from_reader(xml_data);
    reader.config_mut().check_end_names = false;
    reader.config_mut().check_comments = false;
    reader
}

impl Slide {
    /// Parse a slide from its XML data.
    pub(crate) fn parse(
        xml_data: &[u8],
        name: String,
        rels: &Relationships,
        media: &std::collections::HashMap<String, (Vec<u8>, String)>,
        charts: &std::collections::HashMap<String, Vec<String>>,
    ) -> CoreResult<Self> {
        let mut reader = make_content_reader(xml_data);
        let mut shapes = Vec::new();
        let mut background_rgb = None;
        let mut hidden = false;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"sld" => {
                    hidden = xml::optional_attr_str(e, b"show")?
                        .is_some_and(|v| matches!(v.as_ref(), "0" | "false"));
                },
                Event::Start(ref e) if e.local_name().as_ref() == b"bg" => {
                    background_rgb = parse_slide_bg(&mut reader)?;
                },
                Event::Start(ref e) if e.local_name().as_ref() == b"spTree" => {
                    shapes = parse_shape_tree(&mut reader, rels, media, charts)?;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(Slide {
            name,
            shapes,
            notes: None,
            background_rgb,
            hidden,
            comments: Vec::new(),
        })
    }
}

/// Parse `<p:bg>` looking for a single solid-fill colour.
///
/// Returns `Some([r, g, b])` if the background is a `<p:bgPr>` with an
/// `<a:solidFill><a:srgbClr val="RRGGBB"/>`. All other forms (gradient,
/// blip / image, scheme / theme references via `<p:bgRef>`) return
/// `None` — the renderer silently falls back to no background, which
/// matches "minimum theme-background support" per the v0.3.42 plan.
fn parse_slide_bg(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<[u8; 3]>> {
    let mut rgb = None;
    let mut depth = 1u32;
    let mut in_solid_fill = false;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                depth += 1;
                if e.local_name().as_ref() == b"solidFill" {
                    in_solid_fill = true;
                }
            },
            Event::Empty(ref e) => {
                if in_solid_fill && e.local_name().as_ref() == b"srgbClr" {
                    if let Some(val) = xml::optional_attr_str(e, b"val")? {
                        rgb = parse_hex_rgb(val.as_ref());
                    }
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"solidFill" {
                    in_solid_fill = false;
                }
                depth -= 1;
                if depth == 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(rgb)
}

/// Parse a 6-character hex colour (e.g. `"0E273B"`) into `[r, g, b]`.
fn parse_hex_rgb(s: &str) -> Option<[u8; 3]> {
    let bytes = s.as_bytes();
    if bytes.len() != 6 {
        return None;
    }
    let h = |hi, lo| -> Option<u8> {
        let n = |c: u8| match c {
            b'0'..=b'9' => Some(c - b'0'),
            b'a'..=b'f' => Some(c - b'a' + 10),
            b'A'..=b'F' => Some(c - b'A' + 10),
            _ => None,
        };
        Some(n(hi)? * 16 + n(lo)?)
    };
    Some([
        h(bytes[0], bytes[1])?,
        h(bytes[2], bytes[3])?,
        h(bytes[4], bytes[5])?,
    ])
}

// ---------------------------------------------------------------------------
// Shape tree parsing
// ---------------------------------------------------------------------------

fn parse_shape_tree(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    charts: &std::collections::HashMap<String, Vec<String>>,
) -> CoreResult<Vec<Shape>> {
    parse_shape_tree_until(reader, rels, media, charts, b"spTree")
}

/// Shared shape-tree loop, parameterized on the closing tag so it can also
/// read the contents of an `<mc:Choice>`/`<mc:Fallback>` branch (see
/// [`parse_alternate_content`]).
fn parse_shape_tree_until(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    charts: &std::collections::HashMap<String, Vec<String>>,
    end_local: &[u8],
) -> CoreResult<Vec<Shape>> {
    let mut shapes = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"sp" => shapes.push(parse_auto_shape(reader, rels)?),
                b"pic" => shapes.push(parse_picture(reader, rels, media)?),
                b"grpSp" => shapes.push(parse_group_shape(reader, rels, media, charts)?),
                b"graphicFrame" => shapes.push(parse_graphic_frame(reader, rels, charts)?),
                b"cxnSp" => shapes.push(parse_connector(reader)?),
                b"AlternateContent" => {
                    shapes.extend(parse_alternate_content(reader, rels, media, charts)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == end_local => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(shapes)
}

/// `<mc:AlternateContent>` wraps two or more renderings of the same shape
/// behind a markup-compatibility switch — typically a modern extension
/// (`<mc:Choice Requires="…">`, e.g. an OMML equation) and a plain
/// `<mc:Fallback>` for older readers. There's no namespace-support
/// negotiation here: this takes the first `Choice` branch that actually
/// yields a recognized shape, and falls back to `Fallback` otherwise —
/// strictly better than the old behavior of skipping the whole block,
/// which silently dropped shapes like equation text boxes.
fn parse_alternate_content(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    charts: &std::collections::HashMap<String, Vec<String>>,
) -> CoreResult<Vec<Shape>> {
    let mut shapes: Vec<Shape> = Vec::new();
    let mut have_choice = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"Choice" => {
                    let s = parse_shape_tree_until(reader, rels, media, charts, b"Choice")?;
                    if !s.is_empty() {
                        shapes = s;
                        have_choice = true;
                    }
                },
                b"Fallback" => {
                    let s = parse_shape_tree_until(reader, rels, media, charts, b"Fallback")?;
                    if !have_choice && shapes.is_empty() {
                        shapes = s;
                    }
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"AlternateContent" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(shapes)
}

// ---------------------------------------------------------------------------
// AutoShape (p:sp)
// ---------------------------------------------------------------------------

fn parse_auto_shape(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut position = None;
    let mut text_body = None;
    let mut placeholder = None;
    let mut hyperlink = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvSpPr" => {
                    let props = parse_nv_common_props(reader, rels)?;
                    id = props.0;
                    name = props.1;
                    alt_text = props.2;
                    placeholder = props.3;
                    hyperlink = props.4;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader, b"spPr")?;
                },
                b"txBody" => {
                    text_body = Some(parse_text_body(reader, rels)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"sp" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Shape::AutoShape(AutoShape {
        id,
        name,
        alt_text,
        position,
        hyperlink,
        text_body,
        placeholder,
    }))
}

// ---------------------------------------------------------------------------
// PictureShape (p:pic)
// ---------------------------------------------------------------------------

fn parse_picture(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut position = None;
    let mut embed_rid: Option<String> = None;
    let mut hyperlink = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvPicPr" => {
                    let props = parse_nv_pic_props(reader, rels)?;
                    id = props.0;
                    name = props.1;
                    alt_text = props.2;
                    hyperlink = props.3;
                },
                b"blipFill" => {
                    embed_rid = parse_blip_fill_embed(reader)?;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader, b"spPr")?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"pic" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    let (data, format) = match embed_rid.as_deref().and_then(|rid| media.get(rid)) {
        Some((bytes, ext)) => (Some(bytes.clone()), Some(ext.clone())),
        None => (None, None),
    };

    Ok(Shape::Picture(PictureShape {
        id,
        name,
        alt_text,
        position,
        embed_rid,
        data,
        format,
        hyperlink,
    }))
}

/// Parse `<p:blipFill>…<a:blip r:embed="rIdN"/>…</p:blipFill>` and
/// return the `r:embed` attribute, if present. Other contents (stretch,
/// crop, tile) are skipped — only the embed rId is needed to resolve
/// the underlying media part.
fn parse_blip_fill_embed(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<String>> {
    let mut embed: Option<String> = None;
    let mut depth: u32 = 1;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"blip" && embed.is_none() {
                    embed = read_blip_embed_attr(e)?;
                }
                depth += 1;
            },
            Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"blip" && embed.is_none() {
                    embed = read_blip_embed_attr(e)?;
                }
            },
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
    Ok(embed)
}

fn read_blip_embed_attr(e: &quick_xml::events::BytesStart) -> CoreResult<Option<String>> {
    // `<a:blip>` carries `r:embed="rIdN"` (DrawingML namespace `a:`,
    // relationship namespace `r:`). The attribute may be present in
    // either the `Empty` or `Start` form; both routes feed this helper.
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(crate::core::Error::from)?;
        let key = attr.key.as_ref();
        let is_embed = key == b"r:embed" || key.ends_with(b":embed") || key == b"embed";
        if is_embed {
            return Ok(Some(crate::core::xml::unescape_attr_value(&attr)?));
        }
    }
    Ok(None)
}

// ---------------------------------------------------------------------------
// GroupShape (p:grpSp)
// ---------------------------------------------------------------------------

fn parse_group_shape(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    charts: &std::collections::HashMap<String, Vec<String>>,
) -> CoreResult<Shape> {
    // Groups nest, so this is the recursion an adversarial deck drives.
    // Past the limit the subtree is skipped: a stack overflow aborts the
    // process and no caller can catch it.
    let Some(_depth) = xml::DepthGuard::enter() else {
        xml::skip_element_fast(reader)?;
        return Ok(Shape::Group(GroupShape {
            id: 0,
            name: String::new(),
            position: None,
            children: Vec::new(),
        }));
    };
    let mut id = 0u32;
    let mut name = String::new();
    let mut position = None;
    let mut children = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvGrpSpPr" => {
                    let props = parse_nv_id_name(reader, b"nvGrpSpPr")?;
                    id = props.0;
                    name = props.1;
                },
                b"grpSpPr" => {
                    position = parse_shape_properties(reader, b"grpSpPr")?;
                },
                b"sp" => children.push(parse_auto_shape(reader, rels)?),
                b"pic" => children.push(parse_picture(reader, rels, media)?),
                b"grpSp" => children.push(parse_group_shape(reader, rels, media, charts)?),
                b"graphicFrame" => children.push(parse_graphic_frame(reader, rels, charts)?),
                b"cxnSp" => children.push(parse_connector(reader)?),
                b"AlternateContent" => {
                    children.extend(parse_alternate_content(reader, rels, media, charts)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"grpSp" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Shape::Group(GroupShape {
        id,
        name,
        position,
        children,
    }))
}

// ---------------------------------------------------------------------------
// GraphicFrame (p:graphicFrame)
// ---------------------------------------------------------------------------

/// Collect every `<a:t>` text value inside a subtree, reading through the
/// matching `</end_local>`. Used for graphic payloads we don't model
/// structurally (SmartArt, charts) so their words still reach the IR.
fn collect_a_t_text(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<Vec<String>> {
    let mut out = Vec::new();
    let mut depth = 1i32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"t" {
                    let t = xml::read_text_content_fast(reader)?;
                    let t = t.trim();
                    if !t.is_empty() {
                        out.push(t.to_string());
                    }
                } else {
                    depth += 1;
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == end_local && depth <= 1 {
                    break;
                }
                depth -= 1;
                if depth <= 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// Find `<c:chart r:id="…"/>`'s relationship id inside a `<a:graphicData>`
/// subtree, reading through the matching `</end_local>` regardless of
/// whether a chart reference was found (so the reader position stays
/// correct either way).
fn find_chart_rid(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_local: &[u8],
) -> CoreResult<Option<String>> {
    let mut rid = None;
    let mut depth = 1i32;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if rid.is_none() && e.local_name().as_ref() == b"chart" {
                    rid = xml::optional_attr_str(e, b"r:id")?
                        .filter(|v| !v.is_empty())
                        .map(|v| v.into_owned());
                }
                depth += 1;
            },
            Event::Empty(ref e) => {
                if rid.is_none() && e.local_name().as_ref() == b"chart" {
                    rid = xml::optional_attr_str(e, b"r:id")?
                        .filter(|v| !v.is_empty())
                        .map(|v| v.into_owned());
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == end_local && depth <= 1 {
                    break;
                }
                depth -= 1;
                if depth <= 0 {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(rid)
}

fn parse_graphic_frame(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
    charts: &std::collections::HashMap<String, Vec<String>>,
) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut position = None;
    let mut content = GraphicContent::Unknown;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                match e.local_name().as_ref() {
                    b"nvGraphicFramePr" => {
                        let props = parse_nv_id_name(reader, b"nvGraphicFramePr")?;
                        id = props.0;
                        name = props.1;
                    },
                    b"xfrm" => {
                        position = parse_xfrm(reader)?;
                    },
                    // <a:graphic> is a wrapper — keep parsing to find <a:graphicData>
                    b"graphic" => {},
                    b"graphicData" => {
                        let uri = xml::optional_attr_str(e, b"uri")?;
                        if uri.as_deref()
                            == Some("http://schemas.openxmlformats.org/drawingml/2006/table")
                        {
                            content = parse_graphic_data_table(reader, rels)?;
                        } else if uri.as_deref()
                            == Some("http://schemas.openxmlformats.org/drawingml/2006/chart")
                        {
                            // The slide XML holds only a reference —
                            // <c:chart r:id="rIdN"/> — with no text of its
                            // own; the title, axis labels, category names
                            // and cached data values all live in the
                            // separate part that id resolves to
                            // (ppt/charts/chartN.xml), pre-read into
                            // `charts`.
                            let rid = find_chart_rid(reader, b"graphicData")?;
                            let texts = rid.and_then(|r| charts.get(&r)).cloned();
                            content = match texts {
                                Some(t) if !t.is_empty() => GraphicContent::Text(t),
                                _ => GraphicContent::Unknown,
                            };
                        } else {
                            // Everything else — SmartArt diagrams, embedded
                            // objects — used to be skipped wholesale along
                            // with charts. We can't render them, but their
                            // `<a:t>` runs are document text.
                            let texts = collect_a_t_text(reader, b"graphicData")?;
                            content = if texts.is_empty() {
                                GraphicContent::Unknown
                            } else {
                                GraphicContent::Text(texts)
                            };
                        }
                    },
                    _ => {
                        xml::skip_element_fast(reader)?;
                    },
                }
            },
            Event::End(ref e) if e.local_name().as_ref() == b"graphicFrame" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Shape::GraphicFrame(GraphicFrame {
        id,
        name,
        position,
        content,
    }))
}

fn parse_graphic_data_table(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<GraphicContent> {
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"tbl" => {
                let table = parse_table(reader, rels)?;
                // Skip to end of graphicData
                skip_to_end_of(reader, b"graphicData")?;
                return Ok(GraphicContent::Table(table));
            },
            Event::End(ref e) if e.local_name().as_ref() == b"graphicData" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(GraphicContent::Unknown)
}

/// Skip remaining events until the end tag for the given element.
fn skip_to_end_of(reader: &mut quick_xml::Reader<&[u8]>, local: &[u8]) -> CoreResult<()> {
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Start(_) => depth += 1,
            Event::End(ref e) => {
                depth -= 1;
                if depth == 0 && e.local_name().as_ref() == local {
                    return Ok(());
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

// ---------------------------------------------------------------------------
// ConnectorShape (p:cxnSp)
// ---------------------------------------------------------------------------

fn parse_connector(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut position = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvCxnSpPr" => {
                    let props = parse_nv_id_name(reader, b"nvCxnSpPr")?;
                    id = props.0;
                    name = props.1;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader, b"spPr")?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"cxnSp" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Shape::Connector(ConnectorShape { id, name, position }))
}

// ---------------------------------------------------------------------------
// Non-visual property parsing helpers
// ---------------------------------------------------------------------------

/// Parse `p:nvSpPr` → (id, name, alt_text, placeholder)
///
/// Structure:
/// ```xml
/// <p:nvSpPr>
///   <p:cNvPr id="4" name="Title 1" descr="Alt text"/>
///   <p:cNvSpPr/>
///   <p:nvPr><p:ph type="title"/></p:nvPr>
/// </p:nvSpPr>
/// ```
fn parse_nv_common_props(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<(u32, String, Option<String>, Option<PlaceholderInfo>, Option<HyperlinkInfo>)> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut placeholder = None;
    let mut hyperlink = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                {
                    match e.local_name().as_ref() {
                        b"cNvPr" => {
                            id = xml::optional_attr_str(e, b"id")?
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0);
                            name = xml::optional_attr_str(e, b"name")?
                                .map(|v| v.into_owned())
                                .unwrap_or_default();
                            alt_text = xml::optional_attr_str(e, b"descr")?.map(|v| v.into_owned());
                            hyperlink = parse_cnvpr_hyperlink(reader, rels)?;
                        },
                        // p:nvPr contains p:ph — don't skip, keep parsing
                        b"nvPr" => {},
                        _ => {
                            xml::skip_element_fast(reader)?;
                        },
                    }
                }
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"cNvPr" => {
                    id = xml::optional_attr_str(e, b"id")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    name = xml::optional_attr_str(e, b"name")?
                        .map(|v| v.into_owned())
                        .unwrap_or_default();
                    alt_text = xml::optional_attr_str(e, b"descr")?.map(|v| v.into_owned());
                },
                b"ph" => {
                    placeholder = Some(PlaceholderInfo {
                        ph_type: xml::optional_attr_str(e, b"type")?.map(|v| v.into_owned()),
                        idx: xml::optional_attr_str(e, b"idx")?.and_then(|v| v.parse().ok()),
                    });
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"nvSpPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok((id, name, alt_text, placeholder, hyperlink))
}

/// Parse `p:cNvPr`'s children (`a:hlinkClick`/`a:hlinkHover`) after the
/// caller already consumed the `cNvPr` Start event and read its own
/// attributes. Reads through the matching `</p:cNvPr>`. `a:hlinkClick`
/// wins when both are present — a hover-only action with no click
/// target is unusual and click is the primary action.
///
/// This is the shape's own click action (Action Buttons, "jump to
/// slide" navigation icons) — a separate mechanism from the run-level
/// `a:rPr/a:hlinkClick` hyperlink `parse_run_properties` already
/// handles. Both `p:nvSpPr` (AutoShape) and `p:nvPicPr` (PictureShape)
/// used to skip this subtree entirely, so a shape whose only purpose
/// was its click action (typical for Action Buttons, which are drawn
/// as icons with no text) vanished from the IR completely.
fn parse_cnvpr_hyperlink(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<Option<HyperlinkInfo>> {
    let mut hover: Option<HyperlinkInfo> = None;
    let mut click: Option<HyperlinkInfo> = None;
    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"hlinkClick" => {
                click = parse_hlink_click(e, rels)?;
                // A non-empty `<a:hlinkClick>...</a:hlinkClick>` can carry
                // an `<a:snd>` child (Action Button sound); its content
                // has no IR representation, so skip it.
                xml::skip_element_fast(reader)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"hlinkClick" => {
                click = parse_hlink_click(e, rels)?;
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"hlinkHover" => {
                hover = parse_hlink_click(e, rels)?;
                xml::skip_element_fast(reader)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"hlinkHover" => {
                hover = parse_hlink_click(e, rels)?;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"cNvPr" => break,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(click.or(hover))
}

/// Parse `p:nvPicPr` → (id, name, alt_text)
fn parse_nv_pic_props(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<(u32, String, Option<String>, Option<HyperlinkInfo>)> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut hyperlink = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"cNvPr" => {
                id = xml::optional_attr_str(e, b"id")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                name = xml::optional_attr_str(e, b"name")?
                    .map(|v| v.into_owned())
                    .unwrap_or_default();
                alt_text = xml::optional_attr_str(e, b"descr")?.map(|v| v.into_owned());
                hyperlink = parse_cnvpr_hyperlink(reader, rels)?;
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"cNvPr" => {
                id = xml::optional_attr_str(e, b"id")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                name = xml::optional_attr_str(e, b"name")?
                    .map(|v| v.into_owned())
                    .unwrap_or_default();
                alt_text = xml::optional_attr_str(e, b"descr")?.map(|v| v.into_owned());
            },
            Event::End(ref e) if e.local_name().as_ref() == b"nvPicPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok((id, name, alt_text, hyperlink))
}

/// Parse a non-visual-properties wrapper (`p:nvGrpSpPr`, `p:nvGraphicFramePr`,
/// `p:nvCxnSpPr`) → (id, name) from its `p:cNvPr` child. `end_tag` is the
/// wrapper's local name, which is the only thing that differs between them.
fn parse_nv_id_name(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_tag: &[u8],
) -> CoreResult<(u32, String)> {
    let mut id = 0u32;
    let mut name = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"cNvPr" => {
                id = xml::optional_attr_str(e, b"id")?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0);
                name = xml::optional_attr_str(e, b"name")?
                    .map(|v| v.into_owned())
                    .unwrap_or_default();
            },
            Event::End(ref e) if e.local_name().as_ref() == end_tag => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok((id, name))
}

// ---------------------------------------------------------------------------
// Shape properties (a:xfrm within p:spPr or p:grpSpPr)
// ---------------------------------------------------------------------------

/// Parse `p:spPr` or `p:grpSpPr` → extract position from `a:xfrm`.
/// `end_tag` is the wrapper's local name.
fn parse_shape_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
    end_tag: &[u8],
) -> CoreResult<Option<ShapePosition>> {
    let mut position = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"xfrm" => {
                position = Some(parse_xfrm_contents(reader)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == end_tag => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(position)
}

/// Parse `p:xfrm` (used in graphicFrame) → extract position.
fn parse_xfrm(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<Option<ShapePosition>> {
    Ok(Some(parse_xfrm_contents(reader)?))
}

/// Parse the contents of an `a:xfrm` or `p:xfrm` element: `<a:off x y/>`, `<a:ext cx cy/>`.
fn parse_xfrm_contents(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<ShapePosition> {
    let mut x = 0i64;
    let mut y = 0i64;
    let mut cx = 0i64;
    let mut cy = 0i64;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => match e.local_name().as_ref() {
                b"off" => {
                    x = xml::optional_attr_str(e, b"x")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    y = xml::optional_attr_str(e, b"y")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                },
                b"ext" => {
                    cx = xml::optional_attr_str(e, b"cx")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    cy = xml::optional_attr_str(e, b"cy")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                },
                _ => {},
            },
            Event::End(_) => {
                // End of xfrm
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(ShapePosition { x, y, cx, cy })
}

// ---------------------------------------------------------------------------
// Text body parsing (DrawingML a: namespace)
// ---------------------------------------------------------------------------

/// Parse `<p:txBody>` or `<a:txBody>`.
fn parse_text_body(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<TextBody> {
    let mut paragraphs = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"p" => {
                    paragraphs.push(parse_text_paragraph(reader, rels)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(_) => {
                // End of txBody
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TextBody { paragraphs })
}

/// Parse `<a:p>`.
fn parse_text_paragraph(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<TextParagraph> {
    use crate::ir::ParagraphAlignment;
    let mut level = 0u32;
    let mut alignment: Option<ParagraphAlignment> = None;
    let mut space_before_hundredths_pt: Option<u32> = None;
    let mut bullet: Option<BulletStyle> = None;
    let mut content = Vec::new();

    let parse_algn = |e: &quick_xml::events::BytesStart| -> CoreResult<Option<ParagraphAlignment>> {
        Ok(xml::optional_attr_str(e, b"algn")?.and_then(|v| match v.as_ref() {
            "l" => Some(ParagraphAlignment::Left),
            "ctr" => Some(ParagraphAlignment::Center),
            "r" => Some(ParagraphAlignment::Right),
            "just" | "justLow" => Some(ParagraphAlignment::Justify),
            "dist" | "thaiDist" => Some(ParagraphAlignment::Distribute),
            _ => None,
        }))
    };

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"pPr" => {
                    level = xml::optional_attr_str(e, b"lvl")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    alignment = parse_algn(e)?;
                    // <a:pPr> with body — scan for <a:spcBef><a:spcPts/> and
                    // the bullet declaration.
                    let depth_start = 1i32;
                    let mut depth = depth_start;
                    let mut in_spc_bef = false;
                    loop {
                        match reader.read_event()? {
                            Event::Start(ref ee) => {
                                depth += 1;
                                if ee.local_name().as_ref() == b"spcBef" {
                                    in_spc_bef = true;
                                }
                                if let Some(b) = parse_bullet(ee)? {
                                    bullet = Some(b);
                                }
                            },
                            Event::Empty(ref ee) => {
                                if in_spc_bef && ee.local_name().as_ref() == b"spcPts" {
                                    if let Some(v) = xml::optional_attr_str(ee, b"val")? {
                                        if let Ok(n) = v.parse::<u32>() {
                                            space_before_hundredths_pt = Some(n);
                                        }
                                    }
                                }
                                if let Some(b) = parse_bullet(ee)? {
                                    bullet = Some(b);
                                }
                            },
                            Event::End(ref ee) => {
                                depth -= 1;
                                if ee.local_name().as_ref() == b"spcBef" {
                                    in_spc_bef = false;
                                }
                                if depth <= 0 && ee.local_name().as_ref() == b"pPr" {
                                    break;
                                }
                            },
                            Event::Eof => break,
                            _ => {},
                        }
                    }
                },
                b"r" => {
                    content.push(TextContent::Run(parse_text_run(reader, rels)?));
                },
                b"br" => {
                    content.push(TextContent::LineBreak);
                    xml::skip_element_fast(reader)?;
                },
                b"fld" => {
                    content.push(TextContent::Field(parse_text_field(reader, e)?));
                },
                // `<a14:m>` wraps an OMML equation (`<m:oMath>`/`<m:oMathPara>`)
                // as a markup-compatibility extension; there's no structural
                // math model, so pull out every `<m:t>` run so the equation's
                // text isn't silently dropped (PPTX analogue of
                // the DOCX OMML fix).
                b"m" => {
                    let text = collect_a_t_text(reader, b"m")?.concat();
                    if !text.is_empty() {
                        content.push(TextContent::Run(TextRun { text, ..Default::default() }));
                    }
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) => match e.local_name().as_ref() {
                b"pPr" => {
                    level = xml::optional_attr_str(e, b"lvl")?
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    alignment = parse_algn(e)?;
                },
                b"br" => {
                    content.push(TextContent::LineBreak);
                },
                _ => {},
            },
            Event::End(ref e) if e.local_name().as_ref() == b"p" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TextParagraph {
        bullet,
        level,
        alignment,
        space_before_hundredths_pt,
        content,
    })
}

/// Parse `<a:r>` text run.
fn parse_text_run(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<TextRun> {
    let mut text = String::new();
    let mut props = RunProps::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"rPr" => {
                    props = parse_run_properties(reader, e, rels)?;
                },
                b"t" => {
                    text = xml::read_text_content_fast(reader)?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"rPr" => {
                props = run_props_from_attrs(e)?;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"r" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TextRun {
        text,
        bold: props.bold,
        italic: props.italic,
        strikethrough: props.strikethrough,
        hyperlink: props.hyperlink,
        font_size_hundredths_pt: props.font_size_hundredths_pt,
        color_rgb: props.color_rgb,
        underline: props.underline,
        font_name: props.font_name,
        baseline: props.baseline,
        caps: props.caps,
        char_spacing_hundredths_pt: props.char_spacing_hundredths_pt,
    })
}

/// Parse run properties from an `<a:rPr>` Start element (has children like hlinkClick).
fn parse_run_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    rels: &Relationships,
) -> CoreResult<RunProps> {
    let mut props = run_props_from_attrs(start)?;
    let mut hyperlink = None;
    let mut color_rgb: Option<[u8; 3]> = None;
    let mut font_name: Option<String> = None;
    // Track whether we are inside `<a:solidFill>` so we only pick up
    // the inner `<a:srgbClr>` (the fill colour proper) and not
    // unrelated `<a:srgbClr>` elements that may appear in sibling
    // effects (e.g. `<a:hl><a:srgbClr/>` for hyperlink colour).
    let mut in_solid_fill = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                if e.local_name().as_ref() == b"solidFill" {
                    in_solid_fill = true;
                } else if e.local_name().as_ref() == b"hlinkClick" {
                    hyperlink = parse_hlink_click(e, rels)?;
                }
            },
            Event::Empty(ref e) => {
                if e.local_name().as_ref() == b"hlinkClick" {
                    hyperlink = parse_hlink_click(e, rels)?;
                } else if e.local_name().as_ref() == b"latin" && font_name.is_none() {
                    font_name = xml::optional_attr_str(e, b"typeface")?.map(|v| v.into_owned());
                } else if in_solid_fill
                    && e.local_name().as_ref() == b"srgbClr"
                    && color_rgb.is_none()
                {
                    color_rgb = parse_srgb_clr(e);
                }
            },
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"solidFill" {
                    in_solid_fill = false;
                } else if e.local_name().as_ref() == b"rPr" {
                    break;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    props.hyperlink = hyperlink;
    props.color_rgb = color_rgb;
    props.font_name = font_name;
    Ok(props)
}

/// Decode a 6-hex-digit `val="RRGGBB"` attribute from `<a:srgbClr/>`
/// to a `[u8; 3]`. Returns `None` when the attribute is absent or
/// malformed.
fn parse_srgb_clr(e: &quick_xml::events::BytesStart) -> Option<[u8; 3]> {
    let val = xml::optional_attr_str(e, b"val").ok().flatten()?;
    let s = val.as_ref();
    if s.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some([r, g, b])
}

/// Parse a non-negative integer DrawingML attribute (e.g. `sz="1800"`).
/// Returns `None` if the attribute is absent or not parseable.
fn parse_u32_attr(e: &quick_xml::events::BytesStart, key: &[u8]) -> CoreResult<Option<u32>> {
    Ok(xml::optional_attr_str(e, key)?.and_then(|v| v.parse::<u32>().ok()))
}

/// Parse a DrawingML boolean attribute: `b="1"` → Some(true), `b="0"` → Some(false), absent → None.
fn parse_bool_attr(e: &quick_xml::events::BytesStart, key: &[u8]) -> CoreResult<Option<bool>> {
    Ok(xml::optional_attr_str(e, key)?.map(|v| v.as_ref() != "0"))
}

/// Parse `<a:hlinkClick r:id="rId1" tooltip="..."/>` into a HyperlinkInfo.
fn parse_hlink_click(
    e: &quick_xml::events::BytesStart,
    rels: &Relationships,
) -> CoreResult<Option<HyperlinkInfo>> {
    // A *present but empty* `r:id=""` is real, valid PowerPoint output —
    // an Action Button whose only "target" is its own action attribute
    // (e.g. `action="ppaction://noaction"` on a sound-effect button)
    // still writes an empty r:id. Treating it the same as a real,
    // unresolvable id used to give up entirely instead of falling
    // through to `action`, losing the shape completely.
    let r_id = xml::optional_attr_str(e, b"r:id")?.filter(|v| !v.is_empty());
    let tooltip = xml::optional_attr_str(e, b"tooltip")?.map(|v| v.into_owned());
    let action = xml::optional_attr_str(e, b"action")?;

    let target = if let Some(ref r_id) = r_id {
        if let Some(rel) = rels.get_by_id(r_id) {
            if rel.target_mode == TargetMode::External {
                HyperlinkTarget::External(rel.target.clone())
            } else {
                HyperlinkTarget::Internal(rel.target.clone())
            }
        } else {
            return Ok(None);
        }
    } else if let Some(ref action) = action {
        // Internal action like ppaction://hlinksldjump
        HyperlinkTarget::Internal(action.to_string())
    } else {
        return Ok(None);
    };

    Ok(Some(HyperlinkInfo { target, tooltip }))
}

/// Parse `<a:fld type="..." ...>` field element.
fn parse_text_field(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
) -> CoreResult<TextField> {
    let field_type = xml::optional_attr_str(start, b"type")?.map(|v| v.into_owned());
    let mut text = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"t" => {
                text = xml::read_text_content_fast(reader)?;
            },
            Event::End(ref e) if e.local_name().as_ref() == b"fld" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TextField { field_type, text })
}

// ---------------------------------------------------------------------------
// Table parsing (DrawingML a: namespace)
// ---------------------------------------------------------------------------

/// Parse `<a:tbl>`.
fn parse_table(reader: &mut quick_xml::Reader<&[u8]>, rels: &Relationships) -> CoreResult<Table> {
    let mut rows = Vec::new();
    let mut first_row_header = false;
    let mut last_row_header = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"tblPr" => {
                // `firstRow` is how a DrawingML table declares a header row.
                // Assuming row 0 is always a header labelled data rows as
                // headers in every table that does not have one.
                first_row_header = xml::optional_attr_str(e, b"firstRow")?
                    .is_some_and(|v| v.as_ref() == "1" || v.as_ref() == "true");
                last_row_header = xml::optional_attr_str(e, b"lastRow")?
                    .is_some_and(|v| v.as_ref() == "1" || v.as_ref() == "true");
                if matches!(reader.read_event()?, Event::Eof) {
                    break;
                }
            },
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"tr" => {
                    rows.push(parse_table_row(reader, rels)?);
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tbl" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(Table {
        rows,
        first_row_header,
        last_row_header,
    })
}

/// Parse `<a:buNone/>`, `<a:buChar char="•"/>` or
/// `<a:buAutoNum type="…" startAt="…"/>`.
fn parse_bullet(e: &quick_xml::events::BytesStart) -> CoreResult<Option<BulletStyle>> {
    Ok(match e.local_name().as_ref() {
        b"buNone" => Some(BulletStyle::None),
        b"buChar" => Some(BulletStyle::Char(
            xml::optional_attr_str(e, b"char")?
                .map(|v| v.into_owned())
                .unwrap_or_else(|| "\u{2022}".to_string()),
        )),
        b"buAutoNum" => Some(BulletStyle::AutoNum {
            scheme: xml::optional_attr_str(e, b"type")?
                .map(|v| v.into_owned())
                .unwrap_or_else(|| "arabicPeriod".to_string()),
            start_at: xml::optional_attr_str(e, b"startAt")?.and_then(|v| v.parse().ok()),
        }),
        _ => None,
    })
}

/// Parse `<a:tr>`.
fn parse_table_row(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
) -> CoreResult<TableRow> {
    let mut cells = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"tc" => {
                cells.push(parse_table_cell(reader, e, rels)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TableRow { cells })
}

/// Parse `<a:tc>`.
fn parse_table_cell(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    rels: &Relationships,
) -> CoreResult<TableCell> {
    // Spans are clamped at parse time. Unbounded, they reach the DOCX
    // writer's grid loop through the IR and make it iterate
    // gridSpan * rowSpan times, so a hostile deck could hang save_as
    // indefinitely without allocating anything.
    const MAX_SPAN: u32 = 1_000;
    let grid_span: u32 = xml::optional_attr_str(start, b"gridSpan")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(1)
        .clamp(1, MAX_SPAN);
    let row_span: u32 = xml::optional_attr_str(start, b"rowSpan")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(1)
        .clamp(1, MAX_SPAN);
    let h_merge = xml::optional_attr_str(start, b"hMerge")?
        .is_some_and(|v| v.as_ref() == "1" || v.as_ref() == "true");
    let v_merge = xml::optional_attr_str(start, b"vMerge")?
        .is_some_and(|v| v.as_ref() == "1" || v.as_ref() == "true");

    let mut text_body = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"txBody" => {
                text_body = Some(parse_text_body(reader, rels)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"tc" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(TableCell {
        text_body,
        grid_span,
        row_span,
        h_merge,
        v_merge,
    })
}

// ---------------------------------------------------------------------------
// Notes text extraction (used by lib.rs)
// ---------------------------------------------------------------------------

/// Parse a slide comments part.
///
/// Handles both shapes PowerPoint writes: the legacy
/// `<p:cmLst><p:cm authorId="…"><p:text>…` and the modern
/// `<p188:cmLst><p188:cm><p188:txBody><a:p><a:r><a:t>…`. Author names live
/// in a separate `commentAuthors` part, so only ids present in the same
/// file resolve; the text is what matters.
pub(crate) fn parse_comments(xml_data: &[u8]) -> Vec<SlideComment> {
    let mut reader = make_content_reader(xml_data);
    let mut out = Vec::new();
    let mut current = String::new();
    let mut depth_in_comment = 0i32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"cm" => {
                    depth_in_comment = 1;
                    current.clear();
                },
                b"text" | b"t" if depth_in_comment > 0 => {
                    if let Ok(t) = xml::read_text_content_fast(&mut reader) {
                        if !t.trim().is_empty() {
                            if !current.is_empty() {
                                current.push(' ');
                            }
                            current.push_str(t.trim());
                        }
                    }
                },
                _ => {},
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cm" => {
                depth_in_comment = 0;
                if !current.is_empty() {
                    out.push(SlideComment {
                        author: None,
                        text: std::mem::take(&mut current),
                    });
                }
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {},
        }
    }
    out
}

/// Extract the speaker notes body from a notes slide XML. Finds the
/// body placeholder (`type="body"`) and returns its structured
/// `TextBody` — the same model ordinary slide body text uses, so a
/// caller converting it (see `convert_text_body` in `convert_pptx.rs`)
/// gets the same bold/italic/bullet/numbering fidelity for free.
pub(crate) fn extract_notes_body(xml_data: &[u8]) -> Option<TextBody> {
    let rels = Relationships::empty();
    let mut reader = make_content_reader(xml_data);
    let mut shapes = Vec::new();

    // Parse the notes slide's shape tree
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spTree" => {
                shapes = parse_shape_tree(
                    &mut reader,
                    &rels,
                    &std::collections::HashMap::new(),
                    &std::collections::HashMap::new(),
                )
                .ok()?;
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }

    // Find the body placeholder and return its text body.
    for shape in &shapes {
        if let Shape::AutoShape(auto) = shape {
            if let Some(ref ph) = auto.placeholder {
                if ph.ph_type.as_deref() == Some("body") {
                    if let Some(ref tb) = auto.text_body {
                        if !extract_plain_text_from_body(tb).is_empty() {
                            return Some(tb.clone());
                        }
                    }
                }
            }
        }
    }

    None
}

/// Extract plain text from a TextBody.
pub(crate) fn extract_plain_text_from_body(body: &TextBody) -> String {
    let mut parts = Vec::new();
    for para in &body.paragraphs {
        let mut para_text = String::new();
        for content in &para.content {
            match content {
                TextContent::Run(run) => para_text.push_str(&run.text),
                TextContent::LineBreak => para_text.push('\n'),
                TextContent::Field(field) => para_text.push_str(&field.text),
            }
        }
        parts.push(para_text);
    }
    parts.join("\n")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn make_slide_xml(body: &str) -> Vec<u8> {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      {body}
    </p:spTree>
  </p:cSld>
</p:sld>"#
        )
        .into_bytes()
    }

    #[test]
    fn test_parse_auto_shape_with_text() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="Title 1" descr="Alt text"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="274638"/>
      <a:ext cx="8229600" cy="1143000"/>
    </a:xfrm>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r><a:t>Hello World</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, "Slide1".to_string(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new())
                .unwrap();

        assert_eq!(slide.shapes.len(), 1);
        if let Shape::AutoShape(ref auto) = slide.shapes[0] {
            assert_eq!(auto.id, 4);
            assert_eq!(auto.name, "Title 1");
            assert_eq!(auto.alt_text.as_deref(), Some("Alt text"));
            assert!(auto.placeholder.is_some());
            assert_eq!(auto.placeholder.as_ref().unwrap().ph_type.as_deref(), Some("title"));
            let pos = auto.position.as_ref().unwrap();
            assert_eq!(pos.x, 457200);
            assert_eq!(pos.y, 274638);
            assert_eq!(pos.cx, 8229600);
            assert_eq!(pos.cy, 1143000);
            let tb = auto.text_body.as_ref().unwrap();
            assert_eq!(tb.paragraphs.len(), 1);
            assert_eq!(tb.paragraphs[0].content.len(), 1);
            if let TextContent::Run(ref run) = tb.paragraphs[0].content[0] {
                assert_eq!(run.text, "Hello World");
            } else {
                panic!("expected text run");
            }
        } else {
            panic!("expected auto shape");
        }
    }

    /// A shape's own click action (`p:cNvPr > a:hlinkClick`)
    /// used to be discarded entirely (the cNvPr Start branch called
    /// `skip_element_fast`); an Action Button (drawn as an icon with no
    /// text, whose entire purpose is the click target) vanished from the
    /// IR completely. Mirrors the real corpus shape poi_51187.pptx: no
    /// text, an `action="ppaction://hlinksldjump"` jump resolved via
    /// `r:id`.
    #[test]
    fn test_parse_auto_shape_shape_level_hyperlink() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="5" name="Icon 1" descr="SRS_Globe_lr2">
      <a:hlinkClick r:id="rId3" action="ppaction://hlinksldjump"/>
    </p:cNvPr>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
</p:sp>"#,
        );
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slide2.xml"/>
</Relationships>"#;
        let rels = Relationships::parse(rels_xml).unwrap();
        let slide =
            Slide::parse(&xml, "Slide1".to_string(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new())
                .unwrap();

        assert_eq!(slide.shapes.len(), 1);
        let Shape::AutoShape(ref auto) = slide.shapes[0] else {
            panic!("expected auto shape");
        };
        assert_eq!(auto.alt_text.as_deref(), Some("SRS_Globe_lr2"));
        let hl = auto.hyperlink.as_ref().expect("hyperlink must be captured");
        match &hl.target {
            HyperlinkTarget::Internal(target) => assert_eq!(target, "slide2.xml"),
            other => panic!("expected an Internal target, got {other:?}"),
        }
    }

    /// A `<a:hlinkClick>` with a non-empty body (e.g. an `<a:snd>` child
    /// for an Action Button's click sound, as in
    /// aspose-slides_HyperlinkSound.pptx) must still be captured and must
    /// not desync the reader — the child content itself has no IR
    /// representation and is simply skipped.
    #[test]
    fn test_parse_auto_shape_hlink_click_with_child_element() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="6" name="Action Button">
      <a:hlinkClick r:id="rId4" action="ppaction://noaction">
        <a:snd r:embed="rId5" name="push.wav"/>
      </a:hlinkClick>
    </p:cNvPr>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
</p:sp>"#,
        );
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    Target="slide3.xml"/>
</Relationships>"#;
        let rels = Relationships::parse(rels_xml).unwrap();
        let slide =
            Slide::parse(&xml, "Slide1".to_string(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new())
                .unwrap();

        assert_eq!(slide.shapes.len(), 1);
        let Shape::AutoShape(ref auto) = slide.shapes[0] else {
            panic!("expected auto shape");
        };
        assert!(
            auto.hyperlink.is_some(),
            "hlinkClick with a child element must still be captured"
        );
    }

    /// A *present but empty* `r:id=""` (real PowerPoint
    /// output for an Action Button whose only target is its own
    /// `action` attribute, e.g. `action="ppaction://noaction"`) used to
    /// be treated the same as a genuinely unresolvable id and give up
    /// entirely instead of falling through to `action`, losing the
    /// shape completely. Mirrors the real corpus file
    /// aspose-slides_HyperlinkSound.pptx exactly.
    #[test]
    fn test_parse_hlink_click_empty_r_id_falls_back_to_action() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="7" name="Action Button: Sound">
      <a:hlinkClick r:id="" action="ppaction://noaction">
        <a:snd r:embed="rId2" name="push.wav"/>
      </a:hlinkClick>
    </p:cNvPr>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
</p:sp>"#,
        );
        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, "Slide1".to_string(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new())
                .unwrap();

        assert_eq!(slide.shapes.len(), 1);
        let Shape::AutoShape(ref auto) = slide.shapes[0] else {
            panic!("expected auto shape");
        };
        let hl = auto
            .hyperlink
            .as_ref()
            .expect("an empty r:id must fall back to the action attribute");
        match &hl.target {
            HyperlinkTarget::Internal(action) => assert_eq!(action, "ppaction://noaction"),
            other => panic!("expected an Internal action target, got {other:?}"),
        }
    }

    /// The slide XML only ever holds `<c:chart r:id="…"/>` — a reference,
    /// no text — so `parse_graphic_frame` must resolve it against the
    /// pre-extracted `charts` map (keyed by that same rId) rather than
    /// finding nothing via the generic `<a:t>` scan every other
    /// non-table graphic falls back to. XML shape matches
    /// a real corpus file (docx4j_pptx-chart.pptx) byte-for-byte on the
    /// graphicData/c:chart structure.
    #[test]
    fn test_parse_graphic_frame_resolves_chart_text_from_the_charts_map() {
        let xml = make_slide_xml(
            r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="5" name="Test Chart"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               r:id="rId2"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#,
        );
        let mut charts = std::collections::HashMap::new();
        charts.insert(
            "rId2".to_string(),
            vec!["Title: Dollars per Group".to_string(), "Categories: Group 1, Group 2".to_string()],
        );

        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
            &charts,
        )
        .unwrap();

        match &slide.shapes[0] {
            Shape::GraphicFrame(gf) => match &gf.content {
                GraphicContent::Text(lines) => {
                    assert!(lines.contains(&"Title: Dollars per Group".to_string()), "{lines:?}");
                    assert!(
                        lines.contains(&"Categories: Group 1, Group 2".to_string()),
                        "{lines:?}"
                    );
                },
                other => panic!("expected GraphicContent::Text, got {other:?}"),
            },
            other => panic!("expected a GraphicFrame shape, got {other:?}"),
        }
    }

    /// No relationship in `charts` for the rId (e.g. the chart part failed
    /// to open) must not panic — just fall through to Unknown, same as any
    /// other unresolvable graphic.
    #[test]
    fn test_parse_graphic_frame_chart_with_no_matching_charts_entry_is_unknown() {
        let xml = make_slide_xml(
            r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="5" name="Test Chart"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               r:id="rId2"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#,
        );
        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();
        match &slide.shapes[0] {
            Shape::GraphicFrame(gf) => {
                assert!(matches!(gf.content, GraphicContent::Unknown));
            },
            other => panic!("expected a GraphicFrame shape, got {other:?}"),
        }
    }

    #[test]
    fn test_parse_group_shape() {
        let xml = make_slide_xml(
            r#"<p:grpSp>
  <p:nvGrpSpPr>
    <p:cNvPr id="10" name="Group 1"/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
  </p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="100" y="200"/>
      <a:ext cx="5000" cy="3000"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="11" name="Child 1"/>
      <p:cNvSpPr/>
      <p:nvPr/>
    </p:nvSpPr>
    <p:spPr/>
    <p:txBody>
      <a:bodyPr/>
      <a:p><a:r><a:t>Inside group</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>
</p:grpSp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        assert_eq!(slide.shapes.len(), 1);
        if let Shape::Group(ref grp) = slide.shapes[0] {
            assert_eq!(grp.id, 10);
            assert_eq!(grp.name, "Group 1");
            assert_eq!(grp.children.len(), 1);
            if let Shape::AutoShape(ref child) = grp.children[0] {
                assert_eq!(child.name, "Child 1");
                let tb = child.text_body.as_ref().unwrap();
                if let TextContent::Run(ref run) = tb.paragraphs[0].content[0] {
                    assert_eq!(run.text, "Inside group");
                }
            }
        } else {
            panic!("expected group shape");
        }
    }

    /// An `<m:oMath>` equation reaches the slide text via the
    /// `<mc:AlternateContent><mc:Choice Requires="a14"><p:sp>...<a14:m>`
    /// wrapper real PowerPoint output uses (poi-legacy_stress013.pptx,
    /// slides 5/9/10), mirroring the real file's structure exactly.
    #[test]
    fn test_omml_equation_inside_alternate_content_is_not_dropped() {
        let xml = make_slide_xml(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
  <mc:Choice Requires="a14">
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="34" name="TextBox 33"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:p><a:pPr/><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>𝑥</m:t></m:r><m:r><m:t>+1</m:t></m:r></m:oMath></m:oMathPara></a14:m></a:p>
      </p:txBody>
    </p:sp>
  </mc:Choice>
  <mc:Fallback>
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="34" name="fallback pic"/>
        <p:cNvPicPr/>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId99"/></p:blipFill>
      <p:spPr/>
    </p:pic>
  </mc:Fallback>
</mc:AlternateContent>"#,
        );

        let rels = Relationships::empty();
        let slide = Slide::parse(
            &xml,
            String::new(),
            &rels,
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();

        assert_eq!(slide.shapes.len(), 1, "expected the Choice branch's shape, got {:?}", slide.shapes);
        let Shape::AutoShape(ref shape) = slide.shapes[0] else {
            panic!("expected an AutoShape from mc:Choice, got {:?}", slide.shapes[0]);
        };
        let tb = shape.text_body.as_ref().expect("equation text box has a body");
        let TextContent::Run(ref run) = tb.paragraphs[0].content[0] else {
            panic!("expected a run carrying the equation text");
        };
        assert_eq!(run.text, "𝑥+1");
    }

    /// When `mc:Choice`'s content is entirely made of elements this crate
    /// doesn't recognize (yielding zero shapes), the `mc:Fallback` branch
    /// must still be used instead of losing the shape altogether.
    #[test]
    fn test_alternate_content_falls_back_to_fallback_branch_when_choice_yields_nothing() {
        let xml = make_slide_xml(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="somethingUnknown">
    <unknownVendor:thing xmlns:unknownVendor="urn:example:unknown"/>
  </mc:Choice>
  <mc:Fallback>
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="7" name="Fallback Shape"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Fallback text</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </mc:Fallback>
</mc:AlternateContent>"#,
        );

        let rels = Relationships::empty();
        let slide = Slide::parse(
            &xml,
            String::new(),
            &rels,
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();

        assert_eq!(slide.shapes.len(), 1);
        let Shape::AutoShape(ref shape) = slide.shapes[0] else {
            panic!("expected the fallback AutoShape, got {:?}", slide.shapes[0]);
        };
        assert_eq!(shape.name, "Fallback Shape");
    }

    #[test]
    fn test_parse_table_shape() {
        let xml = make_slide_xml(
            r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="20" name="Table 1"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="0" y="0"/>
    <a:ext cx="9144000" cy="3000000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="3048000"/>
          <a:gridCol w="3048000"/>
        </a:tblGrid>
        <a:tr h="370840">
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:t>A1</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:t>B1</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
        <a:tr h="370840">
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:t>A2</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:p><a:r><a:t>B2</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        assert_eq!(slide.shapes.len(), 1);
        if let Shape::GraphicFrame(ref gf) = slide.shapes[0] {
            assert_eq!(gf.name, "Table 1");
            if let GraphicContent::Table(ref tbl) = gf.content {
                assert_eq!(tbl.rows.len(), 2);
                assert_eq!(tbl.rows[0].cells.len(), 2);
                let cell_text =
                    extract_plain_text_from_body(tbl.rows[0].cells[0].text_body.as_ref().unwrap());
                assert_eq!(cell_text, "A1");
            } else {
                panic!("expected table content");
            }
        } else {
            panic!("expected graphic frame");
        }
    }

    #[test]
    fn test_parse_picture_shape() {
        let xml = make_slide_xml(
            r#"<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="30" name="Picture 1" descr="A photo"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="100" y="200"/>
      <a:ext cx="3000" cy="2000"/>
    </a:xfrm>
  </p:spPr>
</p:pic>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        assert_eq!(slide.shapes.len(), 1);
        if let Shape::Picture(ref pic) = slide.shapes[0] {
            assert_eq!(pic.id, 30);
            assert_eq!(pic.name, "Picture 1");
            assert_eq!(pic.alt_text.as_deref(), Some("A photo"));
            let pos = pic.position.as_ref().unwrap();
            assert_eq!(pos.x, 100);
            assert_eq!(pos.cx, 3000);
        } else {
            panic!("expected picture shape");
        }
    }

    #[test]
    fn test_parse_connector_shape() {
        let xml = make_slide_xml(
            r#"<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="40" name="Connector 1"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="500" y="600"/>
      <a:ext cx="1000" cy="0"/>
    </a:xfrm>
  </p:spPr>
</p:cxnSp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        assert_eq!(slide.shapes.len(), 1);
        if let Shape::Connector(ref cxn) = slide.shapes[0] {
            assert_eq!(cxn.id, 40);
            assert_eq!(cxn.name, "Connector 1");
            let pos = cxn.position.as_ref().unwrap();
            assert_eq!(pos.x, 500);
        } else {
            panic!("expected connector shape");
        }
    }

    #[test]
    fn test_parse_text_formatting() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="5" name="Text 1"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r>
        <a:rPr b="1" i="1" strike="sngStrike"/>
        <a:t>formatted</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        if let Shape::AutoShape(ref auto) = slide.shapes[0] {
            let tb = auto.text_body.as_ref().unwrap();
            if let TextContent::Run(ref run) = tb.paragraphs[0].content[0] {
                assert_eq!(run.bold, Some(true));
                assert_eq!(run.italic, Some(true));
                assert!(run.strikethrough);
                assert_eq!(run.text, "formatted");
            }
        }
    }

    #[test]
    fn test_parse_text_field() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="6" name="Slide Number"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:fld type="slidenum">
        <a:rPr/>
        <a:t>3</a:t>
      </a:fld>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();

        if let Shape::AutoShape(ref auto) = slide.shapes[0] {
            let tb = auto.text_body.as_ref().unwrap();
            if let TextContent::Field(ref field) = tb.paragraphs[0].content[0] {
                assert_eq!(field.field_type.as_deref(), Some("slidenum"));
                assert_eq!(field.text, "3");
            } else {
                panic!("expected field");
            }
        }
    }

    #[test]
    fn test_parse_notes_text() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Slide Image"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="sldImg"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>Speaker notes here</a:t></a:r></a:p>
          <a:p><a:r><a:t>Second line</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>"#;

        let body = extract_notes_body(xml).unwrap();
        assert_eq!(extract_plain_text_from_body(&body), "Speaker notes here\nSecond line");
    }

    // ── New: blip rId extraction, font size, alignment, space_before, bg ─

    #[test]
    fn test_run_carries_font_size_from_sz_attr() {
        // <a:rPr sz="1800"/> means 18 pt — should land on the run as
        // 1800 hundredths-of-a-point.
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr><p:cNvPr id="7" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r>
        <a:rPr sz="1800"/>
        <a:t>sized</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let tb = a.text_body.as_ref().unwrap();
            if let TextContent::Run(ref r) = tb.paragraphs[0].content[0] {
                assert_eq!(r.font_size_hundredths_pt, Some(1800));
            } else {
                panic!("expected run");
            }
        }
    }

    #[test]
    fn test_run_font_size_absent_when_sz_missing() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr><p:cNvPr id="8" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r><a:t>unsized</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let tb = a.text_body.as_ref().unwrap();
            if let TextContent::Run(ref r) = tb.paragraphs[0].content[0] {
                assert!(r.font_size_hundredths_pt.is_none());
            }
        }
    }

    #[test]
    fn test_paragraph_alignment_parsed_from_algn_attr() {
        use crate::ir::ParagraphAlignment;
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr><p:cNvPr id="9" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r><a:t>centered</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let para = &a.text_body.as_ref().unwrap().paragraphs[0];
            assert_eq!(para.alignment, Some(ParagraphAlignment::Center));
        }
    }

    #[test]
    fn test_paragraph_alignment_all_variants() {
        use crate::ir::ParagraphAlignment;
        let cases = [
            ("l", ParagraphAlignment::Left),
            ("ctr", ParagraphAlignment::Center),
            ("r", ParagraphAlignment::Right),
            ("just", ParagraphAlignment::Justify),
            ("dist", ParagraphAlignment::Distribute),
        ];
        for (algn, expected) in cases {
            let xml = make_slide_xml(&format!(
                r#"<p:sp>
  <p:nvSpPr><p:cNvPr id="9" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:pPr algn="{algn}"/>
      <a:r><a:t>x</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#
            ));
            let slide = Slide::parse(
                &xml,
                String::new(),
                &Relationships::empty(),
                &std::collections::HashMap::new(),
                &std::collections::HashMap::new(),
            )
            .unwrap();
            if let Shape::AutoShape(ref a) = slide.shapes[0] {
                let para = &a.text_body.as_ref().unwrap().paragraphs[0];
                assert_eq!(para.alignment, Some(expected), "algn={algn}");
            }
        }
    }

    #[test]
    fn test_paragraph_space_before_parsed_from_spc_bef() {
        let xml = make_slide_xml(
            r#"<p:sp>
  <p:nvSpPr><p:cNvPr id="11" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:pPr>
        <a:spcBef><a:spcPts val="1200"/></a:spcBef>
      </a:pPr>
      <a:r><a:t>spaced</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"#,
        );

        let rels = Relationships::empty();
        let slide =
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new(), &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let para = &a.text_body.as_ref().unwrap().paragraphs[0];
            assert_eq!(para.space_before_hundredths_pt, Some(1200));
        }
    }

    #[test]
    fn test_picture_embed_resolves_via_media_map() {
        // Build a media map keyed by the rId used in the slide xml so
        // parse_picture can resolve the embed → bytes.
        let xml = make_slide_xml(
            r#"<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="33" name="Photo"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId7"/>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
  </p:spPr>
</p:pic>"#,
        );

        let mut media = std::collections::HashMap::new();
        media.insert("rId7".to_string(), (vec![0xDEu8, 0xADu8, 0xBEu8, 0xEFu8], "png".to_string()));

        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &media,
            &std::collections::HashMap::new(),
        )
        .unwrap();
        if let Shape::Picture(ref pic) = slide.shapes[0] {
            assert_eq!(pic.embed_rid.as_deref(), Some("rId7"));
            assert_eq!(pic.data.as_deref(), Some(&[0xDEu8, 0xADu8, 0xBEu8, 0xEFu8][..]));
            assert_eq!(pic.format.as_deref(), Some("png"));
        } else {
            panic!("expected picture");
        }
    }

    #[test]
    fn test_picture_embed_without_media_still_carries_rid() {
        // Empty media map: rId is captured but data/format are None.
        let xml = make_slide_xml(
            r#"<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="34" name="Photo"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId9"/></p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm>
  </p:spPr>
</p:pic>"#,
        );

        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();
        if let Shape::Picture(ref pic) = slide.shapes[0] {
            assert_eq!(pic.embed_rid.as_deref(), Some("rId9"));
            assert!(pic.data.is_none());
            assert!(pic.format.is_none());
        }
    }

    #[test]
    fn test_slide_background_solid_rgb() {
        // <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF8800"/>…
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FF8800"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sld>"#;
        let slide = Slide::parse(
            xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();
        assert_eq!(slide.background_rgb, Some([0xFF, 0x88, 0x00]));
    }

    #[test]
    fn test_slide_no_background_returns_none() {
        let xml = make_slide_xml("");
        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
            &std::collections::HashMap::new(),
        )
        .unwrap();
        assert!(slide.background_rgb.is_none());
    }

    #[test]
    fn test_parse_hex_rgb_valid() {
        assert_eq!(parse_hex_rgb("FF8800"), Some([0xFF, 0x88, 0x00]));
        assert_eq!(parse_hex_rgb("000000"), Some([0, 0, 0]));
        assert_eq!(parse_hex_rgb("ffffff"), Some([0xFF, 0xFF, 0xFF]));
    }

    #[test]
    fn test_parse_hex_rgb_invalid() {
        assert_eq!(parse_hex_rgb("FF88"), None); // too short
        assert_eq!(parse_hex_rgb("ZZZZZZ"), None); // not hex
        assert_eq!(parse_hex_rgb(""), None);
    }

    // ── read_blip_embed_attr ────────────────────────────────────────────

    fn first_start_elem(xml: &[u8]) -> quick_xml::events::BytesStart<'static> {
        let mut reader = xml::make_fast_reader(xml);
        loop {
            match reader.read_event().unwrap() {
                Event::Start(e) | Event::Empty(e) => return e.into_owned(),
                Event::Eof => panic!("no start"),
                _ => {},
            }
        }
    }

    #[test]
    fn test_blip_embed_attr_with_r_prefix() {
        let e = first_start_elem(
            br#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       r:embed="rId5"/>"#,
        );
        let rid = read_blip_embed_attr(&e).unwrap();
        assert_eq!(rid.as_deref(), Some("rId5"));
    }

    #[test]
    fn test_blip_embed_attr_arbitrary_prefix() {
        // Some writers use an unrelated prefix bound to the rels namespace.
        let e = first_start_elem(
            br#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:foo="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       foo:embed="rId99"/>"#,
        );
        let rid = read_blip_embed_attr(&e).unwrap();
        assert_eq!(rid.as_deref(), Some("rId99"));
    }

    #[test]
    fn test_blip_embed_attr_absent() {
        let e = first_start_elem(
            br#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        );
        let rid = read_blip_embed_attr(&e).unwrap();
        assert!(rid.is_none());
    }
}
