use quick_xml::events::Event;

use crate::core::relationships::{Relationships, TargetMode};
use crate::core::xml;

use super::shape::{
    AutoShape, ConnectorShape, GraphicContent, GraphicFrame, GroupShape, HyperlinkInfo,
    HyperlinkTarget, PictureShape, PlaceholderInfo, Shape, ShapePosition, Table, TableCell,
    TableRow, TextBody, TextContent, TextField, TextParagraph, TextRun,
};

type CoreResult<T> = crate::core::Result<T>;

/// Parsed run properties: (bold, italic, strikethrough, hyperlink, font_size_hundredths_pt).
///
/// PPTX `<a:rPr sz="..."/>` carries font size in hundredths of a point
/// (e.g. `sz="1800"` = 18 pt). Carrying it through the parser is what
/// keeps PDF→PPTX→PDF round-trips from defaulting every paragraph to
/// the writer's 12 pt fallback (which inflated 8-page A4 sources to
/// ~30 pages).
type RunProps = (
    Option<bool>,
    Option<bool>,
    bool,
    Option<HyperlinkInfo>,
    Option<u32>,
    Option<[u8; 3]>,
);

/// A parsed PPTX slide.
#[derive(Debug, Clone)]
pub struct Slide {
    /// Slide name from the `<p:cSld name="...">` attribute.
    pub name: String,
    /// All top-level shapes on this slide.
    pub shapes: Vec<Shape>,
    /// Speaker notes text, if a notes slide is present.
    pub notes: Option<String>,
    /// Solid background colour (RGB) extracted from the slide's
    /// `<p:cSld><p:bg><p:bgPr><a:solidFill>` element. Only the solid
    /// case is parsed; gradient / image / theme-reference fills are
    /// dropped silently and surface as `None`.
    pub background_rgb: Option<[u8; 3]>,
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
    ) -> CoreResult<Self> {
        let mut reader = make_content_reader(xml_data);
        let mut shapes = Vec::new();
        let mut background_rgb = None;

        loop {
            match reader.read_event()? {
                Event::Start(ref e) if e.local_name().as_ref() == b"bg" => {
                    background_rgb = parse_slide_bg(&mut reader)?;
                },
                Event::Start(ref e) if e.local_name().as_ref() == b"spTree" => {
                    shapes = parse_shape_tree(&mut reader, rels, media)?;
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
) -> CoreResult<Vec<Shape>> {
    let mut shapes = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"sp" => shapes.push(parse_auto_shape(reader, rels)?),
                b"pic" => shapes.push(parse_picture(reader, media)?),
                b"grpSp" => shapes.push(parse_group_shape(reader, rels, media)?),
                b"graphicFrame" => shapes.push(parse_graphic_frame(reader, rels)?),
                b"cxnSp" => shapes.push(parse_connector(reader)?),
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::End(ref e) if e.local_name().as_ref() == b"spTree" => {
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

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvSpPr" => {
                    let props = parse_nv_common_props(reader)?;
                    id = props.0;
                    name = props.1;
                    alt_text = props.2;
                    placeholder = props.3;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader)?;
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
        text_body,
        placeholder,
    }))
}

// ---------------------------------------------------------------------------
// PictureShape (p:pic)
// ---------------------------------------------------------------------------

fn parse_picture(
    reader: &mut quick_xml::Reader<&[u8]>,
    media: &std::collections::HashMap<String, (Vec<u8>, String)>,
) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut position = None;
    let mut embed_rid: Option<String> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvPicPr" => {
                    let props = parse_nv_pic_props(reader)?;
                    id = props.0;
                    name = props.1;
                    alt_text = props.2;
                },
                b"blipFill" => {
                    embed_rid = parse_blip_fill_embed(reader)?;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader)?;
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
) -> CoreResult<Shape> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut position = None;
    let mut children = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"nvGrpSpPr" => {
                    let props = parse_nv_grp_props(reader)?;
                    id = props.0;
                    name = props.1;
                },
                b"grpSpPr" => {
                    position = parse_grp_shape_properties(reader)?;
                },
                b"sp" => children.push(parse_auto_shape(reader, rels)?),
                b"pic" => children.push(parse_picture(reader, media)?),
                b"grpSp" => children.push(parse_group_shape(reader, rels, media)?),
                b"graphicFrame" => children.push(parse_graphic_frame(reader, rels)?),
                b"cxnSp" => children.push(parse_connector(reader)?),
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

fn parse_graphic_frame(
    reader: &mut quick_xml::Reader<&[u8]>,
    rels: &Relationships,
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
                        let props = parse_nv_graphic_frame_props(reader)?;
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
                        } else {
                            xml::skip_element_fast(reader)?;
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
                    let props = parse_nv_cxn_props(reader)?;
                    id = props.0;
                    name = props.1;
                },
                b"spPr" => {
                    position = parse_shape_properties(reader)?;
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
) -> CoreResult<(u32, String, Option<String>, Option<PlaceholderInfo>)> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;
    let mut placeholder = None;

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
                            xml::skip_element_fast(reader)?;
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

    Ok((id, name, alt_text, placeholder))
}

/// Parse `p:nvPicPr` → (id, name, alt_text)
fn parse_nv_pic_props(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<(u32, String, Option<String>)> {
    let mut id = 0u32;
    let mut name = String::new();
    let mut alt_text = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"cNvPr" => {
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

    Ok((id, name, alt_text))
}

/// Parse `p:nvGrpSpPr` → (id, name)
fn parse_nv_grp_props(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<(u32, String)> {
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
            Event::End(ref e) if e.local_name().as_ref() == b"nvGrpSpPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok((id, name))
}

/// Parse `p:nvGraphicFramePr` → (id, name)
fn parse_nv_graphic_frame_props(
    reader: &mut quick_xml::Reader<&[u8]>,
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
            Event::End(ref e) if e.local_name().as_ref() == b"nvGraphicFramePr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok((id, name))
}

/// Parse `p:nvCxnSpPr` → (id, name)
fn parse_nv_cxn_props(reader: &mut quick_xml::Reader<&[u8]>) -> CoreResult<(u32, String)> {
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
            Event::End(ref e) if e.local_name().as_ref() == b"nvCxnSpPr" => {
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

/// Parse `p:spPr` → extract position from `a:xfrm`.
fn parse_shape_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<Option<ShapePosition>> {
    let mut position = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"xfrm" => {
                position = Some(parse_xfrm_contents(reader)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"spPr" => {
                break;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(position)
}

/// Parse `p:grpSpPr` → extract position from `a:xfrm`.
fn parse_grp_shape_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
) -> CoreResult<Option<ShapePosition>> {
    let mut position = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) if e.local_name().as_ref() == b"xfrm" => {
                position = Some(parse_xfrm_contents(reader)?);
            },
            Event::End(ref e) if e.local_name().as_ref() == b"grpSpPr" => {
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
                    // <a:pPr> with body — scan for <a:spcBef><a:spcPts/>
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
                            },
                            Event::Empty(ref ee) => {
                                if in_spc_bef && ee.local_name().as_ref() == b"spcPts" {
                                    if let Some(v) = xml::optional_attr_str(ee, b"val")? {
                                        if let Ok(n) = v.parse::<u32>() {
                                            space_before_hundredths_pt = Some(n);
                                        }
                                    }
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
    let mut bold = None;
    let mut italic = None;
    let mut strikethrough = false;
    let mut hyperlink = None;
    let mut font_size_hundredths_pt = None;
    let mut color_rgb: Option<[u8; 3]> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) => match e.local_name().as_ref() {
                b"rPr" => {
                    let props = parse_run_properties(reader, e, rels)?;
                    bold = props.0;
                    italic = props.1;
                    strikethrough = props.2;
                    hyperlink = props.3;
                    font_size_hundredths_pt = props.4;
                    color_rgb = props.5;
                },
                b"t" => {
                    text = xml::read_text_content_fast(reader)?;
                },
                _ => {
                    xml::skip_element_fast(reader)?;
                },
            },
            Event::Empty(ref e) if e.local_name().as_ref() == b"rPr" => {
                let props = parse_run_properties_empty(e, rels)?;
                bold = props.0;
                italic = props.1;
                strikethrough = props.2;
                hyperlink = props.3;
                font_size_hundredths_pt = props.4;
                color_rgb = props.5;
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
        bold,
        italic,
        strikethrough,
        hyperlink,
        font_size_hundredths_pt,
        color_rgb,
    })
}

/// Parse run properties from an `<a:rPr>` Start element (has children like hlinkClick).
fn parse_run_properties(
    reader: &mut quick_xml::Reader<&[u8]>,
    start: &quick_xml::events::BytesStart,
    rels: &Relationships,
) -> CoreResult<RunProps> {
    let bold = parse_bool_attr(start, b"b")?;
    let italic = parse_bool_attr(start, b"i")?;
    let strike = xml::optional_attr_str(start, b"strike")?;
    let strikethrough = strike.as_deref().is_some_and(|v| v != "noStrike");
    let font_size_hundredths_pt = parse_u32_attr(start, b"sz")?;
    let mut hyperlink = None;
    let mut color_rgb: Option<[u8; 3]> = None;
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

    Ok((bold, italic, strikethrough, hyperlink, font_size_hundredths_pt, color_rgb))
}

/// Parse run properties from an `<a:rPr/>` Empty element. Empty
/// elements cannot carry a `<a:solidFill>` child so `color_rgb`
/// is always `None` on this path.
fn parse_run_properties_empty(
    e: &quick_xml::events::BytesStart,
    _rels: &Relationships,
) -> CoreResult<RunProps> {
    let bold = parse_bool_attr(e, b"b")?;
    let italic = parse_bool_attr(e, b"i")?;
    let strike = xml::optional_attr_str(e, b"strike")?;
    let strikethrough = strike.as_deref().is_some_and(|v| v != "noStrike");
    let font_size_hundredths_pt = parse_u32_attr(e, b"sz")?;
    Ok((bold, italic, strikethrough, None, font_size_hundredths_pt, None))
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
    let r_id = xml::optional_attr_str(e, b"r:id")?;
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

    loop {
        match reader.read_event()? {
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

    Ok(Table { rows })
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
    let grid_span: u32 = xml::optional_attr_str(start, b"gridSpan")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
    let row_span: u32 = xml::optional_attr_str(start, b"rowSpan")?
        .and_then(|v| v.parse().ok())
        .unwrap_or(1);
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

/// Extract speaker notes plain text from a notes slide XML.
/// Finds the body placeholder (type="body") and extracts its text.
pub(crate) fn extract_notes_text(xml_data: &[u8]) -> Option<String> {
    let rels = Relationships::empty();
    let mut reader = make_content_reader(xml_data);
    let mut shapes = Vec::new();

    // Parse the notes slide's shape tree
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spTree" => {
                shapes =
                    parse_shape_tree(&mut reader, &rels, &std::collections::HashMap::new()).ok()?;
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {},
        }
    }

    // Find the body placeholder and extract text
    for shape in &shapes {
        if let Shape::AutoShape(auto) = shape {
            if let Some(ref ph) = auto.placeholder {
                if ph.ph_type.as_deref() == Some("body") {
                    if let Some(ref tb) = auto.text_body {
                        let text = extract_plain_text_from_body(tb);
                        if !text.is_empty() {
                            return Some(text);
                        }
                    }
                }
            }
        }
    }

    None
}

/// Extract plain text from a TextBody.
fn extract_plain_text_from_body(body: &TextBody) -> String {
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
    fn parse_auto_shape_with_text() {
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
            Slide::parse(&xml, "Slide1".to_string(), &rels, &std::collections::HashMap::new())
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

    #[test]
    fn parse_group_shape() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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

    #[test]
    fn parse_table_shape() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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
    fn parse_picture_shape() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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
    fn parse_connector_shape() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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
    fn parse_text_formatting() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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
    fn parse_text_field() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();

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
    fn parse_notes_text() {
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

        let text = extract_notes_text(xml).unwrap();
        assert_eq!(text, "Speaker notes here\nSecond line");
    }

    // ── New: blip rId extraction, font size, alignment, space_before, bg ─

    #[test]
    fn run_carries_font_size_from_sz_attr() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();
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
    fn run_font_size_absent_when_sz_missing() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let tb = a.text_body.as_ref().unwrap();
            if let TextContent::Run(ref r) = tb.paragraphs[0].content[0] {
                assert!(r.font_size_hundredths_pt.is_none());
            }
        }
    }

    #[test]
    fn paragraph_alignment_parsed_from_algn_attr() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let para = &a.text_body.as_ref().unwrap().paragraphs[0];
            assert_eq!(para.alignment, Some(ParagraphAlignment::Center));
        }
    }

    #[test]
    fn paragraph_alignment_all_variants() {
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
            )
            .unwrap();
            if let Shape::AutoShape(ref a) = slide.shapes[0] {
                let para = &a.text_body.as_ref().unwrap().paragraphs[0];
                assert_eq!(para.alignment, Some(expected), "algn={algn}");
            }
        }
    }

    #[test]
    fn paragraph_space_before_parsed_from_spc_bef() {
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
            Slide::parse(&xml, String::new(), &rels, &std::collections::HashMap::new()).unwrap();
        if let Shape::AutoShape(ref a) = slide.shapes[0] {
            let para = &a.text_body.as_ref().unwrap().paragraphs[0];
            assert_eq!(para.space_before_hundredths_pt, Some(1200));
        }
    }

    #[test]
    fn picture_embed_resolves_via_media_map() {
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

        let slide = Slide::parse(&xml, String::new(), &Relationships::empty(), &media).unwrap();
        if let Shape::Picture(ref pic) = slide.shapes[0] {
            assert_eq!(pic.embed_rid.as_deref(), Some("rId7"));
            assert_eq!(pic.data.as_deref(), Some(&[0xDEu8, 0xADu8, 0xBEu8, 0xEFu8][..]));
            assert_eq!(pic.format.as_deref(), Some("png"));
        } else {
            panic!("expected picture");
        }
    }

    #[test]
    fn picture_embed_without_media_still_carries_rid() {
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
        )
        .unwrap();
        if let Shape::Picture(ref pic) = slide.shapes[0] {
            assert_eq!(pic.embed_rid.as_deref(), Some("rId9"));
            assert!(pic.data.is_none());
            assert!(pic.format.is_none());
        }
    }

    #[test]
    fn slide_background_solid_rgb() {
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
        )
        .unwrap();
        assert_eq!(slide.background_rgb, Some([0xFF, 0x88, 0x00]));
    }

    #[test]
    fn slide_no_background_returns_none() {
        let xml = make_slide_xml("");
        let slide = Slide::parse(
            &xml,
            String::new(),
            &Relationships::empty(),
            &std::collections::HashMap::new(),
        )
        .unwrap();
        assert!(slide.background_rgb.is_none());
    }

    #[test]
    fn parse_hex_rgb_valid() {
        assert_eq!(parse_hex_rgb("FF8800"), Some([0xFF, 0x88, 0x00]));
        assert_eq!(parse_hex_rgb("000000"), Some([0, 0, 0]));
        assert_eq!(parse_hex_rgb("ffffff"), Some([0xFF, 0xFF, 0xFF]));
    }

    #[test]
    fn parse_hex_rgb_invalid() {
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
    fn blip_embed_attr_with_r_prefix() {
        let e = first_start_elem(
            br#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       r:embed="rId5"/>"#,
        );
        let rid = read_blip_embed_attr(&e).unwrap();
        assert_eq!(rid.as_deref(), Some("rId5"));
    }

    #[test]
    fn blip_embed_attr_arbitrary_prefix() {
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
    fn blip_embed_attr_absent() {
        let e = first_start_elem(
            br#"<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
        );
        let rid = read_blip_embed_attr(&e).unwrap();
        assert!(rid.is_none());
    }
}
