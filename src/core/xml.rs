use std::borrow::Cow;

use quick_xml::NsReader;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, ResolveResult};

use super::error::{Error, Result};

/// OOXML namespace URI constants. Match by URI, never by prefix.
pub mod ns {
    // OPC package namespaces
    /// `[Content_Types].xml` namespace.
    pub const CONTENT_TYPES: &[u8] =
        b"http://schemas.openxmlformats.org/package/2006/content-types";
    /// `.rels` relationships namespace.
    pub const RELATIONSHIPS: &[u8] =
        b"http://schemas.openxmlformats.org/package/2006/relationships";
    /// Core properties namespace.
    pub const CORE_PROPERTIES: &[u8] =
        b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    // Dublin Core
    /// Dublin Core elements namespace.
    pub const DC: &[u8] = b"http://purl.org/dc/elements/1.1/";
    /// Dublin Core terms namespace.
    pub const DC_TERMS: &[u8] = b"http://purl.org/dc/terms/";

    // DrawingML
    /// DrawingML main namespace (`a:` prefix).
    pub const DRAWING_ML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";

    // Format-specific
    /// WordprocessingML namespace (`w:` prefix).
    pub const WML: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    /// SpreadsheetML namespace (`x:` prefix).
    pub const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// PresentationML namespace (`p:` prefix).
    pub const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";

    // Office document relationships (r: prefix in content XML)
    /// Relationships namespace used inline in content XML (`r:` prefix).
    pub const R: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Extended properties
    /// Extended (application) properties namespace.
    pub const EXTENDED_PROPERTIES: &[u8] =
        b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    // String variants for XML writing (same URIs as above, as &str)
    /// `WML` as a `&str` for use in XML writing.
    pub const WML_STR: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    /// `SML` as a `&str` for use in XML writing.
    pub const SML_STR: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// `PML` as a `&str` for use in XML writing.
    pub const PML_STR: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    /// `DRAWING_ML` as a `&str` for use in XML writing.
    pub const DRAWING_ML_STR: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /// `R` as a `&str` for use in XML writing.
    pub const R_STR: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Strict OOXML variants
    /// ISO 29500 Strict variant of `WML`.
    pub const STRICT_WML: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
    /// ISO 29500 Strict variant of `SML`.
    pub const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
    /// ISO 29500 Strict variant of `PML`.
    pub const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
    /// ISO 29500 Strict variant of `DRAWING_ML`.
    pub const STRICT_DRAWING: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
    /// ISO 29500 Strict variant of `R`.
    pub const STRICT_R: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
}

/// Return the Strict namespace variant for a Transitional namespace, if one exists.
/// This enables transparent parsing of both ISO 29500 Strict and ECMA-376 Transitional documents.
fn strict_alternate(ns: &[u8]) -> Option<&'static [u8]> {
    match ns {
        x if x == ns::WML => Some(ns::STRICT_WML),
        x if x == ns::SML => Some(ns::STRICT_SML),
        x if x == ns::PML => Some(ns::STRICT_PML),
        x if x == ns::DRAWING_ML => Some(ns::STRICT_DRAWING),
        x if x == ns::R => Some(ns::STRICT_R),
        _ => None,
    }
}

/// Check if a resolved namespace + local name matches expected values.
/// Also matches the Strict (ISO 29500) variant of the namespace.
pub fn matches_start(resolve: &ResolveResult, start: &BytesStart, ns: &[u8], local: &[u8]) -> bool {
    start.local_name().as_ref() == local
        && match resolve {
            ResolveResult::Bound(Namespace(n)) => {
                *n == ns || strict_alternate(ns).is_some_and(|s| *n == s)
            },
            _ => false,
        }
}

/// Check if a resolved namespace matches, ignoring local name.
/// Also matches the Strict (ISO 29500) variant of the namespace.
pub fn matches_ns(resolve: &ResolveResult, ns: &[u8]) -> bool {
    match resolve {
        ResolveResult::Bound(Namespace(n)) => {
            *n == ns || strict_alternate(ns).is_some_and(|s| *n == s)
        },
        _ => false,
    }
}

/// Get a required attribute value, returning Error::MissingAttribute if absent.
pub fn required_attr<'a>(event: &'a BytesStart, key: &[u8]) -> Result<Cow<'a, [u8]>> {
    match event.try_get_attribute(key)? {
        Some(attr) => Ok(attr.value),
        None => Err(Error::MissingAttribute {
            element: String::from_utf8_lossy(event.local_name().as_ref()).into_owned(),
            attr: String::from_utf8_lossy(key).into_owned(),
        }),
    }
}

/// Get a required attribute as a UTF-8 string.
pub fn required_attr_str<'a>(event: &'a BytesStart, key: &[u8]) -> Result<Cow<'a, str>> {
    let value = required_attr(event, key)?;
    match value {
        Cow::Borrowed(b) => Ok(Cow::Borrowed(std::str::from_utf8(b)?)),
        Cow::Owned(v) => Ok(Cow::Owned(String::from_utf8(v).map_err(|e| e.utf8_error())?)),
    }
}

/// Get an optional attribute value.
pub fn optional_attr<'a>(event: &'a BytesStart, key: &[u8]) -> Result<Option<Cow<'a, [u8]>>> {
    Ok(event.try_get_attribute(key)?.map(|a| a.value))
}

/// Get an optional attribute as a UTF-8 string.
pub fn optional_attr_str<'a>(event: &'a BytesStart, key: &[u8]) -> Result<Option<Cow<'a, str>>> {
    match optional_attr(event, key)? {
        Some(Cow::Borrowed(b)) => Ok(Some(Cow::Borrowed(std::str::from_utf8(b)?))),
        Some(Cow::Owned(v)) => {
            Ok(Some(Cow::Owned(String::from_utf8(v).map_err(|e| e.utf8_error())?)))
        },
        None => Ok(None),
    }
}

/// Get an optional prefixed attribute by local name, trying all namespace prefixes.
/// For example, `optional_prefixed_attr_str(e, b"id")` matches `r:id`, `d3p1:id`, etc.
/// Falls back to unprefixed `id` if no prefixed match is found.
pub fn optional_prefixed_attr_str<'a>(
    event: &'a BytesStart,
    local_name: &[u8],
) -> Result<Option<Cow<'a, str>>> {
    for attr in event.attributes().flatten() {
        let key = attr.key.as_ref();
        // Check prefixed: look for `:localname` at the end
        if let Some(pos) = key.iter().position(|&b| b == b':') {
            if &key[pos + 1..] == local_name {
                let value = attr.unescape_value()?;
                return Ok(Some(value));
            }
        } else if key == local_name {
            let value = attr.unescape_value()?;
            return Ok(Some(value));
        }
    }
    Ok(None)
}

/// Parse an OOXML boolean toggle element.
///
/// Bare element (`<b/>`) = true, `val="0"` / `val="false"` / `val="off"` = false.
/// The `attr_name` is typically `b"w:val"` (WML) or `b"val"` (SML/DrawingML).
pub fn parse_toggle(e: &BytesStart, attr_name: &[u8]) -> bool {
    match optional_attr_str(e, attr_name) {
        Ok(Some(ref val)) => !matches!(val.as_ref(), "0" | "false" | "off"),
        _ => true,
    }
}

/// Read text content between start and end tags, consuming through the matching end tag.
pub fn read_text_content(reader: &mut NsReader<&[u8]>) -> Result<String> {
    use quick_xml::events::Event;
    let mut text = String::new();
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Text(e) => {
                text.push_str(&e.unescape()?);
            },
            Event::CData(e) => {
                text.push_str(std::str::from_utf8(&e)?);
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
    Ok(text)
}

/// Skip over the current element and all its children (consumes through matching end tag).
pub fn skip_element(reader: &mut NsReader<&[u8]>) -> Result<()> {
    use quick_xml::events::Event;
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
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
    Ok(())
}

/// Create an NsReader configured for OOXML parsing.
pub fn make_reader(xml: &[u8]) -> NsReader<&[u8]> {
    let mut reader = NsReader::from_reader(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = false;
    config.check_comments = false;
    reader
}

// ===========================================================================
// Fast Reader utilities (no namespace resolution — for hot-path parsing)
// ===========================================================================

/// Create a plain Reader (no namespace resolution) configured for OOXML parsing.
/// Use this for format-specific hot paths (worksheets, slides, document body)
/// where all elements are in a single known namespace.
pub fn make_fast_reader(xml: &[u8]) -> quick_xml::Reader<&[u8]> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let config = reader.config_mut();
    config.trim_text(true);
    config.check_end_names = false;
    config.check_comments = false;
    reader
}

/// Read text content between start and end tags using fast Reader.
pub fn read_text_content_fast(reader: &mut quick_xml::Reader<&[u8]>) -> Result<String> {
    use quick_xml::events::Event;
    let mut text = String::new();
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Text(e) => {
                text.push_str(&e.unescape()?);
            },
            Event::CData(e) => {
                text.push_str(&String::from_utf8_lossy(&e));
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
    Ok(text)
}

/// Skip over the current element and all its children using fast Reader.
pub fn skip_element_fast(reader: &mut quick_xml::Reader<&[u8]>) -> Result<()> {
    use quick_xml::events::Event;
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
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
    Ok(())
}

/// Transcode XML bytes to UTF-8 if the XML declaration specifies a non-UTF-8 encoding.
/// Returns `None` if the data is already UTF-8 (the common case), or `Some(transcoded)`
/// if transcoding was needed. Callers should use the returned buffer for parsing.
pub fn ensure_utf8(data: &[u8]) -> Option<Vec<u8>> {
    // Quick check: if it's valid UTF-8 already, skip everything
    if std::str::from_utf8(data).is_ok() {
        return None;
    }

    // Look for encoding="..." in the first 200 bytes of the XML declaration
    let header = &data[..data.len().min(200)];
    let header_str = String::from_utf8_lossy(header);

    let encoding_name = if let Some(pos) = header_str.find("encoding=") {
        let rest = &header_str[pos + 9..];
        let quote = rest.as_bytes().first().copied().unwrap_or(b'"');
        if quote == b'"' || quote == b'\'' {
            let inner = &rest[1..];
            inner.split(quote as char).next().unwrap_or("utf-8")
        } else {
            return None;
        }
    } else {
        // No encoding declaration, try ISO-8859-1 as fallback for non-UTF-8
        "iso-8859-1"
    };

    let encoding = encoding_rs::Encoding::for_label(encoding_name.as_bytes())?;
    if encoding == encoding_rs::UTF_8 {
        return None;
    }

    let (result, _, had_errors) = encoding.decode(data);
    if had_errors {
        return None;
    }

    // Replace the encoding declaration with utf-8 so the XML parser doesn't complain
    let mut utf8 = result.into_owned().into_bytes();
    if let Some(pos) = utf8
        .windows(9)
        .position(|w| w.eq_ignore_ascii_case(b"encoding="))
    {
        let rest = &utf8[pos + 9..];
        if let Some(&quote) = rest.first() {
            if quote == b'"' || quote == b'\'' {
                if let Some(end) = rest[1..].iter().position(|&b| b == quote) {
                    let start = pos + 10;
                    let end = start + end;
                    utf8.splice(start..end, b"utf-8".iter().copied());
                }
            }
        }
    }

    Some(utf8)
}
