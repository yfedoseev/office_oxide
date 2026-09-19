use std::borrow::Cow;

use quick_xml::NsReader;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, ResolveResult};

use super::error::{Error, Result};

/// OOXML namespace URI constants. Match by URI, never by prefix.
pub mod ns {
    // OPC package namespaces
    /// `[Content_Types].xml` namespace.
    pub const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
    /// `.rels` relationships namespace.
    pub const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    /// Core properties namespace.
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    // Dublin Core
    /// Dublin Core elements namespace.
    pub const DC: &str = "http://purl.org/dc/elements/1.1/";
    /// Dublin Core terms namespace.
    pub const DC_TERMS: &str = "http://purl.org/dc/terms/";

    // DrawingML
    /// DrawingML main namespace (`a:` prefix).
    pub const DRAWING_ML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Format-specific
    /// WordprocessingML namespace (`w:` prefix).
    pub const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    /// SpreadsheetML namespace (`x:` prefix).
    pub const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// PresentationML namespace (`p:` prefix).
    pub const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    // Office document relationships (r: prefix in content XML)
    /// Relationships namespace used inline in content XML (`r:` prefix).
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Extended properties
    /// Extended (application) properties namespace.
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    // Aliases kept for the XML writers, which predate the constants above
    // being `&str`.
    /// Alias of [`WML`].
    pub const WML_STR: &str = WML;
    /// Alias of [`SML`].
    pub const SML_STR: &str = SML;
    /// Alias of [`PML`].
    pub const PML_STR: &str = PML;
    /// Alias of [`DRAWING_ML`].
    pub const DRAWING_ML_STR: &str = DRAWING_ML;
    /// Alias of [`R`].
    pub const R_STR: &str = R;

    // Strict OOXML variants
    /// ISO 29500 Strict variant of `WML`.
    pub const STRICT_WML: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    /// ISO 29500 Strict variant of `SML`.
    pub const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    /// ISO 29500 Strict variant of `PML`.
    pub const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    /// ISO 29500 Strict variant of `DRAWING_ML`.
    pub const STRICT_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    /// ISO 29500 Strict variant of `R`.
    pub const STRICT_R: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
}

/// Return the Strict namespace variant for a Transitional namespace, if one exists.
/// This enables transparent parsing of both ISO 29500 Strict and ECMA-376 Transitional documents.
fn strict_alternate(ns: &str) -> Option<&'static str> {
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
pub fn matches_start(resolve: &ResolveResult, start: &BytesStart, ns: &str, local: &str) -> bool {
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
pub fn matches_ns(resolve: &ResolveResult, ns: &str) -> bool {
    match resolve {
        ResolveResult::Bound(Namespace(n)) => {
            *n == ns || strict_alternate(ns).is_some_and(|s| *n == s)
        },
        _ => false,
    }
}

/// Get a required attribute value, returning Error::MissingAttribute if absent.
pub fn required_attr<'a>(event: &'a BytesStart, key: &str) -> Result<Cow<'a, str>> {
    match event.try_get_attribute(key)? {
        Some(attr) => Ok(attr.value),
        None => Err(Error::MissingAttribute {
            element: event.local_name().as_ref().to_string(),
            attr: key.to_string(),
        }),
    }
}

/// Get a required attribute as a UTF-8 string, with XML entity references
/// resolved. See [`optional_attr_str`] for why the unescape matters.
pub fn required_attr_str<'a>(event: &'a BytesStart, key: &str) -> Result<Cow<'a, str>> {
    unescape_cow(required_attr(event, key)?)
}

/// Resolve XML entity references in an attribute value, borrowing when the
/// value contains none (the overwhelmingly common case).
fn unescape_cow(text: Cow<'_, str>) -> Result<Cow<'_, str>> {
    if !text.contains('&') {
        return Ok(text);
    }
    let unescaped = quick_xml::escape::unescape(&text).map_err(quick_xml::Error::from)?;
    Ok(Cow::Owned(unescaped.into_owned()))
}

/// Get an optional attribute value.
pub fn optional_attr<'a>(event: &'a BytesStart, key: &str) -> Result<Option<Cow<'a, str>>> {
    Ok(event.try_get_attribute(key)?.map(|a| a.value))
}

/// Get an optional attribute as a UTF-8 string, with XML entity references
/// resolved.
///
/// The raw bytes quick-xml hands back are still escaped: a `formatCode`
/// written as `#,##0,,&quot; M&quot;` arrives with the six literal
/// characters `&quot;` in place of each `"`. Every consumer that inspects
/// the value then sees text that is not in the document — the number-format
/// scanner read the `M` of `&quot; M&quot;` as a month token and rendered
/// 12,500,000 as the date 36123-11-01 — and every URL, alt text and style
/// name kept its `&amp;` verbatim.
pub fn optional_attr_str<'a>(event: &'a BytesStart, key: &str) -> Result<Option<Cow<'a, str>>> {
    optional_attr(event, key)?.map(unescape_cow).transpose()
}

/// Get an optional prefixed attribute by local name, trying all namespace prefixes.
/// For example, `optional_prefixed_attr_str(e, "id")` matches `r:id`, `d3p1:id`, etc.
/// Falls back to unprefixed `id` if no prefixed match is found.
pub fn optional_prefixed_attr_str<'a>(
    event: &'a BytesStart,
    local_name: &str,
) -> Result<Option<Cow<'a, str>>> {
    for attr in event.attributes().flatten() {
        let key = attr.key.as_ref();
        // Check prefixed: look for `:localname` at the end
        if let Some((_, rest)) = key.split_once(':') {
            if rest == local_name {
                return Ok(Some(Cow::Owned(unescape_attr_value(&attr)?)));
            }
        } else if key == local_name {
            return Ok(Some(Cow::Owned(unescape_attr_value(&attr)?)));
        }
    }
    Ok(None)
}

/// Parse an OOXML boolean toggle element.
///
/// Bare element (`<b/>`) = true, `val="0"` / `val="false"` / `val="off"` = false.
/// The `attr_name` is typically `"w:val"` (WML) or `"val"` (SML/DrawingML).
pub fn parse_toggle(e: &BytesStart, attr_name: &str) -> bool {
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
                text.push_str(&unescape_text(&e)?);
            },
            Event::GeneralRef(e) => {
                text.push_str(&resolve_general_ref(&e)?);
            },
            Event::CData(e) => {
                text.push_str(&e);
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

/// Unescape a `BytesText` event into an owned string.
///
/// quick-xml hands text events over as already-validated UTF-8 (the reader
/// rejects anything else), so only the entity references remain to be
/// resolved. `EscapeError` goes through `quick_xml::Error` to reach our
/// `core::Error`.
pub fn unescape_text(e: &quick_xml::events::BytesText<'_>) -> Result<String> {
    let unescaped = quick_xml::escape::unescape(e).map_err(quick_xml::Error::from)?;
    Ok(unescaped.into_owned())
}

/// Unescape an `Attribute` value into an owned string.
///
/// `Attribute::unescape_value()` is deprecated and, under quick-xml's
/// `encoding` feature, `cfg`-compiled out entirely. Feature unification can
/// turn `encoding` on transitively (e.g. via `calamine`), so relying on it
/// makes the build fragile. Resolving the entity references (`&amp;`,
/// `&lt;`, …) explicitly mirrors `unescape_text` above and is independent
/// of that feature.
///
/// One deliberate difference from `unescape_value()`: that method additionally
/// applied XML attribute-value whitespace normalization (a *literal* tab/CR/LF
/// inside a value collapses to a space), which `escape::unescape` does not do.
/// This never affects real OOXML — attribute values do not contain literal
/// control whitespace, and character references (`&#9;`, `&#10;`) are unescaped
/// identically either way.
pub fn unescape_attr_value(attr: &quick_xml::events::attributes::Attribute<'_>) -> Result<String> {
    let unescaped = quick_xml::escape::unescape(&attr.value).map_err(quick_xml::Error::from)?;
    Ok(unescaped.into_owned())
}

/// Fail when an XML part ends before its root element is closed.
///
/// A parse loop that breaks on `Event::Eof` returns whatever it read, so a
/// `document.xml` cut off mid-element by a failed download or a truncated
/// upload produced a document that looked complete and was not, with no
/// signal at all. Checking that the closing tag is present is cheap — the
/// root close is always the last markup in the part, so only the tail is
/// scanned — and catches exactly that case without a second full parse.
pub fn check_root_closed(data: &[u8], part: &str, root_local: &str) -> Result<()> {
    // The root close is always the last markup in the part, so only the
    // tail is scanned. Small parts are scanned whole.
    const TAIL: usize = 64 * 1024;
    let tail = &data[data.len().saturating_sub(TAIL)..];
    let needle = format!("{root_local}>");
    let needle = needle.as_bytes();

    // Accept `</root>` and `</prefix:root>`: find the local-name-plus-`>`
    // and require a `</` at most one short prefix earlier.
    let found = tail
        .windows(needle.len())
        .enumerate()
        .filter(|(_, w)| *w == needle)
        .any(|(i, _)| {
            let before = &tail[i.saturating_sub(24)..i];
            match before.iter().rposition(|&b| b == b'<') {
                Some(lt) => {
                    let between = &before[lt..];
                    between.starts_with(b"</")
                        && between[2..].iter().all(|&b| b != b'<' && b != b'>')
                },
                None => false,
            }
        });
    if found {
        Ok(())
    } else {
        Err(Error::TruncatedPart(part.to_string()))
    }
}

/// The prefixes bound to an expected namespace by a part's root element.
///
/// Element dispatch throughout this crate matches on *local name* only, so
/// an element from any namespace whose local name happens to match is
/// parsed as if it were the real thing: a `<evil:p><evil:r><evil:t>` inside
/// a `w:body` extracted as ordinary document text that Word never renders.
/// This records which prefixes the root actually bound to the format's
/// namespace so the content parsers can skip everything else.
#[derive(Debug, Clone, Default)]
pub struct NsGuard {
    /// Prefixes bound to the expected namespace. An empty `Vec` with
    /// `permissive` set means "accept everything".
    prefixes: Vec<String>,
    /// Set when the part declared no usable namespace at all, in which case
    /// filtering would reject the whole document. Hand-written and minimal
    /// fixtures do this routinely.
    permissive: bool,
}

impl NsGuard {
    /// A guard that accepts every element. Used where a part's root has not
    /// been inspected.
    pub fn permissive() -> Self {
        Self {
            prefixes: Vec::new(),
            permissive: true,
        }
    }

    /// Build a guard from a part's root start tag.
    ///
    /// `expected` are the namespace URIs that count as the format's own
    /// (Transitional and Strict). Returns `Err` when the root binds its own
    /// prefix to something else entirely — a document claiming to be
    /// WordprocessingML while its `w:` prefix points elsewhere is not the
    /// format it says it is.
    pub fn from_root(root: &BytesStart, expected: &[&str], format: &str) -> Result<Self> {
        let root_prefix = root
            .name()
            .as_ref()
            .split_once(':')
            .map(|(p, _)| p.to_string());

        let mut prefixes = Vec::new();
        let mut root_prefix_bound_elsewhere = false;
        for attr in root.attributes().flatten() {
            let key = attr.key.as_ref();
            let (prefix, is_ns) = if key == "xmlns" {
                (String::new(), true)
            } else if let Some(rest) = key.strip_prefix("xmlns:") {
                (rest.to_string(), true)
            } else {
                (String::new(), false)
            };
            if !is_ns {
                continue;
            }
            if expected.iter().any(|e| *e == attr.value.as_ref()) {
                prefixes.push(prefix);
            } else if root_prefix.as_deref() == Some(prefix.as_str()) {
                root_prefix_bound_elsewhere = true;
            }
        }

        if prefixes.is_empty() {
            if root_prefix_bound_elsewhere {
                return Err(Error::MalformedXml(format!(
                    "root element's namespace is not {format}"
                )));
            }
            // No namespace declaration at all — accept, so minimal
            // hand-written parts keep working.
            return Ok(Self::permissive());
        }
        Ok(Self {
            prefixes,
            permissive: false,
        })
    }

    /// Whether an element belongs to the expected namespace.
    pub fn accepts(&self, e: &BytesStart) -> bool {
        if self.permissive {
            return true;
        }
        let name = e.name();
        let prefix = name.as_ref().split_once(':').map_or("", |(p, _)| p);
        self.prefixes.iter().any(|p| p == prefix)
    }
}

/// Strip characters XML 1.0 forbids from a text value.
///
/// XML 1.0 §2.2 permits only tab, LF, CR and `U+0020..` (minus the
/// surrogate and non-character ranges) — every other C0 control is
/// unrepresentable, *including* as a numeric character reference. Writing
/// one produces a file that Word, Excel and LibreOffice all reject as
/// corrupt, and such characters arrive routinely from PDF text extraction
/// and from database exports. Dropping them is the only lossless-enough
/// option: there is no escape that would round-trip.
pub fn sanitize_xml_text(s: &str) -> std::borrow::Cow<'_, str> {
    fn allowed(c: char) -> bool {
        matches!(c,
            '\u{09}' | '\u{0A}' | '\u{0D}'
            | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
        )
    }
    if s.chars().all(allowed) {
        return std::borrow::Cow::Borrowed(s);
    }
    std::borrow::Cow::Owned(s.chars().filter(|&c| allowed(c)).collect())
}

/// Maximum element-nesting depth accepted by the recursive-descent parsers.
///
/// Without a cap, a small `.docx` holding a few thousand nested `<w:tbl>`
/// elements drives the parser into a stack overflow, which aborts the
/// process — an uncatchable crash no consumer of this library, in any
/// binding, can defend against.
///
/// The value is empirical, and the measurement that matters is the *worst*
/// stack a caller might have, not the best. Nested-table documents built at
/// increasing depths overflow at:
///
/// | build | stack | layer | cliff |
/// |---|---|---|---|
/// | release | 16 MB parse stack | XML parse (`parse_table`) | 3,000-4,000 |
/// | debug | default 2 MiB thread | XML parse (`parse_table`) | 512-1,024 |
/// | debug | default 2 MiB thread | `to_ir()` (`convert_table`) | 150-200 |
///
/// `DepthGuard` bounds the XML-parse recursion, which now runs on its own
/// `PARSE_STACK_SIZE` thread rather than whatever the caller happened to
/// have (see below) — but the *result* is then walked again by `to_ir()`/
/// `plain_text()`/`to_markdown()`, on whatever stack the caller gave *them*,
/// which this crate does not control and is not necessarily large. Those
/// walkers hit their own stack limit well before the XML-parse cliff, since
/// `convert_table`/`plain_text_table` recurse with a much larger frame
/// (multiple local `Vec`s and struct literals per level) than the XML
/// event-loop's `parse_table` does — so they carry their own `DepthGuard`
/// too, and it's the tighter of the two cliffs, not the
/// parse-stack one, that this constant must stay under.
///
/// 100 is chosen with real margin under the 150-200 debug/2 MiB `to_ir()`
/// cliff — still ~20x deeper than any document a human authoring tool
/// produces.
///
/// That the XML-parse cliff has its own large margin only ever held where
/// the parse actually got the stack it was measured against, and it often
/// did not: `needs_stack_thread` inferred the answer from `RLIMIT_STACK`,
/// which describes the main thread rather than the running one, so an
/// unlimited limit ran the parse inline on an ordinary 2 MiB thread. Every
/// threaded platform now parses on a `PARSE_STACK_SIZE` stack, so *that*
/// part of this constant is calibrated against a stack the library owns —
/// the `to_ir()`/`plain_text()` cliff is not, and never can be, since it
/// runs on the caller's own stack.
///
/// Re-measure if the parser or IR-conversion structs grow — this moves with
/// the frame size, and it moved once already: earlier releases' cliffs were
/// measured only against the XML-parse layer and were 5,000-10,000 (release)
/// / 512-1,024 (debug) before this release's added fields (and before the
/// conversion-layer `DepthGuard`s existed at all) pulled the real, tighter
/// limit down to what's measured above.
pub const MAX_NESTING_DEPTH: usize = 100;

thread_local! {
    static NESTING_DEPTH: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static TRUNCATED_SUBTREES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// RAII guard tracking recursion depth in the parsers.
///
/// [`DepthGuard::enter`] returns `None` once [`MAX_NESTING_DEPTH`] is
/// reached; the caller then skips the over-deep subtree instead of
/// recursing into it. The counter is thread-local, so parallel worksheet
/// and slide parsing each get their own budget.
pub struct DepthGuard(());

impl DepthGuard {
    /// Enter one level of nesting, or return `None` when the limit is hit.
    ///
    /// Every `None` also increments the thread-local truncated-subtree
    /// counter (see [`truncated_subtrees`]) — content past the bound used
    /// to be dropped with only a `log::warn!`, leaving a caller with no
    /// way to learn that the document they just wrote or read is missing
    /// content.
    pub fn enter() -> Option<Self> {
        NESTING_DEPTH.with(|d| {
            let cur = d.get();
            if cur >= MAX_NESTING_DEPTH {
                TRUNCATED_SUBTREES.with(|t| t.set(t.get() + 1));
                None
            } else {
                d.set(cur + 1);
                Some(DepthGuard(()))
            }
        })
    }
}

/// Number of subtrees truncated by [`DepthGuard::enter`] returning `None`
/// on this thread since the last [`reset_truncated_subtrees`] call.
pub fn truncated_subtrees() -> usize {
    TRUNCATED_SUBTREES.with(std::cell::Cell::get)
}

/// Reset the truncation counter. Called at the start of a top-level
/// write/parse so its count reflects only that call, not a previous one
/// on the same thread.
pub fn reset_truncated_subtrees() {
    TRUNCATED_SUBTREES.with(|t| t.set(0));
}

impl Drop for DepthGuard {
    fn drop(&mut self) {
        NESTING_DEPTH.with(|d| d.set(d.get().saturating_sub(1)));
    }
}

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

/// Resolve an `Event::GeneralRef` — an `&name;` or `&#NN;` reference — into
/// the text it stands for.
///
/// quick-xml reports every entity reference as its own event rather than
/// folding it into the surrounding `Event::Text`, so a reader that only
/// handles `Event::Text` silently *deletes* them: `AT&amp;T` came out as
/// `ATT` and `&#8212;` vanished. Character references resolve numerically,
/// the five XML predefined entities resolve from the spec, and anything
/// else (a DTD-declared entity we cannot expand) is preserved verbatim as
/// `&name;` so no characters are lost.
pub fn resolve_general_ref(e: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(ch) = e.resolve_char_ref()? {
        return Ok(ch.to_string());
    }
    Ok(match e.as_ref() {
        "lt" => "<".to_string(),
        "gt" => ">".to_string(),
        "amp" => "&".to_string(),
        "apos" => "'".to_string(),
        "quot" => "\"".to_string(),
        other => format!("&{other};"),
    })
}

/// Read text content between start and end tags using fast Reader.
pub fn read_text_content_fast(reader: &mut quick_xml::Reader<&[u8]>) -> Result<String> {
    use quick_xml::events::Event;
    let mut text = String::new();
    let mut depth = 1u32;
    loop {
        match reader.read_event()? {
            Event::Text(e) => {
                text.push_str(&unescape_text(&e)?);
            },
            Event::GeneralRef(e) => {
                text.push_str(&resolve_general_ref(&e)?);
            },
            Event::CData(e) => {
                text.push_str(&e);
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
///
/// Bytes that are not valid in their declared (or fallback) encoding are
/// replaced by U+FFFD rather than handed to the XML reader raw: the reader
/// validates UTF-8 per event and would otherwise fail the whole part on the
/// first bad byte, losing everything around it. The replacement character
/// is visible in the output, so the damage is not silent.
pub fn ensure_utf8(data: &[u8]) -> Option<Vec<u8>> {
    // UTF-16 must be settled before the valid-UTF-8 check below. A UTF-16
    // part whose characters are all ASCII is a run of NUL-interleaved bytes,
    // and NUL is itself valid UTF-8 — so `from_utf8` accepts it, the real
    // encoding is never noticed, and every tag name arrives split by nulls.
    // That silently emptied a UTF-16BE `xl/workbook.xml`.
    if let Some(decoded) = decode_utf16_xml(data) {
        return Some(decoded);
    }

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
            return Some(lossy_utf8(data));
        }
    } else {
        // No encoding declaration, try ISO-8859-1 as fallback for non-UTF-8
        "iso-8859-1"
    };

    let Some(encoding) = encoding_rs::Encoding::for_label(encoding_name.as_bytes()) else {
        return Some(lossy_utf8(data));
    };
    if encoding == encoding_rs::UTF_8 {
        return Some(lossy_utf8(data));
    }

    let (result, _, had_errors) = encoding.decode(data);
    if had_errors {
        return Some(lossy_utf8(data));
    }

    // Replace the encoding declaration with utf-8 so the XML parser doesn't complain
    Some(rewrite_encoding_decl(result.into_owned().into_bytes()))
}

/// Replace every invalid UTF-8 sequence with U+FFFD, warning once per part.
fn lossy_utf8(data: &[u8]) -> Vec<u8> {
    log::warn!("XML part is not valid UTF-8; invalid bytes replaced with U+FFFD");
    String::from_utf8_lossy(data).into_owned().into_bytes()
}

/// Rewrite an XML declaration's `encoding="..."` value to `utf-8`.
///
/// Called after transcoding: leaving the original label in place makes the
/// bytes self-contradictory, and a strict downstream processor would reject
/// them.
fn rewrite_encoding_decl(mut utf8: Vec<u8>) -> Vec<u8> {
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
    utf8
}

/// Decode a UTF-16 XML part to UTF-8, or `None` if `data` isn't UTF-16.
///
/// Endianness comes from a byte-order mark when one is present. Without a
/// BOM, XML 1.0 §4.3.3 requires a UTF-16 document to begin with the text
/// declaration, so the NUL-interleaved `<?` prolog identifies it
/// unambiguously.
fn decode_utf16_xml(data: &[u8]) -> Option<Vec<u8>> {
    let big_endian = match data {
        [0xFE, 0xFF, ..] => true,
        [0xFF, 0xFE, ..] => false,
        // BOM-less: `<?` as UTF-16BE / UTF-16LE code units.
        [0x00, 0x3C, 0x00, 0x3F, ..] => true,
        [0x3C, 0x00, 0x3F, 0x00, ..] => false,
        _ => return None,
    };

    let encoding = if big_endian {
        encoding_rs::UTF_16BE
    } else {
        encoding_rs::UTF_16LE
    };
    // `decode` strips a leading BOM itself, so the result never carries one.
    let (result, _, had_errors) = encoding.decode(data);
    if had_errors {
        return None;
    }

    Some(rewrite_encoding_decl(result.into_owned().into_bytes()))
}

#[cfg(test)]
mod attr_tests {
    use super::unescape_attr_value;
    use quick_xml::events::BytesStart;

    /// Parse `<e {attrs}>` and unescape the value of attribute `key`.
    fn attr_value(attrs: &str, key: &str) -> String {
        let start = BytesStart::from_content(format!("e {attrs}"), 1);
        let attr = start
            .attributes()
            .map(|a| a.unwrap())
            .find(|a| a.key.as_ref() == key)
            .expect("attribute present");
        unescape_attr_value(&attr).unwrap()
    }

    #[test]
    fn test_unescapes_predefined_and_numeric_entities() {
        assert_eq!(attr_value(r#"v="a &amp; b &lt;x&gt; &#65;""#, "v"), "a & b <x> A");
    }

    #[test]
    fn test_passes_plain_value_through_unchanged() {
        assert_eq!(attr_value(r#"r:id="rId7""#, "r:id"), "rId7");
    }

    #[test]
    fn test_unescapes_ampersand_in_hyperlink_target() {
        assert_eq!(
            attr_value(r#"Target="https://x/?a=1&amp;b=2""#, "Target"),
            "https://x/?a=1&b=2"
        );
    }
}

#[cfg(test)]
mod ensure_utf8_tests {
    use super::ensure_utf8;

    #[test]
    fn test_valid_utf8_is_left_alone() {
        assert!(ensure_utf8("<a>caf\u{e9}</a>".as_bytes()).is_none());
    }

    /// A part declared UTF-8 with one invalid byte in a text node must still
    /// parse: the reader validates UTF-8 per event, so handing the raw bytes
    /// over would fail the whole part on that one byte. The bad byte becomes
    /// U+FFFD and everything around it survives.
    #[test]
    fn test_invalid_utf8_under_utf8_declaration_is_replaced_not_dropped() {
        let mut xml = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><a>caf".to_vec();
        xml.push(0xE9); // Latin-1 e-acute, not valid UTF-8
        xml.extend_from_slice(b" ok</a>");
        let out = ensure_utf8(&xml).expect("invalid bytes need transcoding");
        let out = String::from_utf8(out).unwrap();
        assert_eq!(out, "<?xml version=\"1.0\" encoding=\"UTF-8\"?><a>caf\u{FFFD} ok</a>");

        let mut reader = super::make_fast_reader(out.as_bytes());
        let mut text = String::new();
        loop {
            match reader.read_event().unwrap() {
                quick_xml::events::Event::Text(t) => text.push_str(&t),
                quick_xml::events::Event::Eof => break,
                _ => {},
            }
        }
        assert_eq!(text, "caf\u{FFFD} ok");
    }

    /// Without a declaration, undeclared 8-bit bytes are read as ISO-8859-1
    /// (the existing fallback) rather than replaced.
    #[test]
    fn test_undeclared_latin1_is_transcoded() {
        let xml = b"<a>caf\xE9</a>";
        let out = String::from_utf8(ensure_utf8(xml).unwrap()).unwrap();
        assert_eq!(out, "<a>caf\u{e9}</a>");
    }
}
