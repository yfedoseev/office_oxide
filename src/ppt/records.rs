#![allow(dead_code)]
//! PowerPoint binary record types and parsing.
//!
//! PPT records have an 8-byte header:
//! - Bits 0-3: recVer (version)
//! - Bits 4-15: recInstance
//! - Bytes 2-3: recType (u16)
//! - Bytes 4-7: recLen (u32)
//!
//! Container records (recVer = 0xF) contain child records.
//! Atom records contain raw data.

use super::error::{PptError, Result};

// ── Record type IDs ──
pub const RT_DOCUMENT: u16 = 0x03E8;
pub const RT_SLIDE: u16 = 0x03EE;
pub const RT_SLIDE_BASE: u16 = 0x03EC;
/// SlideAtom ([MS-PPT] 2.4.2, record type 1007) — a `Slide` container's
/// own atom; carries `masterIdRef`, the persist ID of the main master
/// this slide inherits formatting from (issue #335).
pub const RT_SLIDE_ATOM: u16 = 0x03EF;
/// MainMaster container ([MS-PPT] 2.5.3, record type 1016) — the slide
/// master; among its children are up to several `TxMasterStyleAtom`s,
/// one per text type, each holding the level-indexed default
/// character/paragraph formatting a placeholder shape falls back to for
/// any property its own direct `StyleTextPropAtom` didn't set (issue
/// #335).
pub const RT_MAIN_MASTER: u16 = 0x03F8;
/// TxMasterStyleAtom ([MS-PPT] 2.9.5, record type 4003). `recInstance`
/// is itself the `TextTypeEnum` value this atom's styles apply to — "the
/// atom instance value is the text type", per Apache POI's own doc
/// comment on the equivalent class (issue #335).
pub const RT_TX_MASTER_STYLE_ATOM: u16 = 0x0FA3;
pub const RT_NOTES: u16 = 0x03F0;
pub const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
pub const RT_TEXT_HEADER: u16 = 0x0F9F;
/// OutlineTextRefAtom ([MS-PPT] 2.4.15.6) — a shape's text stored *by
/// reference* as a zero-based index into the TextHeaderAtom sequence that
/// follows its slide's SlidePersistAtom in SlideListWithTextContainer,
/// instead of embedded directly in the shape.
pub const RT_OUTLINE_TEXT_REF_ATOM: u16 = 0x0F9E;
pub const RT_TEXT_CHARS: u16 = 0x0FA0;
pub const RT_TEXT_BYTES: u16 = 0x0FA8;
pub const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
/// Optional child of a `Slide` container: which slideshow transition to
/// use, and whether the slide is hidden ([MS-PPT] 2.5.1 `SlideContainer`,
/// 2.4.15.4/2.13.24 `SlideShowSlideInfoAtom`, record type 1017 = 0x3F9).
/// Issue #297.
pub const RT_SLIDE_SHOW_SLIDE_INFO_ATOM: u16 = 0x03F9;
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
/// PersistDirectoryAtom ([MS-PPT] 2.3.4).
pub const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
pub const RT_CURRENT_USER_ATOM: u16 = 0x0FF6;
/// HeadersFootersContainer: the header/footer/user-date text applied to
/// a presentation ([MS-PPT] 2.4.16, `HeadersFootersContainer`, record
/// type 4057 = 0x0FD9). The pre-existing constant here had the wrong
/// value (0x0FDA, which is actually `HeadersFootersAtom` — the small
/// flags atom nested *inside* this container, not the container itself)
/// — verified byte-for-byte against a real corpus file's `DocumentContainer`
/// child (issue #308).
pub const RT_HEADER_FOOTER: u16 = 0x0FD9;
/// `HeadersFootersAtom`: the flags atom nested inside a
/// `HeadersFootersContainer` ([MS-PPT] 2.4.17, record type 4058 =
/// 0x0FDA). Not decoded — this fix only surfaces the container's
/// `CString` text children, not the show/hide flag bits.
pub const RT_HEADER_FOOTER_ATOM: u16 = 0x0FDA;
pub const RT_TEXT_SPECIAL_INFO: u16 = 0x0FAA;
pub const RT_TEXT_RULER: u16 = 0x0FA2;
pub const RT_STYLE_TEXT_PROP: u16 = 0x0FA1;
pub const RT_CSTRING: u16 = 0x0FBA;
/// `TargetAtom`'s own `rh.recInstance` value ([MS-PPT] 2.10.19) — the
/// `RT_CSTRING` sibling within an `ExHyperlinkContainer` that carries the
/// hyperlink's actual target URL/path, as opposed to `FriendlyNameAtom`
/// or `LocationAtom` (other `RT_CSTRING` children at different instances).
pub const CSTRING_INSTANCE_TARGET: u16 = 0x001;

/// [MS-ODRAW] §2.2.16 `OfficeArtSpgrContainer` — a group of shapes. Its
/// first child `RT_SHAPE` is the group's own placeholder shape (no
/// `RT_CHILD_ANCHOR`); every subsequent `RT_SHAPE` child is a real member
/// of the group, positioned via `RT_CHILD_ANCHOR` (issue #255).
pub const RT_SPGR_CONTAINER: u16 = 0xF003;
/// [MS-ODRAW] §2.2.16 `OfficeArtSpContainer` — a single shape (its
/// properties, anchor, client data, and text box, as children).
pub const RT_SHAPE: u16 = 0xF004;
/// [MS-ODRAW] §2.2.16 `OfficeArtChildAnchor` — a group member shape's
/// position in its group's local coordinate space: `xLeft`/`yTop`/
/// `xRight`/`yBottom`, each a signed 32-bit integer (issue #255).
pub const RT_CHILD_ANCHOR: u16 = 0xF00F;
/// [MS-ODRAW] §2.2.9 `OfficeArtFOPT` — a shape's property table
/// (`rh.recInstance` = property count), including `pib` ("Blip to
/// display"), which resolves a picture shape to its actual image
/// (issue #256).
pub const RT_FOPT: u16 = 0xF00B;
/// [MS-PPT] 2.7.3 `OfficeArtClientData` — a shape's PPT-specific data
/// (placeholder role, animation, interactive info), a child of `RT_SHAPE`.
pub const RT_CLIENT_DATA: u16 = 0xF011;
/// [MS-PPT] 2.10.1 `ExObjListContainer` — the document-wide table of
/// external objects (hyperlinks, embedded media), a child of the top-level
/// `DocumentContainer`.
pub const RT_EXTERNAL_OBJECT_LIST: u16 = 0x0409;
/// [MS-PPT] 2.10.16 `ExHyperlinkContainer` — one hyperlink's own id +
/// target, found inside `RT_EXTERNAL_OBJECT_LIST`.
pub const RT_EXTERNAL_HYPERLINK: u16 = 0x0FD7;
/// [MS-PPT] 2.10.17 `ExHyperlinkAtom` — the numeric id (`exHyperlinkId`)
/// that `InteractiveInfoAtom.exHyperlinkIdRef` refers back to.
pub const RT_EXTERNAL_HYPERLINK_ATOM: u16 = 0x0FD3;
/// [MS-PPT] 2.6.9 `MouseClickInteractiveInfoContainer` /
/// `MouseOverInteractiveInfoContainer` — a shape's click/hover action,
/// found inside its `RT_CLIENT_DATA`.
pub const RT_INTERACTIVE_INFO: u16 = 0x0FF2;
/// [MS-PPT] 2.6.10 `InteractiveInfoAtom` — the action type + hyperlink id
/// reference, inside `RT_INTERACTIVE_INFO`.
pub const RT_INTERACTIVE_INFO_ATOM: u16 = 0x0FF3;
/// [MS-PPT] 2.6.11 `MouseClickTextInteractiveInfoAtom` /
/// `MouseOverTextInteractiveInfoAtom` — the character range (within the
/// text of the nearest preceding `TextHeaderAtom`) that a *sibling*
/// `RT_INTERACTIVE_INFO` (appearing directly in a `ClientTextbox`, not
/// nested in `RT_CLIENT_DATA`) anchors its hyperlink to. This is the
/// text-run-level hyperlink mechanism — distinct from, and far more
/// common than, the whole-shape one via `RT_CLIENT_DATA` (issue #257).
pub const RT_TEXT_INTERACTIVE_INFO_ATOM: u16 = 0x0FDF;
/// [MS-PPT] 2.10.20 `ExOleObjAtom` — one embedded/linked/ActiveX OLE
/// object's identity: `objID` (joins back to a shape's `ExObjRefAtom`),
/// `subType` (Excel/Word/Equation Editor/etc.), and `type` (embedded=0,
/// linked=1, control=2). Found inside an `ExEmbed` container
/// (`RT_EXTERNAL_OLE_EMBED`), itself inside `RT_EXTERNAL_OBJECT_LIST`
/// (issue #337).
pub const RT_EXTERNAL_OLE_OBJECT_ATOM: u16 = 0x0FC3;
/// [MS-PPT] 2.10.19 `ExEmbed` — the container wrapping one
/// `ExOleObjAtom` (plus `CString` menu/progId/clipboard-format names
/// this crate doesn't need for identity purposes), a child of
/// `RT_EXTERNAL_OBJECT_LIST` (issue #337).
pub const RT_EXTERNAL_OLE_EMBED: u16 = 0x0FCC;
/// [MS-PPT] 2.4.9.2 `ExObjRefAtom` — a shape's own reference (`exObjIdRef`)
/// to an external object (an `ExOleObjAtom` or `ExMediaAtom`), found
/// directly inside its `RT_CLIENT_DATA` (issue #337).
pub const RT_EXTERNAL_OBJECT_REF_ATOM: u16 = 0x0BC1;

// ── SlideListWithText `rh.recInstance` discriminants ([MS-PPT] 2.4.14) ──
/// `rh.recInstance` value identifying a `SlideListWithTextContainer` (real slides).
pub const SLWT_SLIDES: u16 = 0;
/// `rh.recInstance` value identifying a `MasterListWithTextContainer`.
pub const SLWT_MASTERS: u16 = 1;
/// `rh.recInstance` value identifying a `NotesListWithTextContainer`.
pub const SLWT_NOTES: u16 = 2;

/// A parsed PPT record header.
#[derive(Debug, Clone, Copy)]
pub struct RecordHeader {
    pub rec_ver: u8,
    pub rec_instance: u16,
    pub rec_type: u16,
    pub rec_len: u32,
}

impl RecordHeader {
    pub fn is_container(&self) -> bool {
        self.rec_ver == 0x0F
    }

    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(PptError::InvalidRecord("record header too short".into()));
        }
        let ver_instance = u16::from_le_bytes([data[0], data[1]]);
        let rec_ver = (ver_instance & 0x0F) as u8;
        let rec_instance = ver_instance >> 4;
        let rec_type = u16::from_le_bytes([data[2], data[3]]);
        let rec_len = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
        Ok(Self {
            rec_ver,
            rec_instance,
            rec_type,
            rec_len,
        })
    }
}

/// A PPT record with its header and data.
#[derive(Debug, Clone)]
pub struct PptRecord {
    pub header: RecordHeader,
    pub data: Vec<u8>,
    pub offset: usize,
}

/// Iterate over the immediate children of one record region (single level,
/// non-recursive).
///
/// Each yielded record's `data` is bounded strictly by that record's own
/// declared `rec_len`, clamped to whatever bytes actually remain in the slice
/// passed to [`RecordIter::new`] — never by descending into (or trusting the
/// declared lengths of) anything nested further down. For a container record,
/// `data` is its bounded *children* region; recurse by constructing a new
/// `RecordIter::new(&container_record.data)`.
///
/// This bounding is what keeps a corrupted or maliciously oversized `rec_len`
/// on one record from swallowing bytes that belong to its siblings, or to
/// anything outside its own enclosing container: skipping to the next sibling
/// only ever relies on the current record's own header, never on correctly
/// parsing what's inside it.
pub struct RecordIter<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> RecordIter<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }
}

impl<'a> Iterator for RecordIter<'a> {
    type Item = Result<PptRecord>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.pos + 8 > self.data.len() {
            return None;
        }

        let header = match RecordHeader::parse(&self.data[self.pos..]) {
            Ok(h) => h,
            Err(e) => return Some(Err(e)),
        };

        let record_offset = self.pos;
        let data_start = (self.pos + 8).min(self.data.len());
        // Clamp to *this* slice's own end — never trust rec_len past what's
        // actually here, whether the record is an atom or a container.
        let data_end = data_start
            .saturating_add(header.rec_len as usize)
            .min(self.data.len());

        self.pos = data_end;

        Some(Ok(PptRecord {
            header,
            data: self.data[data_start..data_end].to_vec(),
            offset: record_offset,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_atom(rec_type: u16, instance: u16, data: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = instance << 4; // ver=0 (atom)
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u32).to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    fn make_container(rec_type: u16, instance: u16, children: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = (instance << 4) | 0x0F; // ver=0xF (container)
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(children.len() as u32).to_le_bytes());
        buf.extend_from_slice(children);
        buf
    }

    #[test]
    fn parse_record_header() {
        let data = make_atom(0x0FA0, 0, &[0x41, 0x00]);
        let header = RecordHeader::parse(&data).unwrap();
        assert_eq!(header.rec_type, RT_TEXT_CHARS);
        assert_eq!(header.rec_ver, 0);
        assert!(!header.is_container());
        assert_eq!(header.rec_len, 2);
    }

    #[test]
    fn parse_container_header() {
        let child = make_atom(RT_TEXT_CHARS, 0, &[0x41, 0x00]);
        let data = make_container(RT_SLIDE, 0, &child);
        let header = RecordHeader::parse(&data).unwrap();
        assert!(header.is_container());
        assert_eq!(header.rec_type, RT_SLIDE);
    }

    #[test]
    fn iterate_flat_atoms() {
        let mut stream = make_atom(RT_TEXT_HEADER, 0, &[0x00, 0x00, 0x00, 0x00]);
        stream.extend(make_atom(RT_TEXT_CHARS, 0, &[0x48, 0x00, 0x69, 0x00]));
        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 2);
        assert_eq!(records[0].header.rec_type, RT_TEXT_HEADER);
        assert_eq!(records[1].header.rec_type, RT_TEXT_CHARS);
    }

    #[test]
    fn container_children_are_bounded_not_flattened() {
        let child1 = make_atom(RT_TEXT_HEADER, 0, &[0x00, 0x00, 0x00, 0x00]);
        let child2 = make_atom(RT_TEXT_CHARS, 0, &[0x41, 0x00]);
        let mut children = child1.clone();
        children.extend(&child2);
        let container = make_container(RT_SLIDE_LIST_WITH_TEXT, 0, &children);

        // A single-level iteration over the container yields the container
        // itself, not its descendants — it must not implicitly flatten.
        let records: Vec<_> = RecordIter::new(&container)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(records.len(), 1);
        assert_eq!(records[0].header.rec_type, RT_SLIDE_LIST_WITH_TEXT);
        assert!(records[0].header.is_container());
        assert_eq!(records[0].data, children);

        // Recursing explicitly into the container's own bounded child region
        // reaches both children.
        let nested: Vec<_> = RecordIter::new(&records[0].data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(nested.len(), 2);
        assert_eq!(nested[0].header.rec_type, RT_TEXT_HEADER);
        assert_eq!(nested[1].header.rec_type, RT_TEXT_CHARS);
    }

    #[test]
    fn empty_stream() {
        let records: Vec<_> = RecordIter::new(&[])
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert!(records.is_empty());
    }

    /// A single record with a corrupted/oversized declared length, nested
    /// inside a container, must not swallow bytes belonging to a sibling
    /// record *outside* that container. Skipping to the next sibling must
    /// rely only on the container's own header length, never on successfully
    /// parsing (or bounding) what's inside it.
    #[test]
    fn corrupt_nested_record_length_does_not_swallow_top_level_siblings() {
        // A child atom that declares a wildly oversized length with no data
        // behind it (simulating real-world corrupted/non-conformant files).
        let mut corrupt_child = Vec::new();
        corrupt_child.extend_from_slice(&0u16.to_le_bytes()); // ver=0 (atom), instance=0
        corrupt_child.extend_from_slice(&RT_TEXT_CHARS.to_le_bytes());
        corrupt_child.extend_from_slice(&1_000_000u32.to_le_bytes()); // bogus length

        let outer = make_container(RT_SLIDE, 0, &corrupt_child);

        let mut stream = outer;
        stream.extend(make_atom(RT_TEXT_BYTES, 0, b"sibling text"));

        let records: Vec<_> = RecordIter::new(&stream)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(
            records.len(),
            2,
            "corrupt length nested inside the container must not swallow the top-level sibling after it"
        );
        assert_eq!(records[0].header.rec_type, RT_SLIDE);
        assert_eq!(records[1].header.rec_type, RT_TEXT_BYTES);
        assert_eq!(records[1].data, b"sibling text");

        // Descending into the corrupt container's own children must clamp to
        // the bytes actually available, not panic or run past the container.
        let inner: Vec<_> = RecordIter::new(&records[0].data)
            .collect::<std::result::Result<_, _>>()
            .unwrap();
        assert_eq!(inner.len(), 1);
        assert!(inner[0].data.len() < 1_000_000);
    }
}
