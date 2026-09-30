//! Minimal synthetic legacy PowerPoint (`.ppt`) writer for tests.
//!
//! Builds a "PowerPoint Document" record stream the way a real save lays
//! it out — a `DocumentContainer`, one `Slide` container per slide, an
//! optional `Notes` container per slide, a `PersistDirectoryAtom` and a
//! `UserEditAtom` — plus the "Current User" stream that points at the
//! edit, wrapped in a CFB container. Covers only what `PptDocument`
//! reads; every record layout cites its [MS-PPT] section.

#![allow(dead_code)]

/// `DocumentContainer` ([MS-PPT]).
pub const RT_DOCUMENT: u16 = 0x03E8;
/// `SlideContainer` ([MS-PPT]).
pub const RT_SLIDE: u16 = 0x03EE;
/// `SlideAtom` ([MS-PPT]).
pub const RT_SLIDE_ATOM: u16 = 0x03EF;
/// `NotesContainer` ([MS-PPT]).
pub const RT_NOTES: u16 = 0x03F0;
/// `NotesAtom` ([MS-PPT]).
pub const RT_NOTES_ATOM: u16 = 0x03F1;
/// `EnvironmentContainer` ([MS-PPT]).
pub const RT_ENVIRONMENT: u16 = 0x03F2;
/// `SlidePersistAtom` ([MS-PPT]).
pub const RT_SLIDE_PERSIST_ATOM: u16 = 0x03F3;
/// `FontCollectionContainer` ([MS-PPT]).
pub const RT_FONT_COLLECTION: u16 = 0x07D5;
/// `TextHeaderAtom` ([MS-PPT]).
pub const RT_TEXT_HEADER: u16 = 0x0F9F;
/// `TextCharsAtom` ([MS-PPT]).
pub const RT_TEXT_CHARS: u16 = 0x0FA0;
/// `StyleTextPropAtom` ([MS-PPT]).
pub const RT_STYLE_TEXT_PROP: u16 = 0x0FA1;
/// `TextBytesAtom` ([MS-PPT]).
pub const RT_TEXT_BYTES: u16 = 0x0FA8;
/// `FontEntityAtom` ([MS-PPT]).
pub const RT_FONT_ENTITY_ATOM: u16 = 0x0FB7;
/// `CString` ([MS-PPT] atom used by many containers).
pub const RT_CSTRING: u16 = 0x0FBA;
/// `HeadersFootersContainer` ([MS-PPT]).
pub const RT_HEADERS_FOOTERS: u16 = 0x0FD9;
/// `HeadersFootersAtom` ([MS-PPT]).
pub const RT_HEADERS_FOOTERS_ATOM: u16 = 0x0FDA;
/// `SlideListWithTextContainer` family ([MS-PPT]/.6).
pub const RT_SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
/// `UserEditAtom` ([MS-PPT]).
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;
/// `CurrentUserAtom` ([MS-PPT]).
pub const RT_CURRENT_USER_ATOM: u16 = 0x0FF6;
/// `PersistDirectoryAtom` ([MS-PPT]).
pub const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;
/// OfficeArt `SpContainer` ([MS-ODRAW]).
pub const RT_SHAPE: u16 = 0xF004;
/// `OfficeArtClientTextbox` ([MS-PPT]): a shape's text holder.
pub const RT_CLIENT_TEXTBOX: u16 = 0xF00D;

/// `CurrentUserAtom.headerToken` of an unencrypted document ([MS-PPT]
/// 2.3.2).
pub const HEADER_TOKEN_PLAIN: u32 = 0xE391_C05F;
/// `CurrentUserAtom.headerToken` of an encrypted document.
pub const HEADER_TOKEN_ENCRYPTED: u32 = 0xF3D1_C4DF;

/// An atom record: 8-byte header (`recVer` 0) plus body.
pub fn atom(rec_type: u16, instance: u16, body: &[u8]) -> Vec<u8> {
    let mut v = (instance << 4).to_le_bytes().to_vec();
    v.extend_from_slice(&rec_type.to_le_bytes());
    v.extend_from_slice(&(body.len() as u32).to_le_bytes());
    v.extend_from_slice(body);
    v
}

/// A container record (`recVer` 0xF).
pub fn container(rec_type: u16, instance: u16, children: &[u8]) -> Vec<u8> {
    let mut v = ((instance << 4) | 0xF).to_le_bytes().to_vec();
    v.extend_from_slice(&rec_type.to_le_bytes());
    v.extend_from_slice(&(children.len() as u32).to_le_bytes());
    v.extend_from_slice(children);
    v
}

/// UTF-16LE bytes of `s`.
pub fn utf16(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// A text shape: `TextHeaderAtom` (`text_type`, [MS-PPT] `TextTypeEnum`)
/// plus a `TextCharsAtom`, followed by `extra` records (e.g. a
/// `StyleTextPropAtom`), inside a `ClientTextbox` inside an `SpContainer`.
pub fn text_shape(text_type: u32, text: &str, extra: &[u8]) -> Vec<u8> {
    let mut tb = atom(RT_TEXT_HEADER, 0, &text_type.to_le_bytes());
    tb.extend(atom(RT_TEXT_CHARS, 0, &utf16(text)));
    tb.extend_from_slice(extra);
    container(RT_SHAPE, 0, &container(RT_CLIENT_TEXTBOX, 0, &tb))
}

/// One slide to write.
#[derive(Default, Clone)]
pub struct PptSlide {
    /// The slide's shapes and other children, after its `SlideAtom`.
    pub shapes: Vec<u8>,
    /// Speaker notes text; `Some` writes a `Notes` container for it.
    pub notes: Option<String>,
    /// Link the notes through `SlideAtom.notesIdRef` (the normal way);
    /// `false` leaves it 0 so only `NotesAtom.slideIdRef` links them.
    pub link_notes_from_slide: bool,
}

/// A deck to write.
#[derive(Default, Clone)]
pub struct PptBuilder {
    pub slides: Vec<PptSlide>,
    /// Extra `DocumentContainer` children (headers/footers, ExObjList, …).
    pub doc_children: Vec<u8>,
    /// `FontEntityAtom` face names, in font-collection order.
    pub fonts: Vec<String>,
    /// `CurrentUserAtom.headerToken`; `None` writes the unencrypted token.
    pub header_token: Option<u32>,
    /// `UserEditAtom.encryptSessionPersistIdRef`; `Some` writes the
    /// optional field (present only in encrypted documents).
    pub encrypt_session_persist_id: Option<u32>,
    /// Extra CFB root streams.
    pub extra_streams: Vec<(String, Vec<u8>)>,
}

impl PptBuilder {
    /// The "PowerPoint Document" and "Current User" stream bytes.
    pub fn streams(&self) -> (Vec<u8>, Vec<u8>) {
        // Persist ids: 1 = document, then per slide its slide container
        // and (if any) its notes container.
        let mut next_id = 2u32;
        let mut slide_ids = Vec::new(); // (persist id, slide id)
        let mut notes_ids = Vec::new(); // Option<(persist id, notes id)>
        for (i, s) in self.slides.iter().enumerate() {
            slide_ids.push((next_id, 256 + i as u32));
            next_id += 1;
            if s.notes.is_some() {
                notes_ids.push(Some((next_id, 0x1000 + i as u32)));
                next_id += 1;
            } else {
                notes_ids.push(None);
            }
        }

        // SlidePersistAtom ([MS-PPT]): persistIdRef, flags,
        // cTexts, slideId, reserved.
        let persist_atom = |persist_id: u32, slide_id: u32| {
            let mut b = persist_id.to_le_bytes().to_vec();
            b.extend_from_slice(&0u32.to_le_bytes());
            b.extend_from_slice(&0u32.to_le_bytes());
            b.extend_from_slice(&slide_id.to_le_bytes());
            b.extend_from_slice(&0u32.to_le_bytes());
            atom(RT_SLIDE_PERSIST_ATOM, 0, &b)
        };

        let mut doc_children = Vec::new();
        if !self.fonts.is_empty() {
            let mut fonts = Vec::new();
            for (i, name) in self.fonts.iter().enumerate() {
                // FontEntityAtom ([MS-PPT]): lfFaceName (32 UTF-16
                // units, NUL-padded), then charset/flags/pitch (4 bytes).
                let mut b = vec![0u8; 68];
                let n = utf16(name);
                b[..n.len().min(62)].copy_from_slice(&n[..n.len().min(62)]);
                fonts.extend(atom(RT_FONT_ENTITY_ATOM, i as u16, &b));
            }
            doc_children.extend(container(
                RT_ENVIRONMENT,
                0,
                &container(RT_FONT_COLLECTION, 0, &fonts),
            ));
        }
        let mut slwt = Vec::new();
        for &(pid, sid) in &slide_ids {
            slwt.extend(persist_atom(pid, sid));
        }
        doc_children.extend(container(RT_SLIDE_LIST_WITH_TEXT, 0, &slwt));
        let mut nlwt = Vec::new();
        for (pid, nid) in notes_ids.iter().flatten() {
            nlwt.extend(persist_atom(*pid, *nid));
        }
        if !nlwt.is_empty() {
            // NotesListWithTextContainer: recInstance 2 ([MS-PPT]).
            doc_children.extend(container(RT_SLIDE_LIST_WITH_TEXT, 2, &nlwt));
        }
        doc_children.extend_from_slice(&self.doc_children);

        let mut stream = Vec::new();
        let mut offsets: Vec<(u32, u32)> = Vec::new();
        offsets.push((1, stream.len() as u32));
        stream.extend(container(RT_DOCUMENT, 0, &doc_children));
        for (i, s) in self.slides.iter().enumerate() {
            let (pid, sid) = slide_ids[i];
            // SlideAtom ([MS-PPT]): geom + rgPlaceholderTypes (12),
            // masterIdRef @12, notesIdRef @16, slideFlags @20, unused.
            let mut sa = vec![0u8; 24];
            if let (Some((_, nid)), true) = (notes_ids[i], s.link_notes_from_slide) {
                sa[16..20].copy_from_slice(&nid.to_le_bytes());
            }
            let mut children = atom(RT_SLIDE_ATOM, 2, &sa);
            children.extend_from_slice(&s.shapes);
            offsets.push((pid, stream.len() as u32));
            stream.extend(container(RT_SLIDE, 0, &children));
            if let (Some((npid, _)), Some(text)) = (notes_ids[i], &s.notes) {
                // NotesAtom ([MS-PPT]): slideIdRef, slideFlags, unused.
                let mut na = sid.to_le_bytes().to_vec();
                na.extend_from_slice(&[0u8; 4]);
                let mut nc = atom(RT_NOTES_ATOM, 1, &na);
                nc.extend(text_shape(2, text, &[])); // Tx_TYPE_NOTES
                offsets.push((npid, stream.len() as u32));
                stream.extend(container(RT_NOTES, 0, &nc));
            }
        }

        // PersistDirectoryAtom ([MS-PPT]): one entry per id.
        let pd_offset = stream.len() as u32;
        let mut pd = Vec::new();
        for (id, off) in &offsets {
            pd.extend_from_slice(&((1u32 << 20) | id).to_le_bytes());
            pd.extend_from_slice(&off.to_le_bytes());
        }
        stream.extend(atom(RT_PERSIST_DIRECTORY_ATOM, 0, &pd));

        // UserEditAtom ([MS-PPT]): lastSlideIdRef, version(2),
        // minorVersion, majorVersion, offsetLastEdit, offsetPersistDirectory,
        // docPersistIdRef, persistIdSeed, lastView(2), unused(2), then the
        // optional encryptSessionPersistIdRef.
        let edit_offset = stream.len() as u32;
        let mut ue = Vec::new();
        ue.extend_from_slice(&256u32.to_le_bytes());
        ue.extend_from_slice(&[0, 0, 0, 3]);
        ue.extend_from_slice(&0u32.to_le_bytes());
        ue.extend_from_slice(&pd_offset.to_le_bytes());
        ue.extend_from_slice(&1u32.to_le_bytes());
        ue.extend_from_slice(&next_id.to_le_bytes());
        ue.extend_from_slice(&[1, 0, 0, 0]);
        if let Some(id) = self.encrypt_session_persist_id {
            ue.extend_from_slice(&id.to_le_bytes());
        }
        stream.extend(atom(RT_USER_EDIT_ATOM, 0, &ue));

        // CurrentUserAtom ([MS-PPT]): size (0x14), headerToken,
        // offsetToCurrentEdit, lenUserName, docFileVersion, major, minor,
        // unused, ansiUserName, relVersion.
        let mut cu = Vec::new();
        cu.extend_from_slice(&0x14u32.to_le_bytes());
        cu.extend_from_slice(
            &self
                .header_token
                .unwrap_or(HEADER_TOKEN_PLAIN)
                .to_le_bytes(),
        );
        cu.extend_from_slice(&edit_offset.to_le_bytes());
        cu.extend_from_slice(&1u16.to_le_bytes());
        cu.extend_from_slice(&0x03F4u16.to_le_bytes());
        cu.extend_from_slice(&[3, 0, 0, 0]);
        cu.push(b'u');
        cu.extend_from_slice(&8u32.to_le_bytes());
        let current_user = atom(RT_CURRENT_USER_ATOM, 0, &cu);
        (stream, current_user)
    }

    /// The complete `.ppt` file.
    pub fn build(&self) -> Vec<u8> {
        let (stream, current_user) = self.streams();
        let mut streams: Vec<(&str, &[u8])> = vec![
            ("PowerPoint Document", &stream),
            ("Current User", &current_user),
        ];
        for (name, data) in &self.extra_streams {
            streams.push((name, data));
        }
        super::cfb_with_streams(&streams)
    }
}

/// A `StyleTextPropAtom` ([MS-PPT]) for `char_count` characters:
/// one paragraph run with no properties, then the given character runs
/// `(count, masks, fields)` — `fields` being the `TextCFException` bytes
/// the masks select, in spec order.
pub fn style_text_prop(char_count: u32, cf_runs: &[(u32, u32, Vec<u8>)]) -> Vec<u8> {
    let mut b = Vec::new();
    // TextPFRun: count, indentLevel, PFMasks = 0.
    b.extend_from_slice(&(char_count + 1).to_le_bytes());
    b.extend_from_slice(&0u16.to_le_bytes());
    b.extend_from_slice(&0u32.to_le_bytes());
    for (count, masks, fields) in cf_runs {
        b.extend_from_slice(&count.to_le_bytes());
        b.extend_from_slice(&masks.to_le_bytes());
        b.extend_from_slice(fields);
    }
    atom(RT_STYLE_TEXT_PROP, 0, &b)
}

/// A `HeadersFootersContainer` ([MS-PPT]) of the given instance
/// (3 = slides, 4 = notes/handouts) with its `HeadersFootersAtom` flags
/// and `(CString instance, text)` strings (0 user date, 1 header, 2
/// footer).
pub fn headers_footers(instance: u16, flags: u16, strings: &[(u16, &str)]) -> Vec<u8> {
    let mut b = 0u16.to_le_bytes().to_vec(); // formatId
    b.extend_from_slice(&flags.to_le_bytes());
    let mut children = atom(RT_HEADERS_FOOTERS_ATOM, 0, &b);
    for (inst, text) in strings {
        children.extend(atom(RT_CSTRING, *inst, &utf16(text)));
    }
    container(RT_HEADERS_FOOTERS, instance, &children)
}
