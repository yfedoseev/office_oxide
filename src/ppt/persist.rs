//! Persist object directory resolution for legacy PPT97 binary files.
//!
//! PowerPoint 97 uses incremental ("fast") saves: each save appends new or
//! changed top-level records to the end of the "PowerPoint Document" stream
//! and writes a fresh `UserEditAtom` + `PersistDirectoryAtom` recording where
//! the *current* copy of every persist object now lives. Older, superseded
//! copies of records (e.g. a `Slide` container from before an edit) are left
//! behind in the stream — nothing compacts them away.
//!
//! A naive sequential scan of the stream from front to back therefore cannot
//! tell current content from stale, orphaned leftovers: whichever copy
//! happens to come first wins. Locating the *current* content for a given
//! slide requires resolving its `persistIdRef` through this directory, built
//! by walking `CurrentUserAtom.offsetToCurrentEdit` → `UserEditAtom` chain
//! (via `offsetLastEdit`) → merged `PersistDirectoryAtom` entries, per
//! [MS-PPT] 2.1.2.

use std::collections::{HashMap, HashSet};

use super::records::{
    RT_CURRENT_USER_ATOM, RT_PERSIST_DIRECTORY_ATOM, RT_USER_EDIT_ATOM, RecordHeader, RecordIter,
};

/// Guards against a corrupted or cyclic `offsetLastEdit` chain.
const MAX_EDIT_CHAIN_LEN: usize = 4096;

/// Resolves a persist object identifier to its current byte offset within the
/// "PowerPoint Document" stream.
#[derive(Debug, Default, Clone)]
pub struct PersistDirectory {
    offsets: HashMap<u32, u32>,
    /// Persist ID of the current `DocumentContainer` (`UserEditAtom.docPersistIdRef`).
    pub doc_persist_id: u32,
}

impl PersistDirectory {
    /// Byte offset of the persist object with the given ID, if known.
    pub fn resolve(&self, persist_id: u32) -> Option<usize> {
        self.offsets.get(&persist_id).map(|&o| o as usize)
    }
}

/// Build the persist object directory for `stream` (the raw "PowerPoint
/// Document" bytes).
///
/// `current_user` is the raw "Current User" stream, if present; its
/// `offsetToCurrentEdit` field is the normal, spec-correct way to find the
/// most recent `UserEditAtom`. When that stream is absent, or its offset
/// doesn't land on a valid `UserEditAtom`, falls back to a bounded top-level
/// scan of `stream` for the highest-offset `UserEditAtom` — every save
/// appends a fresh one at the end, so the highest offset is always the most
/// recent edit regardless of what the `Current User` stream says.
///
/// Returns `None` if no `UserEditAtom` can be found at all (e.g. a minimal
/// hand-built stream in a test) — callers should fall back to a different
/// extraction strategy in that case.
pub fn build(stream: &[u8], current_user: Option<&[u8]>) -> Option<PersistDirectory> {
    let start = current_user
        .and_then(offset_to_current_edit)
        .filter(|&off| user_edit_atom_at(stream, off).is_some())
        .or_else(|| find_latest_user_edit_atom(stream))?;

    let mut dir = PersistDirectory::default();
    let mut offset = Some(start);
    let mut seen = HashSet::new();
    let mut doc_persist_id = None;

    for _ in 0..MAX_EDIT_CHAIN_LEN {
        let Some(off) = offset else { break };
        if !seen.insert(off) {
            break; // cyclic offsetLastEdit chain
        }

        let Some(edit) = user_edit_atom_at(stream, off) else {
            break;
        };
        if doc_persist_id.is_none() {
            doc_persist_id = Some(edit.doc_persist_id_ref);
        }

        if let Some(entries) = persist_directory_entries_at(stream, edit.offset_persist_directory) {
            // Visiting newest-edit-first: only fill in IDs not already known,
            // so an older edit's entry never overwrites a newer one.
            for (id, off) in entries {
                dir.offsets.entry(id).or_insert(off);
            }
        }

        offset = if edit.offset_last_edit == 0 {
            None
        } else {
            Some(edit.offset_last_edit as usize)
        };
    }

    dir.doc_persist_id = doc_persist_id.unwrap_or(1);
    Some(dir)
}

struct UserEditAtom {
    offset_last_edit: u32,
    offset_persist_directory: u32,
    doc_persist_id_ref: u32,
}

/// Parse a `UserEditAtom` ([MS-PPT] 2.3.3) at an absolute offset in `stream`.
///
/// Body layout: `lastSlideIdRef`(4)@0, packed version fields(4)@4,
/// `offsetLastEdit`(4)@8, `offsetPersistDirectory`(4)@12,
/// `docPersistIdRef`(4)@16, `persistIdSeed`(4)@20, `lastView`+`unused`(4)@24.
fn user_edit_atom_at(stream: &[u8], offset: usize) -> Option<UserEditAtom> {
    let header_bytes = stream.get(offset..offset + 8)?;
    let header = RecordHeader::parse(header_bytes).ok()?;
    if header.rec_type != RT_USER_EDIT_ATOM || header.is_container() {
        return None;
    }
    let body_start = offset + 8;
    let body_end = body_start
        .saturating_add(header.rec_len as usize)
        .min(stream.len());
    let body = stream.get(body_start..body_end)?;
    if body.len() < 20 {
        return None;
    }
    Some(UserEditAtom {
        offset_last_edit: u32::from_le_bytes([body[8], body[9], body[10], body[11]]),
        offset_persist_directory: u32::from_le_bytes([body[12], body[13], body[14], body[15]]),
        doc_persist_id_ref: u32::from_le_bytes([body[16], body[17], body[18], body[19]]),
    })
}

/// Parse a `CurrentUserAtom` ([MS-PPT] 2.3.2) `offsetToCurrentEdit` field
/// (body offset 8, 4 bytes) from the raw "Current User" stream.
fn offset_to_current_edit(current_user: &[u8]) -> Option<usize> {
    let header = RecordHeader::parse(current_user.get(0..8)?).ok()?;
    if header.rec_type != RT_CURRENT_USER_ATOM || header.is_container() {
        return None;
    }
    let body = current_user.get(8..)?;
    let field = body.get(8..12)?;
    Some(u32::from_le_bytes([field[0], field[1], field[2], field[3]]) as usize)
}

/// Parse a `PersistDirectoryAtom` ([MS-PPT] 2.3.4/2.3.5) at an absolute
/// offset in `stream`, returning `(persistId, streamOffset)` pairs.
///
/// Each `PersistDirectoryEntry` is a 4-byte header (`persistId` in the low 20
/// bits, `cPersist` count in the high 12 bits) followed by `cPersist` 4-byte
/// stream offsets for sequential persist IDs starting at `persistId`.
fn persist_directory_entries_at(stream: &[u8], offset: u32) -> Option<Vec<(u32, u32)>> {
    let offset = offset as usize;
    let header_bytes = stream.get(offset..offset + 8)?;
    let header = RecordHeader::parse(header_bytes).ok()?;
    if header.rec_type != RT_PERSIST_DIRECTORY_ATOM || header.is_container() {
        return None;
    }
    let body_start = offset + 8;
    let body_end = body_start
        .saturating_add(header.rec_len as usize)
        .min(stream.len());
    let body = stream.get(body_start..body_end)?;

    let mut entries = Vec::new();
    let mut pos = 0usize;
    while pos + 4 <= body.len() {
        let word = u32::from_le_bytes([body[pos], body[pos + 1], body[pos + 2], body[pos + 3]]);
        let persist_id = word & 0x000F_FFFF;
        let c_persist = (word >> 20) as usize;
        pos += 4;

        for i in 0..c_persist {
            let Some(chunk) = body.get(pos..pos + 4) else {
                break;
            };
            let off = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
            entries.push((persist_id.wrapping_add(i as u32), off));
            pos += 4;
        }
    }
    Some(entries)
}

/// Bounded top-level scan for the highest-offset `UserEditAtom` in `stream`.
///
/// `UserEditAtom`s only ever appear at the top level of the stream (never
/// nested inside another container), so a single-level, non-recursive walk
/// is sufficient — and safe even when some unrelated top-level record has a
/// corrupted/oversized length, since `RecordIter` bounds each record by its
/// own declared length rather than by successfully parsing its contents.
fn find_latest_user_edit_atom(stream: &[u8]) -> Option<usize> {
    let mut latest = None;
    for rec in RecordIter::new(stream) {
        let Ok(rec) = rec else { break };
        if rec.header.rec_type == RT_USER_EDIT_ATOM && !rec.header.is_container() {
            latest = Some(rec.offset);
        }
    }
    latest
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_atom(rec_type: u16, instance: u16, data: &[u8]) -> Vec<u8> {
        let ver_instance: u16 = instance << 4;
        let mut buf = Vec::new();
        buf.extend_from_slice(&ver_instance.to_le_bytes());
        buf.extend_from_slice(&rec_type.to_le_bytes());
        buf.extend_from_slice(&(data.len() as u32).to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    fn user_edit_atom_bytes(
        offset_last_edit: u32,
        offset_persist_directory: u32,
        doc_persist_id_ref: u32,
    ) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0u32.to_le_bytes()); // lastSlideIdRef
        body.extend_from_slice(&0u32.to_le_bytes()); // version/minor/major
        body.extend_from_slice(&offset_last_edit.to_le_bytes());
        body.extend_from_slice(&offset_persist_directory.to_le_bytes());
        body.extend_from_slice(&doc_persist_id_ref.to_le_bytes());
        body.extend_from_slice(&0u32.to_le_bytes()); // persistIdSeed
        body.extend_from_slice(&0u32.to_le_bytes()); // lastView + unused
        make_atom(RT_USER_EDIT_ATOM, 0, &body)
    }

    fn persist_directory_bytes(entries: &[(u32, u32)]) -> Vec<u8> {
        assert!(!entries.is_empty());
        let persist_id = entries[0].0;
        let c_persist = entries.len() as u32;
        let header = persist_id | (c_persist << 20);
        let mut body = header.to_le_bytes().to_vec();
        for (_, off) in entries {
            body.extend_from_slice(&off.to_le_bytes());
        }
        make_atom(RT_PERSIST_DIRECTORY_ATOM, 0, &body)
    }

    fn current_user_bytes(offset_to_current_edit: u32) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0u32.to_le_bytes()); // size
        body.extend_from_slice(&0u32.to_le_bytes()); // headerToken
        body.extend_from_slice(&offset_to_current_edit.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // lenUserName
        body.extend_from_slice(&0u16.to_le_bytes()); // docFileVersion
        body.push(0); // majorVersion
        body.push(0); // minorVersion
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        make_atom(RT_CURRENT_USER_ATOM, 0, &body)
    }

    #[test]
    fn persist_directory_entry_bit_layout() {
        // persistId=5 (low 20 bits), cPersist=3 (high 12 bits), 3 offsets.
        let pd = persist_directory_bytes(&[(5, 100), (6, 200), (7, 300)]);
        let mut stream = vec![0u8; 16];
        stream.extend(pd);
        let entries = persist_directory_entries_at(&stream, 16).unwrap();
        assert_eq!(entries, vec![(5, 100), (6, 200), (7, 300)]);
    }

    #[test]
    fn resolves_current_edit_via_current_user_stream() {
        let mut stream = Vec::new();
        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, 0), (2, 42)]));
        let edit_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));

        let current_user = current_user_bytes(edit_offset);
        let dir = build(&stream, Some(&current_user)).unwrap();

        assert_eq!(dir.doc_persist_id, 1);
        assert_eq!(dir.resolve(1), Some(0));
        assert_eq!(dir.resolve(2), Some(42));
    }

    #[test]
    fn falls_back_to_brute_force_scan_without_current_user_stream() {
        let mut stream = Vec::new();
        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, 0), (2, 42)]));
        stream.extend(user_edit_atom_bytes(0, pd_offset, 1));

        let dir = build(&stream, None).unwrap();
        assert_eq!(dir.resolve(2), Some(42));
    }

    #[test]
    fn newer_edit_entry_wins_over_older_edit_for_same_persist_id() {
        let mut stream = Vec::new();

        // Oldest edit: persist id 2 -> stale offset 10.
        let pd1_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(2, 10)]));
        let edit1_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(0, pd1_offset, 1));

        // Newest edit: persist id 2 -> current offset 99, chained back to edit1.
        let pd2_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(2, 99)]));
        let edit2_offset = stream.len() as u32;
        stream.extend(user_edit_atom_bytes(edit1_offset, pd2_offset, 1));

        let current_user = current_user_bytes(edit2_offset);
        let dir = build(&stream, Some(&current_user)).unwrap();
        assert_eq!(dir.resolve(2), Some(99), "newest edit's directory entry must win");
    }

    #[test]
    fn cyclic_offset_last_edit_does_not_hang() {
        let mut stream = Vec::new();
        let pd_offset = stream.len() as u32;
        stream.extend(persist_directory_bytes(&[(1, 0)]));
        let edit_offset = stream.len() as u32;
        // offsetLastEdit points at itself.
        stream.extend(user_edit_atom_bytes(edit_offset, pd_offset, 1));

        let current_user = current_user_bytes(edit_offset);
        let dir = build(&stream, Some(&current_user)).unwrap();
        assert_eq!(dir.resolve(1), Some(0));
    }

    #[test]
    fn no_user_edit_atom_returns_none() {
        assert!(build(&[], None).is_none());
        assert!(build(b"not a ppt stream at all", None).is_none());
    }
}
