//! Embedded OLE object recognition (Excel workbooks, Equation Editor/
//! MathType, OLE Package, embedded Word/PowerPoint, etc.) inside a
//! DOC's `Storage:ObjectPool` storage.
//!
//! At minimum: recognize that an embedded object exists
//! and surface its identity, even without extracting its native
//! payload — the DOC analogue of the PPT OLE fix. Before
//! this, `ObjectPool` was never traversed at all, so a file with (for
//! example) 12 embedded Excel workbooks showed no trace of any of them
//! anywhere in the IR.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, DirEntry, EntryType};

/// Sentinel meaning "no sibling/child" in a CFB directory entry, mirrors
/// `cfb::directory::NO_ENTRY` (not reachable from here — that module is
/// private to `cfb`).
const NO_ENTRY: u32 = 0xFFFF_FFFF;

/// One embedded OLE object found under `ObjectPool`, identified only by
/// the presence of a well-known stream name within its own storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedOleObject {
    /// Human-readable identity, e.g. `"Embedded Microsoft Excel
    /// Workbook"`.
    pub description: String,
    /// The main text of an embedded Word document, when it could be read
    /// within the document's embedded-package account.
    pub text: Option<String>,
}

/// The description of an embedded Word document.
pub(super) const EMBEDDED_WORD: &str = "Embedded Microsoft Word Document";

/// Find every embedded object under the root's `ObjectPool` storage (if
/// present) and classify each by the presence of a well-known stream
/// name inside its own storage. Returns an empty `Vec` (not an error)
/// when there's no `ObjectPool` at all, or its structure doesn't match
/// expectations — this is a best-effort enhancement, never a hard
/// requirement for reading the rest of the document.
pub fn extract_ole_objects<R: Read + Seek>(reader: &CfbReader<R>) -> Vec<EmbeddedOleObject> {
    extract_ole_objects_with_storage(reader)
        .into_iter()
        .map(|(object, _)| object)
        .collect()
}

/// As [`extract_ole_objects`], with the directory index of each object's
/// storage, for reading what is inside it.
pub(super) fn extract_ole_objects_with_storage<R: Read + Seek>(
    reader: &CfbReader<R>,
) -> Vec<(EmbeddedOleObject, usize)> {
    classify_objects(reader.entries(), MAX_ENTRIES_WALKED).0
}

/// The walk behind [`extract_ole_objects`], over a bare directory, with
/// the per-walk cap as a parameter; also returns the entries the object
/// walks took.
fn classify_objects(
    entries: &[DirEntry],
    max_walk: usize,
) -> (Vec<(EmbeddedOleObject, usize)>, usize) {
    if entries.is_empty() || entries[0].entry_type != EntryType::RootStorage {
        return (Vec::new(), 0);
    }
    let Some(pool_idx) = find_direct_child(entries, entries[0].child, "ObjectPool") else {
        return (Vec::new(), 0);
    };
    let pool = &entries[pool_idx];
    if !matches!(pool.entry_type, EntryType::Storage | EntryType::RootStorage) {
        return (Vec::new(), 0);
    }

    let mut out = Vec::new();
    // The pool listing and the object walks each share one visited set, so
    // the pass is bounded by the directory, not by objects times entries.
    let mut objects = Walker::new(entries, max_walk);
    for obj_idx in Walker::new(entries, max_walk).collect(pool.child) {
        let obj = &entries[obj_idx];
        if obj.entry_type != EntryType::Storage {
            continue; // a stray non-storage entry directly under ObjectPool
        }
        let children = objects.collect(obj.child);
        let names: Vec<&str> = children.iter().map(|&i| entries[i].name.as_str()).collect();
        if let Some(description) = classify(&names) {
            out.push((
                EmbeddedOleObject {
                    description,
                    text: None,
                },
                obj_idx,
            ));
        }
    }
    (out, objects.steps)
}

/// Classify one embedded object's storage by the well-known stream
/// names it directly contains, most-specific first. `\u{1}CompObj`
/// carries the real ProgID (e.g. "Excel.Sheet.8") but requires parsing
/// [MS-OLEDS]'s `CompObjStream`; presence-based classification via the
/// payload stream name gets the common cases without that additional
/// parsing surface.
fn classify(names: &[&str]) -> Option<String> {
    let has = |n: &str| names.iter().any(|x| x.eq_ignore_ascii_case(n));
    if has("Workbook") || has("Book") {
        Some("Embedded Microsoft Excel Workbook".to_string())
    } else if has("Equation Native") {
        Some("Embedded Equation Editor/MathType Object".to_string())
    } else if has("WordDocument") {
        Some(EMBEDDED_WORD.to_string())
    } else if has("PowerPoint Document") {
        Some("Embedded Microsoft PowerPoint Presentation".to_string())
    } else if has("Package") {
        Some("Embedded OLE Package".to_string())
    } else if has("\u{1}CompObj") || has("\u{1}Ole") {
        Some("Embedded OLE Object".to_string())
    } else {
        None
    }
}

/// Find a direct child of the RB-tree rooted at `root` matching `name`,
/// case-insensitively. Explicit-stack with a visited set — the tree
/// comes from an untrusted file, so a corrupted/adversarial sibling
/// cycle must not overflow the stack or loop unboundedly. Each entry is
/// visited once and the walk stops at the first match.
pub(super) fn find_direct_child(entries: &[DirEntry], root: u32, name: &str) -> Option<usize> {
    let mut stack = vec![root];
    let mut seen = vec![false; entries.len()];
    while let Some(node) = stack.pop() {
        if node == NO_ENTRY || node as usize >= entries.len() || seen[node as usize] {
            continue;
        }
        seen[node as usize] = true;
        let entry = &entries[node as usize];
        if entry.name.eq_ignore_ascii_case(name) {
            return Some(node as usize);
        }
        stack.push(entry.left_sibling);
        stack.push(entry.right_sibling);
    }
    None
}

/// Bound on how many directory entries one walk may take, whatever the
/// directory's size. Well above any real document's embedded-object count.
const MAX_ENTRIES_WALKED: usize = 50_000;

/// Walks directory trees for one classification pass. The visited set is
/// allocated once and shared by every walk of the pass, so an entry an
/// earlier walk claimed is never walked again: the pass costs at most one
/// step per directory entry however many objects name the same subtree or
/// however the sibling pointers loop. Each single walk also stops at
/// `max_walk` entries.
struct Walker<'a> {
    entries: &'a [DirEntry],
    seen: Vec<bool>,
    max_walk: usize,
    /// Entries taken over every walk so far.
    steps: usize,
}

impl<'a> Walker<'a> {
    fn new(entries: &'a [DirEntry], max_walk: usize) -> Self {
        Self {
            entries,
            seen: vec![false; entries.len()],
            max_walk,
            steps: 0,
        }
    }

    /// The entry indices of the RB tree rooted at `root` that no earlier
    /// walk of this pass has taken. Explicit stack: the tree comes from an
    /// untrusted file and may be cyclic.
    fn collect(&mut self, root: u32) -> Vec<usize> {
        let mut out = Vec::new();
        let mut stack = vec![root];
        while let Some(node) = stack.pop() {
            if node == NO_ENTRY || node as usize >= self.entries.len() || self.seen[node as usize] {
                continue;
            }
            if out.len() >= self.max_walk {
                break;
            }
            self.seen[node as usize] = true;
            self.steps += 1;
            let entry = &self.entries[node as usize];
            out.push(node as usize);
            stack.push(entry.left_sibling);
            stack.push(entry.right_sibling);
        }
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(name: &str, entry_type: EntryType, child: u32, left: u32, right: u32) -> DirEntry {
        DirEntry {
            name: name.to_string(),
            entry_type,
            color: 1,
            left_sibling: left,
            right_sibling: right,
            child,
            start_sector: 0,
            stream_size: 0,
        }
    }

    /// A minimal but realistic tree: Root -> ObjectPool -> one embedded
    /// object storage -> {CompObj, Workbook} streams.
    fn build_tree_with_one_excel_object() -> Vec<DirEntry> {
        vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY), // 0
            entry("ObjectPool", EntryType::Storage, 2, NO_ENTRY, NO_ENTRY),     // 1
            entry("_1234567890", EntryType::Storage, 3, NO_ENTRY, NO_ENTRY),    // 2
            entry("\u{1}CompObj", EntryType::Stream, NO_ENTRY, NO_ENTRY, 4),    // 3
            entry("Workbook", EntryType::Stream, NO_ENTRY, NO_ENTRY, NO_ENTRY), // 4
        ]
    }

    #[test]
    fn test_finds_one_embedded_excel_workbook() {
        let entries = build_tree_with_one_excel_object();
        let (out, _) = classify_objects(&entries, MAX_ENTRIES_WALKED);
        assert_eq!(
            out,
            vec![(
                EmbeddedOleObject {
                    description: "Embedded Microsoft Excel Workbook".to_string(),
                    text: None,
                },
                2
            )]
        );
    }

    #[test]
    fn test_no_object_pool_yields_empty() {
        let entries = vec![entry(
            "Root Entry",
            EntryType::RootStorage,
            NO_ENTRY,
            NO_ENTRY,
            NO_ENTRY,
        )];
        assert!(find_direct_child(&entries, entries[0].child, "ObjectPool").is_none());
    }

    #[test]
    fn test_classify_recognizes_every_known_type() {
        assert_eq!(classify(&["Workbook"]).as_deref(), Some("Embedded Microsoft Excel Workbook"));
        assert_eq!(
            classify(&["Equation Native"]).as_deref(),
            Some("Embedded Equation Editor/MathType Object")
        );
        assert_eq!(
            classify(&["WordDocument"]).as_deref(),
            Some("Embedded Microsoft Word Document")
        );
        assert_eq!(
            classify(&["PowerPoint Document"]).as_deref(),
            Some("Embedded Microsoft PowerPoint Presentation")
        );
        assert_eq!(classify(&["Package"]).as_deref(), Some("Embedded OLE Package"));
        assert_eq!(classify(&["\u{1}CompObj"]).as_deref(), Some("Embedded OLE Object"));
        assert_eq!(classify(&["SomethingUnknown"]), None);
        assert_eq!(classify(&[]), None);
    }

    /// A sibling cycle (corrupted/adversarial file) must not hang.
    #[test]
    fn test_cyclic_siblings_do_not_hang() {
        let entries = vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY),
            // node 1's right sibling points back to itself.
            entry("Loop", EntryType::Storage, NO_ENTRY, NO_ENTRY, 1),
        ];
        let mut walker = Walker::new(&entries, MAX_ENTRIES_WALKED);
        assert_eq!(walker.collect(entries[0].child), vec![1]);
        assert_eq!(walker.steps, 1, "a cyclic entry is taken once, not once per step");
    }

    /// Object storages that loop on themselves and share one looping child
    /// tree: the output is the one object and the walk takes each entry once.
    #[test]
    fn test_self_looping_object_pool_walks_each_entry_once() {
        let n: u32 = 200;
        let mut entries = vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY),
            entry("ObjectPool", EntryType::Storage, 2, NO_ENTRY, NO_ENTRY),
        ];
        let stream = 2 + n;
        for i in 0..n {
            entries.push(entry(&format!("_{i}"), EntryType::Storage, stream, 2 + i, 2 + i));
        }
        entries.push(entry("Workbook", EntryType::Stream, NO_ENTRY, stream, stream));
        let (out, steps) = classify_objects(&entries, 8);
        assert_eq!(out.len(), 1);
        assert!(steps <= entries.len(), "steps {steps} exceed the {} entries", entries.len());
    }

    /// One walk stops at the per-walk cap, however long the sibling chain.
    #[test]
    fn test_one_walk_stops_at_the_per_walk_cap() {
        let len: u32 = 100;
        let mut entries = vec![entry(
            "Root Entry",
            EntryType::RootStorage,
            1,
            NO_ENTRY,
            NO_ENTRY,
        )];
        for i in 0..len {
            let right = if i + 1 < len { 2 + i } else { NO_ENTRY };
            entries.push(entry(&format!("s{i}"), EntryType::Stream, NO_ENTRY, NO_ENTRY, right));
        }
        let mut walker = Walker::new(&entries, 5);
        assert_eq!(walker.collect(1).len(), 5);
        assert_eq!(walker.steps, 5, "the walk went past its cap");
    }

    /// Many object storages whose child links all land in one large
    /// sibling chain: without a pass-wide visited set each walks the rest of
    /// the chain (objects x entries); with one, the chain is walked once in
    /// total and the work stays within the entry count, however small the
    /// per-walk cap.
    #[test]
    fn test_many_objects_sharing_one_large_subtree_cost_one_walk() {
        let objects: u32 = 400;
        let chain: u32 = 400;
        // Pool siblings: object i's right sibling is object i + 1.
        let first_obj = 2;
        let chain_start = first_obj + objects;
        let mut entries = vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY),
            entry("ObjectPool", EntryType::Storage, first_obj, NO_ENTRY, NO_ENTRY),
        ];
        for i in 0..objects {
            let right = if i + 1 < objects {
                first_obj + i + 1
            } else {
                NO_ENTRY
            };
            // Every object's child is a different entry of the same chain.
            entries.push(entry(
                &format!("_{i}"),
                EntryType::Storage,
                chain_start + (i % chain),
                NO_ENTRY,
                right,
            ));
        }
        for j in 0..chain {
            let right = if j + 1 < chain {
                chain_start + j + 1
            } else {
                NO_ENTRY
            };
            let name = if j == chain - 1 { "Workbook" } else { "x" };
            entries.push(entry(name, EntryType::Stream, NO_ENTRY, NO_ENTRY, right));
        }
        let (out, steps) = classify_objects(&entries, 50);
        assert!(
            steps <= chain as usize,
            "the shared chain was walked {steps} steps for {chain} entries"
        );
        assert!(out.len() <= 1, "an entry is claimed by one object only");
    }
}
