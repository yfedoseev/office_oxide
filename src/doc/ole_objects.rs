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

use std::collections::HashMap;
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
}

/// Find every embedded object under the root's `ObjectPool` storage (if
/// present) and classify each by the presence of a well-known stream
/// name inside its own storage. Returns an empty `Vec` (not an error)
/// when there's no `ObjectPool` at all, or its structure doesn't match
/// expectations — this is a best-effort enhancement, never a hard
/// requirement for reading the rest of the document.
pub fn extract_ole_objects<R: Read + Seek>(reader: &CfbReader<R>) -> Vec<EmbeddedOleObject> {
    classify_objects(reader.entries())
}

/// The walk behind [`extract_ole_objects`], over a bare directory.
fn classify_objects(entries: &[DirEntry]) -> Vec<EmbeddedOleObject> {
    if entries.is_empty() || entries[0].entry_type != EntryType::RootStorage {
        return Vec::new();
    }
    let Some(pool_idx) = find_direct_child(entries, entries[0].child, "ObjectPool") else {
        return Vec::new();
    };
    let pool = &entries[pool_idx];
    if !matches!(pool.entry_type, EntryType::Storage | EntryType::RootStorage) {
        return Vec::new();
    }

    let mut out = Vec::new();
    // Many object storages can name the same child tree; classify each
    // distinct tree once rather than walking it again for every object.
    let mut classified: HashMap<u32, Option<String>> = HashMap::new();
    for obj_idx in collect_children(entries, pool.child) {
        let obj = &entries[obj_idx];
        if obj.entry_type != EntryType::Storage {
            continue; // a stray non-storage entry directly under ObjectPool
        }
        let description = classified.entry(obj.child).or_insert_with(|| {
            let children = collect_children(entries, obj.child);
            let names: Vec<&str> = children.iter().map(|&i| entries[i].name.as_str()).collect();
            classify(&names)
        });
        if let Some(description) = description {
            out.push(EmbeddedOleObject {
                description: description.clone(),
            });
        }
    }
    out
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
        Some("Embedded Microsoft Word Document".to_string())
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
fn find_direct_child(entries: &[DirEntry], root: u32, name: &str) -> Option<usize> {
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

/// All directory entry indices in the RB tree rooted at `root`,
/// explicit-stack with a visited set for the same reason as
/// [`find_direct_child`]: each entry appears once, so the walk is bounded
/// by the directory however its pointers are wired.
fn collect_children(entries: &[DirEntry], root: u32) -> Vec<usize> {
    let mut out = Vec::new();
    let mut stack = vec![root];
    let mut seen = vec![false; entries.len()];
    while let Some(node) = stack.pop() {
        if node == NO_ENTRY || node as usize >= entries.len() || seen[node as usize] {
            continue;
        }
        seen[node as usize] = true;
        let entry = &entries[node as usize];
        out.push(node as usize);
        stack.push(entry.left_sibling);
        stack.push(entry.right_sibling);
    }
    out
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
        // Exercise the free functions directly — CfbReader construction
        // needs real sector data, which this test isn't building.
        let pool_idx = find_direct_child(&entries, entries[0].child, "ObjectPool").unwrap();
        let pool = &entries[pool_idx];
        let mut out = Vec::new();
        for obj_idx in collect_children(&entries, pool.child) {
            let obj = &entries[obj_idx];
            if obj.entry_type != EntryType::Storage {
                continue;
            }
            let children = collect_children(&entries, obj.child);
            let names: Vec<&str> = children.iter().map(|&i| entries[i].name.as_str()).collect();
            if let Some(description) = classify(&names) {
                out.push(description);
            }
        }
        assert_eq!(out, vec!["Embedded Microsoft Excel Workbook".to_string()]);
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
        let out = collect_children(&entries, entries[0].child);
        assert_eq!(out, vec![1], "a cyclic entry is reported once, not once per step");
    }

    /// Object storages whose sibling pointers loop on themselves, and a
    /// pool of many such objects, cost a walk bounded by the directory
    /// (every entry once), not by a per-walk step cap times the objects.
    #[test]
    fn test_self_looping_object_pool_walks_each_entry_once() {
        let n: u32 = 200;
        let mut entries = vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY),
            entry("ObjectPool", EntryType::Storage, 2, NO_ENTRY, NO_ENTRY),
        ];
        // Objects 2..2+n: each is its own left and right sibling, and its
        // child tree is the shared looping stream below.
        let stream = 2 + n;
        for i in 0..n {
            entries.push(entry(&format!("_{i}"), EntryType::Storage, stream, 2 + i, 2 + i));
        }
        entries.push(entry("Workbook", EntryType::Stream, NO_ENTRY, stream, stream));
        let pool = collect_children(&entries, entries[1].child);
        assert_eq!(pool.len(), 1, "the self-looping object is listed once");
        assert_eq!(classify_objects(&entries).len(), 1);
    }
}
