//! Embedded OLE object recognition (Excel workbooks, Equation Editor/
//! MathType, OLE Package, embedded Word/PowerPoint, etc.) inside a
//! DOC's `Storage:ObjectPool` storage.
//!
//! At minimum (issue #284): recognize that an embedded object exists
//! and surface its identity, even without extracting its native
//! payload — the DOC analogue of the PPT fix in issue #337. Before
//! this, `ObjectPool` was never traversed at all, so a file with (for
//! example) 12 embedded Excel workbooks showed no trace of any of them
//! anywhere in the IR.

use std::io::{Read, Seek};

use crate::cfb::{CfbReader, DirEntry, EntryType};

/// Sentinel meaning "no sibling/child" in a CFB directory entry, mirrors
/// `cfb::directory::NO_ENTRY` (not reachable from here — that module is
/// private to `cfb`).
const NO_ENTRY: u32 = 0xFFFF_FFFF;

/// Bound on how many directory entries a single walk may visit —
/// `ObjectPool`'s own storage tree comes from an untrusted file, so a
/// corrupted/adversarial sibling chain must not spin unboundedly. Well
/// above any real document's actual embedded-object count.
const MAX_ENTRIES_WALKED: usize = 50_000;

/// One embedded OLE object found under `ObjectPool`, identified only by
/// the presence of a well-known stream name within its own storage
/// (issue #284).
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
    let entries = reader.entries();
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
    for obj_idx in collect_children(entries, pool.child) {
        let obj = &entries[obj_idx];
        if obj.entry_type != EntryType::Storage {
            continue; // a stray non-storage entry directly under ObjectPool
        }
        let children = collect_children(entries, obj.child);
        let names: Vec<&str> = children.iter().map(|&i| entries[i].name.as_str()).collect();
        if let Some(description) = classify(&names) {
            out.push(EmbeddedOleObject { description });
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
/// case-insensitively. Explicit-stack and entry-count-bounded — the
/// tree comes from an untrusted file, so a corrupted/adversarial
/// sibling cycle must not overflow the stack or loop unboundedly.
fn find_direct_child(entries: &[DirEntry], root: u32, name: &str) -> Option<usize> {
    let mut stack = vec![root];
    let mut visited = 0usize;
    while let Some(node) = stack.pop() {
        visited += 1;
        if visited > MAX_ENTRIES_WALKED || node == NO_ENTRY || node as usize >= entries.len() {
            continue;
        }
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
/// explicit-stack and entry-count-bounded for the same reason as
/// [`find_direct_child`].
fn collect_children(entries: &[DirEntry], root: u32) -> Vec<usize> {
    let mut out = Vec::new();
    let mut stack = vec![root];
    let mut visited = 0usize;
    while let Some(node) = stack.pop() {
        visited += 1;
        if visited > MAX_ENTRIES_WALKED || node == NO_ENTRY || node as usize >= entries.len() {
            continue;
        }
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
    fn finds_one_embedded_excel_workbook() {
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
    fn no_object_pool_yields_empty() {
        let entries = vec![entry("Root Entry", EntryType::RootStorage, NO_ENTRY, NO_ENTRY, NO_ENTRY)];
        assert!(find_direct_child(&entries, entries[0].child, "ObjectPool").is_none());
    }

    #[test]
    fn classify_recognizes_every_known_type() {
        assert_eq!(classify(&["Workbook"]).as_deref(), Some("Embedded Microsoft Excel Workbook"));
        assert_eq!(
            classify(&["Equation Native"]).as_deref(),
            Some("Embedded Equation Editor/MathType Object")
        );
        assert_eq!(classify(&["WordDocument"]).as_deref(), Some("Embedded Microsoft Word Document"));
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
    fn cyclic_siblings_do_not_hang() {
        let entries = vec![
            entry("Root Entry", EntryType::RootStorage, 1, NO_ENTRY, NO_ENTRY),
            // node 1's right sibling points back to itself.
            entry("Loop", EntryType::Storage, NO_ENTRY, NO_ENTRY, 1),
        ];
        let out = collect_children(&entries, entries[0].child);
        assert!(!out.is_empty());
    }
}
