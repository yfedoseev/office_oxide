//! A package embedded in a `.docx` (`<o:OLEObject Type="Embed">`) is opened to
//! fold its text into the document's, and it can embed packages of its own.
//! What that costs is charged to the one document that holds it: the depth is
//! bounded, a blob reached through several entry names is opened once, the
//! packages opened and the bytes they take are capped, and a nested open
//! never waits for a parse slot (a chain as long as the pool of slots, or as
//! many documents embedding one package as the pool has slots, would
//! otherwise never finish).
//!
//! The limits are process-global, so every test holds the `LimitScope` (or the
//! default limits), as the other limit tests do.

mod common;

use std::io::Cursor;
use std::sync::mpsc;
use std::time::{Duration, Instant};

use common::LimitScope;
use office_oxide::limits;
use office_oxide::{Document, DocumentFormat};

const WNS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml""#;

fn crc32(data: &[u8]) -> u32 {
    let mut crc = flate2::Crc::new();
    crc.update(data);
    crc.sum()
}

/// A zip of stored `blobs`, with `entries` naming them: several entries may
/// name one blob, and then share its local header and data, as a hand-built
/// archive can (APPNOTE 4.3.7, 4.3.12).
fn zip_with_aliases(blobs: &[Vec<u8>], entries: &[(String, usize)]) -> Vec<u8> {
    let mut out = Vec::new();
    let mut offsets = vec![usize::MAX; blobs.len()];
    let mut first_name: Vec<Option<&str>> = vec![None; blobs.len()];
    for (name, blob) in entries {
        first_name[*blob].get_or_insert(name);
    }
    for (i, blob) in blobs.iter().enumerate() {
        let Some(name) = first_name[i] else { continue };
        offsets[i] = out.len();
        out.extend_from_slice(&0x0403_4b50u32.to_le_bytes());
        out.extend_from_slice(&[20, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        out.extend_from_slice(&crc32(blob).to_le_bytes());
        out.extend_from_slice(&(blob.len() as u32).to_le_bytes());
        out.extend_from_slice(&(blob.len() as u32).to_le_bytes());
        out.extend_from_slice(&(name.len() as u16).to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(name.as_bytes());
        out.extend_from_slice(blob);
    }
    let directory = out.len();
    for (name, blob) in entries {
        let data = &blobs[*blob];
        out.extend_from_slice(&0x0201_4b50u32.to_le_bytes());
        out.extend_from_slice(&[20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        out.extend_from_slice(&crc32(data).to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(&(name.len() as u16).to_le_bytes());
        out.extend_from_slice(&[0; 12]);
        out.extend_from_slice(&(offsets[*blob] as u32).to_le_bytes());
        out.extend_from_slice(name.as_bytes());
    }
    let size = out.len() - directory;
    out.extend_from_slice(&0x0605_4b50u32.to_le_bytes());
    out.extend_from_slice(&[0; 4]);
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    out.extend_from_slice(&(entries.len() as u16).to_le_bytes());
    out.extend_from_slice(&(size as u32).to_le_bytes());
    out.extend_from_slice(&(directory as u32).to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out
}

/// A `.docx` whose body says `text` and embeds each of `inner` as a package;
/// the blob `inner[i]` is named by `copies[i]` entries (aliases of one
/// another).
fn docx_embedding(text: &str, inner: &[(&Vec<u8>, usize)]) -> Vec<u8> {
    let mut body = format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
    let mut rels = String::new();
    let mut blobs: Vec<Vec<u8>> = vec![
        br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
    ];
    let mut entries: Vec<(String, usize)> =
        vec![("[Content_Types].xml".into(), 0), ("_rels/.rels".into(), 1)];
    let mut id = 0;
    for (blob, copies) in inner {
        let index = blobs.len();
        blobs.push((*blob).clone());
        for _ in 0..*copies {
            body.push_str(&format!(
                r#"<w:p><w:r><w:object><v:shape/><o:OLEObject Type="Embed" ProgID="Word.Document.12" r:id="rE{id}"/></w:object></w:r></w:p>"#
            ));
            rels.push_str(&format!(
                r#"<Relationship Id="rE{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/e{id}.docx"/>"#
            ));
            entries.push((format!("word/embeddings/e{id}.docx"), index));
            id += 1;
        }
    }
    let document = blobs.len();
    blobs.push(format!(r#"<w:document {WNS}><w:body>{body}</w:body></w:document>"#).into_bytes());
    blobs.push(
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        )
        .into_bytes(),
    );
    entries.push(("word/document.xml".into(), document));
    entries.push(("word/_rels/document.xml.rels".into(), document + 1));
    zip_with_aliases(&blobs, &entries)
}

fn leaf(text: &str) -> Vec<u8> {
    docx_embedding(text, &[])
}

/// A tree of `depth` levels in which each package embeds `fan_out` aliased
/// entries naming the package below it; level `n` says `L{n}`.
fn aliased_tree(depth: usize, fan_out: usize) -> Vec<u8> {
    let mut current = leaf(&format!("L{depth}"));
    for level in (0..depth).rev() {
        current = docx_embedding(&format!("L{level}"), &[(&current, fan_out)]);
    }
    current
}

/// Run `f` on a thread and give it `secs` to finish.
fn within<T: Send + 'static>(secs: u64, f: impl FnOnce() -> T + Send + 'static) -> T {
    let (tx, rx) = mpsc::channel();
    std::thread::spawn(move || {
        let _ = tx.send(f());
    });
    rx.recv_timeout(Duration::from_secs(secs))
        .expect("did not finish in time")
}

fn open(bytes: Vec<u8>) -> Document {
    Document::from_reader(Cursor::new(bytes), DocumentFormat::Docx).unwrap()
}

fn skipped(doc: &Document) -> Vec<String> {
    doc.as_docx()
        .unwrap()
        .unreadable_parts
        .iter()
        .filter(|p| p.contains("skipped embedded package"))
        .cloned()
        .collect()
}

#[test]
fn test_a_long_chain_of_embedded_packages_is_cut_at_the_depth_limit() {
    let _limits = common::default_limits();
    // Longer than any pool of parse slots.
    let chain = aliased_tree(40, 1);
    let (doc, took) = within(60, move || {
        let started = Instant::now();
        let doc = open(chain);
        let _ = doc.plain_text();
        (doc, started.elapsed())
    });
    assert!(took < Duration::from_secs(30), "{took:?}");
    let text = doc.plain_text();
    // The document and the three levels below it are read; the fourth is not.
    assert_eq!(limits::DEFAULT_MAX_EMBEDDED_DEPTH, 3);
    assert!(text.contains("L3"), "{text}");
    assert!(!text.contains("L4"), "{text}");
    assert_eq!(skipped(&doc).len(), 1, "{:?}", skipped(&doc));
    assert!(doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_embedded_packages_that_are_aliases_of_one_blob_are_opened_once() {
    let mut scope = LimitScope::new();
    // Ten entries name each level's package, five levels down. Opened once
    // per blob, five packages are opened; opened once per entry, 10^5.
    scope.set(limits::max_embedded_depth, limits::set_max_embedded_depth, 10);
    scope.set(limits::max_embedded_packages, limits::set_max_embedded_packages, 5);
    let tree = aliased_tree(5, 10);
    let doc = within(60, move || open(tree));
    assert!(skipped(&doc).is_empty(), "{:?}", skipped(&doc));
    let text = doc.plain_text();
    assert!(text.contains("L5"), "{text}");
    assert!(!doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_a_deep_aliased_tree_is_bounded_in_time_by_default() {
    let _limits = common::default_limits();
    let tree = aliased_tree(5, 10);
    let started = Instant::now();
    let doc = within(60, move || open(tree));
    assert!(started.elapsed() < Duration::from_secs(20), "{:?}", started.elapsed());
    // Each of the ten entries naming the package below the limit is
    // a part that was not opened, and is reported.
    assert_eq!(skipped(&doc).len(), 10, "{:?}", skipped(&doc));
    assert!(skipped(&doc).iter().all(|p| p.contains("levels deep")));
}

#[test]
fn test_distinct_embedded_packages_are_capped_in_number() {
    let mut scope = LimitScope::new();
    scope.set(limits::max_embedded_packages, limits::set_max_embedded_packages, 4);
    let leaves: Vec<Vec<u8>> = (0..10).map(|i| leaf(&format!("child{i}x"))).collect();
    let inner: Vec<(&Vec<u8>, usize)> = leaves.iter().map(|l| (l, 1)).collect();
    let doc = open(docx_embedding("top", &inner));
    let text = doc.plain_text();
    let found = (0..10)
        .filter(|i| text.contains(&format!("child{i}x")))
        .count();
    assert_eq!(found, 4, "{text}");
    assert_eq!(skipped(&doc).len(), 6, "{:?}", skipped(&doc));
}

#[test]
fn test_embedded_packages_share_one_byte_allowance() {
    let mut scope = LimitScope::new();
    let leaves: Vec<Vec<u8>> = (0..4).map(|i| leaf(&format!("child{i}x"))).collect();
    // Room for two of the packages and not a third.
    let each = leaves[0].len() as u64;
    scope.set(limits::max_embedded_bytes, limits::set_max_embedded_bytes, each * 2 + each / 2);
    let inner: Vec<(&Vec<u8>, usize)> = leaves.iter().map(|l| (l, 1)).collect();
    let doc = open(docx_embedding("top", &inner));
    let text = doc.plain_text();
    let found = (0..4)
        .filter(|i| text.contains(&format!("child{i}x")))
        .count();
    assert!(found < 4, "{text}");
    assert!(!skipped(&doc).is_empty());
}

#[test]
fn test_embedded_packages_share_the_text_budget() {
    let mut scope = LimitScope::new();
    let leaves: Vec<Vec<u8>> = (0..6)
        .map(|i| leaf(&format!("child{i}x{}", "y".repeat(50))))
        .collect();
    scope.set(limits::max_text_chars, limits::set_max_text_chars, 130);
    let inner: Vec<(&Vec<u8>, usize)> = leaves.iter().map(|l| (l, 1)).collect();
    let doc = open(docx_embedding("top", &inner));
    let found = (0..6)
        .filter(|i| doc.plain_text().contains(&format!("child{i}x")))
        .count();
    assert!(found < 6, "every embedded package was copied: {found}");
    assert!(!skipped(&doc).is_empty());
}

#[test]
fn test_a_document_that_embeds_one_package_still_yields_its_text() {
    let _limits = common::default_limits();
    let doc = open(docx_embedding("outer", &[(&leaf("inner text"), 1)]));
    let text = doc.plain_text();
    assert!(text.contains("outer") && text.contains("inner text"), "{text}");
    assert!(skipped(&doc).is_empty());
    assert!(!doc.to_ir().metadata.text_truncated);
}

#[test]
fn test_many_documents_embedding_a_package_do_not_wait_for_each_other() {
    let _limits = common::default_limits();
    // More simultaneous parses than the pool of parse slots can hold (at
    // most 64), each long enough to parse that they all hold a slot before
    // any reaches its embedded package.
    const PARSES: usize = 200;
    let filler = "</w:t></w:r></w:p><w:p><w:r><w:t>pad".repeat(4_000);
    let bytes = docx_embedding(&format!("outer{filler}"), &[(&leaf("inner text"), 1)]);
    let (tx, rx) = mpsc::channel();
    for _ in 0..PARSES {
        let (tx, bytes) = (tx.clone(), bytes.clone());
        std::thread::spawn(move || {
            let doc = open(bytes);
            let _ = tx.send(doc.plain_text().contains("inner text"));
        });
    }
    drop(tx);
    for _ in 0..PARSES {
        assert!(
            rx.recv_timeout(Duration::from_secs(120))
                .expect("a parse is stuck")
        );
    }
}

#[test]
fn test_what_an_embedded_package_could_not_read_is_reported_by_the_document_that_holds_it() {
    let _limits = common::default_limits();
    // The inner package's body is malformed after its first paragraph.
    let inner = leaf(r#"kept inner</w:t></w:r></w:p><w:p><w:r a=""#);
    let doc = open(docx_embedding("outer", &[(&inner, 1)]));
    let text = doc.plain_text();
    assert!(text.contains("kept inner"), "{text}");
    let reported = &doc.as_docx().unwrap().unreadable_parts;
    assert!(
        reported
            .iter()
            .any(|p| p.contains("embeddings/e0.docx") && p.contains("word/document.xml")),
        "{reported:?}"
    );
    // A sound inner package reports nothing.
    let sound = open(docx_embedding("outer", &[(&leaf("fine"), 1)]));
    assert!(sound.as_docx().unwrap().unreadable_parts.is_empty());
}

/// A package reached through several entry names is charged to the byte
/// allowance and to the package count once, not once per name: five names for
/// one blob fit an allowance of two and a half packages (five charges would
/// need five), and all five places in the document carry its text.
#[test]
fn test_one_blob_under_several_names_is_charged_once() {
    let mut scope = LimitScope::new();
    let blob = leaf("shared child");
    scope.set(
        limits::max_embedded_bytes,
        limits::set_max_embedded_bytes,
        blob.len() as u64 * 5 / 2,
    );
    scope.set(limits::max_embedded_packages, limits::set_max_embedded_packages, 1);
    let doc = open(docx_embedding("top", &[(&blob, 5)]));
    assert!(skipped(&doc).is_empty(), "{:?}", skipped(&doc));
    assert_eq!(doc.plain_text().matches("shared child").count(), 5);
}
