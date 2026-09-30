//! DOCX editing via raw XML text replacement.
//!
//! Uses the `EditablePackage` from core to preserve all parts, replacing
//! text in the `<w:t>` elements of the body and of the headers, footers,
//! footnotes, endnotes and comments it references. A match may span the
//! runs of one paragraph; everything else in each part is kept byte for
//! byte.

use crate::core::editable::EditablePackage;
use crate::core::opc::PartName;

use super::Result;

/// An editable DOCX document that supports text replacement and saving.
pub struct EditableDocx {
    package: EditablePackage,
    main_part: PartName,
}

impl EditableDocx {
    /// Open a DOCX file for editing.
    pub fn open(path: impl AsRef<std::path::Path>) -> Result<Self> {
        let package = EditablePackage::open(&path)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Open from any `Read + Seek` source.
    pub fn from_reader<R: std::io::Read + std::io::Seek>(reader: R) -> Result<Self> {
        let package = EditablePackage::from_reader(reader)?;
        let main_part = PartName::new("/word/document.xml")?;
        Ok(Self { package, main_part })
    }

    /// Replace all occurrences of `find` with `replace` in the document's
    /// text: the body, and the headers, footers, footnotes, endnotes and
    /// comments it references. A match may span runs (Word splits a phrase
    /// at every formatting, spelling or revision boundary) but not
    /// paragraphs; the replacement takes the formatting of the run the
    /// match starts in. Returns the number of replacements made.
    pub fn replace_text(&mut self, find: &str, replace: &str) -> usize {
        let mut parts = vec![self.main_part.clone()];
        if let Some(rels) = self.package.part_rels(&self.main_part) {
            use crate::core::relationships::{TargetMode, rel_types};
            for rel in rels.all() {
                let text_part = [
                    rel_types::HEADER,
                    rel_types::FOOTER,
                    rel_types::FOOTNOTES,
                    rel_types::ENDNOTES,
                    rel_types::COMMENTS,
                ]
                .contains(&rel.rel_type.as_str());
                if text_part && rel.target_mode == TargetMode::Internal {
                    if let Ok(part) = self.main_part.resolve_relative(&rel.target) {
                        if !parts.contains(&part) {
                            parts.push(part);
                        }
                    }
                }
            }
        }

        let mut total = 0;
        for part in parts {
            let Some(data) = self.package.get_part(&part) else {
                continue;
            };
            let xml_str = String::from_utf8_lossy(data);
            let (new_xml, count) = replace_across_runs(&xml_str, find, replace);
            if count > 0 {
                self.package.set_part(part, new_xml.into_bytes());
                total += count;
            }
        }
        total
    }

    /// Save the edited document to a file.
    pub fn save(&self, path: impl AsRef<std::path::Path>) -> Result<()> {
        self.package.save(path)?;
        Ok(())
    }

    /// Write the edited document to any `Write + Seek` destination.
    pub fn write_to<W: std::io::Write + std::io::Seek>(&self, writer: W) -> Result<()> {
        self.package.write_to(writer)?;
        Ok(())
    }
}

/// One `<w:t>` text node: where its tag and content sit in the source, its
/// decoded text, and the paragraph it belongs to.
struct TextNode {
    tag_start: usize,
    tag_end: usize,
    content_end: usize,
    text: String,
    paragraph: usize,
}

/// Replace `find` with `replace` in the text of each paragraph of a
/// WordprocessingML part, matching across the `<w:t>` nodes of the
/// paragraph's runs. A node's text belongs to its innermost enclosing
/// `<w:p>` (a text box's paragraphs nest inside the anchoring one), and a
/// match never spans paragraphs. The replacement goes into the node where
/// the match starts; the rest of the matched text is removed from the
/// nodes it covered. Nodes that do not change keep their source bytes.
/// Returns the new XML and the number of replacements.
fn replace_across_runs(xml: &str, find: &str, replace: &str) -> (String, usize) {
    if find.is_empty() {
        return (xml.to_string(), 0);
    }
    let nodes = scan_text_nodes(xml);

    // Group the nodes by paragraph, keeping document order.
    let mut groups: Vec<Vec<usize>> = Vec::new();
    let mut group_of: std::collections::HashMap<usize, usize> = std::collections::HashMap::new();
    for (i, n) in nodes.iter().enumerate() {
        let g = *group_of.entry(n.paragraph).or_insert_with(|| {
            groups.push(Vec::new());
            groups.len() - 1
        });
        groups[g].push(i);
    }

    let mut new_text: Vec<Option<String>> = vec![None; nodes.len()];
    let mut count = 0;
    for group in &groups {
        let joined: String = group.iter().map(|&i| nodes[i].text.as_str()).collect();
        let matches: Vec<(usize, usize)> = joined
            .match_indices(find)
            .map(|(s, m)| (s, s + m.len()))
            .collect();
        if matches.is_empty() {
            continue;
        }
        count += matches.len();
        // Each node's span in `joined`.
        let mut spans = Vec::with_capacity(group.len());
        let mut at = 0;
        for &i in group {
            spans.push((at, at + nodes[i].text.len()));
            at += nodes[i].text.len();
        }
        let mut out: Vec<String> = vec![String::new(); group.len()];
        // Copy `joined[a..b]` into the nodes that hold it.
        let copy = |out: &mut Vec<String>, a: usize, b: usize| {
            for (k, &(s, e)) in spans.iter().enumerate() {
                let (lo, hi) = (a.max(s), b.min(e));
                if lo < hi {
                    out[k].push_str(&joined[lo..hi]);
                }
            }
        };
        let mut cur = 0;
        for &(s, e) in &matches {
            copy(&mut out, cur, s);
            // The node the match starts in (`s < e`, so it holds `s`).
            let k = spans
                .iter()
                .position(|&(a, b)| a <= s && s < b)
                .unwrap_or(0);
            out[k].push_str(replace);
            cur = e;
        }
        copy(&mut out, cur, joined.len());
        for (k, &i) in group.iter().enumerate() {
            if out[k] != nodes[i].text {
                new_text[i] = Some(std::mem::take(&mut out[k]));
            }
        }
    }
    if count == 0 {
        return (xml.to_string(), 0);
    }

    let mut result = String::with_capacity(xml.len());
    let mut pos = 0;
    for (n, text) in nodes.iter().zip(&new_text) {
        let Some(text) = text else { continue };
        let tag = &xml[n.tag_start..n.tag_end];
        result.push_str(&xml[pos..n.tag_start]);
        // Text that now starts or ends with whitespace needs
        // `xml:space="preserve"`, or Word drops that whitespace.
        let edge_space =
            text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace);
        if edge_space && !tag.contains("xml:space") {
            result.push_str(&tag[..tag.len() - 1]);
            result.push_str(r#" xml:space="preserve">"#);
        } else {
            result.push_str(tag);
        }
        result.push_str(&quick_xml::escape::escape(text.as_str()));
        pos = n.content_end;
    }
    result.push_str(&xml[pos..]);
    (result, count)
}

/// Every non-empty-element `<w:t>` in `xml`, with the paragraph it belongs
/// to (`usize::MAX` for text outside any paragraph).
fn scan_text_nodes(xml: &str) -> Vec<TextNode> {
    let mut nodes = Vec::new();
    // Open paragraphs, innermost last, each with its own id.
    let mut open: Vec<usize> = Vec::new();
    let mut next_id = 0usize;
    let mut pos = 0;
    while let Some(off) = xml[pos..].find('<') {
        let lt = pos + off;
        let rest = &xml[lt..];
        // Comments, CDATA and processing instructions carry no markup.
        let skip_to = if rest.starts_with("<!--") {
            Some("-->")
        } else if rest.starts_with("<![CDATA[") {
            Some("]]>")
        } else if rest.starts_with("<?") {
            Some("?>")
        } else {
            None
        };
        if let Some(end) = skip_to {
            pos = rest.find(end).map_or(xml.len(), |i| lt + i + end.len());
            continue;
        }
        let Some(gt) = rest.find('>') else { break };
        let tag_end = lt + gt + 1;
        let tag = &xml[lt..tag_end];
        let (closing, body) = match tag[1..].strip_prefix('/') {
            Some(body) => (true, body),
            None => (false, &tag[1..]),
        };
        let name = &body[..body
            .find(|c: char| c.is_whitespace() || c == '>' || c == '/')
            .unwrap_or(body.len())];
        let self_closing = tag.ends_with("/>");
        match name {
            "w:p" if closing => {
                open.pop();
            },
            "w:p" if !self_closing => {
                open.push(next_id);
                next_id += 1;
            },
            "w:t" if !closing && !self_closing => {
                let Some(close) = xml[tag_end..].find("</w:t>") else {
                    break;
                };
                let content_end = tag_end + close;
                let raw = &xml[tag_end..content_end];
                let text = quick_xml::escape::unescape(raw)
                    .map(|c| c.into_owned())
                    .unwrap_or_else(|_| raw.to_string());
                nodes.push(TextNode {
                    tag_start: lt,
                    tag_end,
                    content_end,
                    text,
                    paragraph: open.last().copied().unwrap_or(usize::MAX),
                });
                pos = content_end + "</w:t>".len();
                continue;
            },
            _ => {},
        }
        pos = tag_end;
    }
    nodes
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_replace_in_wt_simple() {
        let xml = r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#;
        let (result, count) = replace_across_runs(xml, "World", "Rust");
        assert_eq!(count, 1);
        assert!(result.contains("<w:t>Hello Rust</w:t>"));
    }

    #[test]
    fn test_replace_in_wt_multiple() {
        let xml = r#"<w:r><w:t>foo bar foo</w:t></w:r>"#;
        let (result, count) = replace_across_runs(xml, "foo", "baz");
        assert_eq!(count, 2);
        assert!(result.contains("<w:t>baz bar baz</w:t>"));
    }

    #[test]
    fn test_replace_preserves_attributes() {
        let xml = r#"<w:r><w:t xml:space="preserve"> Hello </w:t></w:r>"#;
        let (result, count) = replace_across_runs(xml, "Hello", "World");
        assert_eq!(count, 1);
        assert!(result.contains(r#"xml:space="preserve"> World </w:t>"#));
    }

    /// Word splits one visible phrase across runs at every formatting,
    /// spelling or revision boundary. Matching inside each `<w:t>` alone
    /// found nothing and reported 0 for text plainly in the document.
    #[test]
    fn test_replace_matches_text_split_across_runs() {
        let xml = r#"<w:p><w:r><w:t>Say Hel</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>lo Wor</w:t></w:r><w:r><w:t>ld!</w:t></w:r></w:p>"#;
        let (result, count) = replace_across_runs(xml, "Hello World", "Bye");
        assert_eq!(count, 1);
        assert_eq!(
            result,
            r#"<w:p><w:r><w:t>Say Bye</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r><w:r><w:t>!</w:t></w:r></w:p>"#
        );
    }

    #[test]
    fn test_replace_does_not_join_text_across_paragraphs() {
        let xml = r#"<w:p><w:r><w:t>Hel</w:t></w:r></w:p><w:p><w:r><w:t>lo</w:t></w:r></w:p>"#;
        let (result, count) = replace_across_runs(xml, "Hello", "X");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }

    /// A text box's paragraphs nest inside the anchoring paragraph; each
    /// is matched on its own text only.
    #[test]
    fn test_replace_keeps_nested_paragraphs_apart() {
        let xml = concat!(
            r#"<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:pict><w:txbxContent>"#,
            r#"<w:p><w:r><w:t>cd</w:t></w:r></w:p></w:txbxContent></w:pict></w:r>"#,
            r#"<w:r><w:t>ef</w:t></w:r></w:p>"#
        );
        assert_eq!(replace_across_runs(xml, "bc", "X").1, 0);
        let (result, count) = replace_across_runs(xml, "be", "X");
        assert_eq!(count, 1, "the outer paragraph's text is ab + ef");
        assert!(result.contains("<w:t>aX</w:t>") && result.contains("<w:t>f</w:t>"), "{result}");
    }

    /// Text that now begins or ends with a space needs
    /// `xml:space="preserve"`, or Word drops the space.
    #[test]
    fn test_replace_preserves_new_edge_whitespace() {
        let xml = r#"<w:p><w:r><w:t>a-b</w:t></w:r></w:p>"#;
        let (result, _) = replace_across_runs(xml, "-", " ");
        assert_eq!(result, r#"<w:p><w:r><w:t>a b</w:t></w:r></w:p>"#);
        let (result, _) = replace_across_runs(xml, "a", " ");
        assert_eq!(result, r#"<w:p><w:r><w:t xml:space="preserve"> -b</w:t></w:r></w:p>"#);
    }

    /// Headers, footers, footnotes, endnotes and comments are separate
    /// parts; only `document.xml` was edited, so a replacement reported
    /// success while the same text stayed everywhere else.
    #[test]
    fn test_replace_reaches_headers_footers_notes_and_comments() {
        use crate::core::opc::OpcWriter;
        use crate::core::relationships::rel_types;
        use std::io::Cursor;

        const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;
        let para = "<w:p><w:r><w:t>ACME</w:t></w:r></w:p>";
        let mut opc = OpcWriter::new(Cursor::new(Vec::new())).unwrap();
        let doc = PartName::new("/word/document.xml").unwrap();
        opc.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        let ct = "application/vnd.openxmlformats-officedocument.wordprocessingml";
        opc.add_part(
            &doc,
            &format!("{ct}.document.main+xml"),
            format!("<w:document {W}><w:body>{para}</w:body></w:document>").as_bytes(),
        )
        .unwrap();
        for (file, root, rel, kind) in [
            ("header1.xml", "w:hdr", rel_types::HEADER, "header"),
            ("footer1.xml", "w:ftr", rel_types::FOOTER, "footer"),
            ("footnotes.xml", "w:footnotes", rel_types::FOOTNOTES, "footnotes"),
            ("endnotes.xml", "w:endnotes", rel_types::ENDNOTES, "endnotes"),
            ("comments.xml", "w:comments", rel_types::COMMENTS, "comments"),
        ] {
            let part = PartName::new(&format!("/word/{file}")).unwrap();
            let body = match root {
                "w:hdr" | "w:ftr" => para.to_string(),
                "w:comments" => format!(r#"<w:comment w:id="0" w:author="A">{para}</w:comment>"#),
                "w:footnotes" => format!(r#"<w:footnote w:id="1">{para}</w:footnote>"#),
                _ => format!(r#"<w:endnote w:id="1">{para}</w:endnote>"#),
            };
            opc.add_part(
                &part,
                &format!("{ct}.{kind}+xml"),
                format!("<{root} {W}>{body}</{root}>").as_bytes(),
            )
            .unwrap();
            opc.add_part_rel(&doc, rel, file);
        }
        let bytes = opc.finish().unwrap().into_inner();

        let mut edit = EditableDocx::from_reader(Cursor::new(bytes)).unwrap();
        assert_eq!(edit.replace_text("ACME", "Globex"), 6);
        let mut out = Cursor::new(Vec::new());
        edit.write_to(&mut out).unwrap();
        let mut zip = zip::ZipArchive::new(Cursor::new(out.into_inner())).unwrap();
        for part in [
            "document.xml",
            "header1.xml",
            "footer1.xml",
            "footnotes.xml",
            "endnotes.xml",
            "comments.xml",
        ] {
            let mut xml = String::new();
            std::io::Read::read_to_string(
                &mut zip.by_name(&format!("word/{part}")).unwrap(),
                &mut xml,
            )
            .unwrap();
            assert!(xml.contains("Globex") && !xml.contains("ACME"), "{part}: {xml}");
        }
    }

    #[test]
    fn test_replace_with_an_empty_pattern_changes_nothing() {
        let xml = r#"<w:p><w:r><w:t>abc</w:t></w:r></w:p>"#;
        assert_eq!(replace_across_runs(xml, "", "x"), (xml.to_string(), 0));
    }

    #[test]
    fn test_no_match_returns_zero() {
        let xml = r#"<w:r><w:t>Hello</w:t></w:r>"#;
        let (result, count) = replace_across_runs(xml, "xyz", "abc");
        assert_eq!(count, 0);
        assert_eq!(result, xml);
    }
}
