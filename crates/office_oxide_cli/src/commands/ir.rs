use office_oxide::Document;

pub fn run(file: &str) -> Result<(), Box<dyn std::error::Error>> {
    let doc = Document::open(file)?;
    let ir = doc.to_ir();
    let json = ir_to_json(&ir);
    println!("{}", serde_json::to_string_pretty(&json)?);
    Ok(())
}

fn ir_to_json(ir: &office_oxide::DocumentIR) -> serde_json::Value {
    use serde_json::json;

    json!({
        "metadata": {
            "format": format!("{:?}", ir.metadata.format),
            "title": ir.metadata.title,
            "author": ir.metadata.author,
            "subject": ir.metadata.subject,
            "keywords": ir.metadata.keywords,
            "created": ir.metadata.created,
            "modified": ir.metadata.modified,
            "description": ir.metadata.description,
            "has_macros": ir.metadata.has_macros,
            "text_truncated": ir.metadata.text_truncated,
        },
        "sections": ir.sections.iter().map(|s| {
            // speaker_notes is a sibling of `elements`, not one of them.
            // Projecting only `elements` dropped a slide's notes from this
            // surface entirely once they stopped being a paragraph.
            json!({
                "title": s.title,
                "hidden": s.hidden,
                "speaker_notes": s.speaker_notes,
                "conditional_formats": s.conditional_formats.iter().map(|cf| json!({
                    "range": cf.range,
                    "rule_type": cf.rule_type,
                    "operator": cf.operator,
                    "formulas": cf.formulas,
                })).collect::<Vec<_>>(),
                "elements": s.elements.iter().map(element_to_json).collect::<Vec<_>>(),
            })
        }).collect::<Vec<_>>(),
    })
}

fn element_to_json(elem: &office_oxide::ir::Element) -> serde_json::Value {
    use office_oxide::ir::*;
    use serde_json::json;

    match elem {
        Element::Heading(h) => json!({
            "type": "heading",
            "level": h.level,
            "alignment": h.alignment,
            "content": inline_to_json(&h.content),
        }),
        Element::Paragraph(p) => json!({
            "type": "paragraph",
            "alignment": p.alignment,
            "content": inline_to_json(&p.content),
        }),
        Element::Table(t) => json!({
            "type": "table",
            "rows": t.rows.iter().map(|r| json!({
                "is_header": r.is_header,
                "cells": r.cells.iter().map(|c| json!({
                    "col_span": c.col_span,
                    "row_span": c.row_span,
                    "content": c.content.iter().map(element_to_json).collect::<Vec<_>>(),
                    "data_type": c.data_type,
                    "raw_number": c.raw_number,
                    "number_format": c.number_format,
                    "number_format_id": c.number_format_id,
                    "formula": c.formula,
                })).collect::<Vec<_>>(),
            })).collect::<Vec<_>>(),
        }),
        Element::List(l) => json!({
            "type": "list",
            "ordered": l.ordered,
            "items": list_items_to_json(&l.items),
        }),
        Element::Image(img) => json!({
            "type": "image",
            "alt_text": img.alt_text,
            "hyperlink": img.hyperlink,
        }),
        Element::ThematicBreak => json!({ "type": "thematic_break" }),
        Element::TextBox(tb) => json!({
            "type": "text_box",
            "elements": tb.content.iter().map(element_to_json).collect::<Vec<_>>(),
        }),
        Element::PageBreak => json!({ "type": "page_break" }),
        Element::ColumnBreak => json!({ "type": "column_break" }),
        Element::Footnote(n) => json!({
            "type": "footnote",
            "id": n.id,
            "marker": n.marker,
            "elements": n.content.iter().map(element_to_json).collect::<Vec<_>>(),
        }),
        Element::Endnote(n) => json!({
            "type": "endnote",
            "id": n.id,
            "marker": n.marker,
            "elements": n.content.iter().map(element_to_json).collect::<Vec<_>>(),
        }),
        Element::CodeBlock(cb) => json!({
            "type": "code_block",
            "language": cb.language,
            "content": cb.content,
        }),
        Element::Shape(s) => json!({
            "type": "shape",
            "kind": format!("{:?}", s.kind),
            "x_emu": s.x_emu,
            "y_emu": s.y_emu,
            "width_emu": s.width_emu,
            "height_emu": s.height_emu,
        }),
        // `Element` is `#[non_exhaustive]`, so rustc requires a wildcard
        // arm here regardless — a genuinely new variant can't be turned
        // into a compile error from outside the defining crate. This is
        // the fallback of last resort: it at least carries the Debug
        // dump, so a new variant is *visibly incomplete* on this surface
        // rather than indistinguishable from a variant that was properly
        // handled (issue #221; #221's own maximal-Section round-trip
        // test, and this crate's Shape-specific test, are what actually
        // catch a future miss like this one).
        other => json!({ "type": "unimplemented", "debug": format!("{other:?}") }),
    }
}

fn inline_to_json(content: &[office_oxide::ir::InlineContent]) -> Vec<serde_json::Value> {
    use office_oxide::ir::*;
    use serde_json::json;

    content
        .iter()
        .map(|item| match item {
            InlineContent::Text(span) => json!({
                "type": "text",
                "text": span.text,
                "bold": span.bold,
                "italic": span.italic,
                "strikethrough": span.strikethrough,
                "hyperlink": span.hyperlink,
                "underline": span.underline,
                "font_size_half_pt": span.font_size_half_pt,
                "color": span.color.map(|[r, g, b]| format!("{r:02X}{g:02X}{b:02X}")),
                "font_name": span.font_name,
                "highlight": span.highlight.map(|[r, g, b]| format!("{r:02X}{g:02X}{b:02X}")),
                "vertical_align": span.vertical_align,
                "all_caps": span.all_caps,
                "small_caps": span.small_caps,
                "char_spacing_half_pt": span.char_spacing_half_pt,
            }),
            InlineContent::LineBreak => json!({ "type": "line_break" }),
            InlineContent::FootnoteRef(r) => json!({
                "type": "footnote_ref",
                "id": r.note_id,
            }),
            InlineContent::EndnoteRef(r) => json!({
                "type": "endnote_ref",
                "id": r.note_id,
            }),
            // `InlineContent` is also `#[non_exhaustive]`; same reasoning
            // as element_to_json above (issue #221).
            other => json!({ "type": "unimplemented", "debug": format!("{other:?}") }),
        })
        .collect()
}

fn list_items_to_json(items: &[office_oxide::ir::ListItem]) -> Vec<serde_json::Value> {
    use serde_json::json;

    items
        .iter()
        .map(|item| {
            let mut obj = json!({
                "content": item.content.iter().map(element_to_json).collect::<Vec<_>>(),
            });
            if let Some(ref nested) = item.nested {
                obj["nested"] = json!({
                    "ordered": nested.ordered,
                    "items": list_items_to_json(&nested.items),
                });
            }
            obj
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use office_oxide::DocumentIR;
    use office_oxide::format::DocumentFormat;
    use office_oxide::ir::*;

    use super::ir_to_json;

    /// issue #221 — speaker_notes is a sibling of `Section::elements`, not
    /// one of its items; the CLI's JSON projection missed it once already
    /// (only caught by a multi-thousand-file corpus sweep). Locked in here
    /// as a fast unit test.
    #[test]
    fn test_speaker_notes_reach_the_json_projection() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Pptx,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![],
                speaker_notes: Some("SPEAKER_NOTES_MARKER".to_string()),
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(
            rendered.contains("SPEAKER_NOTES_MARKER"),
            "the ir command's JSON projection must include speaker_notes: {rendered}"
        );
    }

    /// A `Shape` element (the one variant the exhaustive match in
    /// `element_to_json` was missing) must render as its own type, not
    /// silently fall through to a generic "unknown".
    #[test]
    fn test_shape_element_does_not_render_as_unknown() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Docx,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![Element::Shape(Shape {
                    kind: ShapeGeom::Rect,
                    ..Default::default()
                })],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(
            rendered.contains(r#""type":"shape""#),
            "a Shape element must render as its own type, not unknown: {rendered}"
        );
        assert!(!rendered.contains("unknown"), "no element should render as unknown: {rendered}");
    }

    /// issue #299 — `Image::hyperlink` (a shape's own click action) was
    /// added to the IR but the CLI's JSON projection only ever surfaced
    /// `alt_text`, the same "field added, one consumer missed" shape
    /// #221 already found once for `speaker_notes`.
    #[test]
    fn test_image_hyperlink_reaches_the_json_projection() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Pptx,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![Element::Image(office_oxide::ir::Image {
                    hyperlink: Some("#slide2.xml".to_string()),
                    ..Default::default()
                })],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(
            rendered.contains("#slide2.xml"),
            "the ir command's JSON projection must include Image::hyperlink: {rendered}"
        );
    }

    /// issue #332 — `.doc` comments and real endnotes both reach the IR as
    /// `Element::Endnote` (there's no dedicated `Element::Comment`), and
    /// `Note::marker` ("comment" vs "endnote") is the only thing that tells
    /// them apart. The JSON projection dropped it, making the two
    /// indistinguishable in `ir` output.
    #[test]
    fn test_note_marker_reaches_the_json_projection() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Doc,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![
                    Element::Endnote(Note {
                        id: 0,
                        marker: Some("comment".to_string()),
                        content: vec![],
                    }),
                    Element::Footnote(Note {
                        id: 1,
                        marker: Some("footnote".to_string()),
                        content: vec![],
                    }),
                ],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(
            rendered.contains(r#""marker":"comment""#),
            "a comment's marker must distinguish it from a real endnote: {rendered}"
        );
        assert!(
            rendered.contains(r#""marker":"footnote""#),
            "a footnote's marker must also reach the projection: {rendered}"
        );
    }

    /// issue #333 — `inline_to_json` used to project only `text`/`bold`/
    /// `italic`/`strikethrough`/`hyperlink` from a `TextSpan`, silently
    /// dropping every other formatting field the IR actually carries.
    #[test]
    fn test_text_span_formatting_fields_reach_the_json_projection() {
        use office_oxide::ir::{TextSpan, UnderlineStyle, VerticalAlign};

        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Ppt,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                elements: vec![Element::Paragraph(Paragraph {
                    content: vec![InlineContent::Text(TextSpan {
                        underline: Some(UnderlineStyle::Single),
                        font_size_half_pt: Some(36),
                        color: Some([0x12, 0x34, 0x56]),
                        font_name: Some("Calibri".to_string()),
                        vertical_align: Some(VerticalAlign::Superscript),
                        all_caps: true,
                        ..TextSpan::plain("styled")
                    })],
                    alignment: Some(ParagraphAlignment::Center),
                    ..Default::default()
                })],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(rendered.contains(r#""underline":"single""#), "{rendered}");
        assert!(rendered.contains(r#""font_size_half_pt":36"#), "{rendered}");
        assert!(rendered.contains(r#""color":"123456""#), "{rendered}");
        assert!(rendered.contains(r#""font_name":"Calibri""#), "{rendered}");
        assert!(rendered.contains(r#""vertical_align":"superscript""#), "{rendered}");
        assert!(rendered.contains(r#""all_caps":true"#), "{rendered}");
        assert!(rendered.contains(r#""alignment":"center""#), "{rendered}");
    }

    /// issue #252 — a worksheet's conditional formatting rules must reach
    /// the `ir` command's JSON output.
    #[test]
    fn test_conditional_formats_reach_the_json_projection() {
        let ir = DocumentIR {
            metadata: Metadata {
                format: DocumentFormat::Xlsx,
                title: None,
                ..Default::default()
            },
            sections: vec![Section {
                conditional_formats: vec![office_oxide::ir::ConditionalFormat {
                    range: "A1:A10".to_string(),
                    rule_type: "cellIs".to_string(),
                    operator: Some("greaterThan".to_string()),
                    formulas: vec!["100".to_string()],
                }],
                ..Default::default()
            }],
            defined_names: Vec::new(),
        };
        let json = ir_to_json(&ir);
        let rendered = serde_json::to_string(&json).unwrap();
        assert!(
            rendered.contains(r#""range":"A1:A10""#) && rendered.contains(r#""rule_type":"cellIs""#),
            "the ir command's JSON projection must include Section::conditional_formats: {rendered}"
        );
    }
}
