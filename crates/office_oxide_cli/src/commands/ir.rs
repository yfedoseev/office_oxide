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
        },
        "sections": ir.sections.iter().map(|s| {
            json!({
                "title": s.title,
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
            "content": inline_to_json(&h.content),
        }),
        Element::Paragraph(p) => json!({
            "type": "paragraph",
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
        }),
        Element::ThematicBreak => json!({ "type": "thematic_break" }),
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
            }),
            InlineContent::LineBreak => json!({ "type": "line_break" }),
        })
        .collect()
}

fn list_items_to_json(items: &[office_oxide::ir::ListItem]) -> Vec<serde_json::Value> {
    use serde_json::json;

    items
        .iter()
        .map(|item| {
            let mut obj = json!({
                "content": inline_to_json(&item.content),
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
