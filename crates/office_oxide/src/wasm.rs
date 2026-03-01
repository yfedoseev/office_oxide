use std::io::Cursor;

use wasm_bindgen::prelude::*;

use crate::format::DocumentFormat;
use crate::Document;

/// A parsed Office document for use in JavaScript/WASM.
#[wasm_bindgen]
pub struct WasmDocument {
    inner: Document,
}

#[wasm_bindgen]
impl WasmDocument {
    /// Create a new document from raw bytes and a format string ("docx", "xlsx", "pptx").
    #[wasm_bindgen(constructor)]
    pub fn new(data: &[u8], format: &str) -> Result<WasmDocument, JsValue> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| JsValue::from_str(&format!("unsupported format: {format}")))?;
        let cursor = Cursor::new(data.to_vec());
        let inner = Document::from_reader(cursor, fmt)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(WasmDocument { inner })
    }

    /// Return the format name: "docx", "xlsx", or "pptx".
    #[wasm_bindgen(js_name = "formatName")]
    pub fn format_name(&self) -> String {
        format!("{:?}", self.inner.format()).to_lowercase()
    }

    /// Extract plain text from the document.
    #[wasm_bindgen(js_name = "plainText")]
    pub fn plain_text(&self) -> String {
        self.inner.plain_text()
    }

    /// Convert the document to markdown.
    #[wasm_bindgen(js_name = "toMarkdown")]
    pub fn to_markdown(&self) -> String {
        self.inner.to_markdown()
    }

    /// Convert the document to a JSON IR representation.
    #[wasm_bindgen(js_name = "toIr")]
    pub fn to_ir(&self) -> Result<JsValue, JsValue> {
        let ir = self.inner.to_ir();
        serde_wasm_bindgen::to_value(&ir_to_serializable(&ir))
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }
}

fn ir_to_serializable(ir: &crate::ir::DocumentIR) -> serde_json::Value {
    use crate::ir::*;
    use serde_json::json;

    json!({
        "metadata": {
            "format": format!("{:?}", ir.metadata.format),
            "title": ir.metadata.title,
        },
        "sections": ir.sections.iter().map(|s| {
            json!({
                "title": s.title,
                "elements": s.elements.iter().map(elem_json).collect::<Vec<_>>(),
            })
        }).collect::<Vec<_>>(),
    })
}

fn elem_json(elem: &crate::ir::Element) -> serde_json::Value {
    use crate::ir::*;
    use serde_json::json;

    match elem {
        Element::Heading(h) => json!({
            "type": "heading", "level": h.level,
            "content": inline_json(&h.content),
        }),
        Element::Paragraph(p) => json!({
            "type": "paragraph",
            "content": inline_json(&p.content),
        }),
        Element::Table(t) => json!({
            "type": "table",
            "rows": t.rows.iter().map(|r| json!({
                "is_header": r.is_header,
                "cells": r.cells.iter().map(|c| json!({
                    "col_span": c.col_span, "row_span": c.row_span,
                    "content": c.content.iter().map(elem_json).collect::<Vec<_>>(),
                })).collect::<Vec<_>>(),
            })).collect::<Vec<_>>(),
        }),
        Element::List(l) => json!({
            "type": "list", "ordered": l.ordered,
            "items": list_json(&l.items),
        }),
        Element::Image(img) => json!({ "type": "image", "alt_text": img.alt_text }),
        Element::ThematicBreak => json!({ "type": "thematic_break" }),
    }
}

fn inline_json(content: &[crate::ir::InlineContent]) -> Vec<serde_json::Value> {
    use crate::ir::*;
    use serde_json::json;
    content.iter().map(|item| match item {
        InlineContent::Text(span) => json!({
            "type": "text", "text": span.text,
            "bold": span.bold, "italic": span.italic,
            "strikethrough": span.strikethrough, "hyperlink": span.hyperlink,
        }),
        InlineContent::LineBreak => json!({ "type": "line_break" }),
    }).collect()
}

fn list_json(items: &[crate::ir::ListItem]) -> Vec<serde_json::Value> {
    use serde_json::json;
    items.iter().map(|item| {
        let mut obj = json!({ "content": inline_json(&item.content) });
        if let Some(ref nested) = item.nested {
            obj["nested"] = json!({ "ordered": nested.ordered, "items": list_json(&nested.items) });
        }
        obj
    }).collect()
}
