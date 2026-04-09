use std::io::Cursor;

use wasm_bindgen::prelude::*;

use crate::Document;
use crate::format::DocumentFormat;

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
        let inner =
            Document::from_reader(cursor, fmt).map_err(|e| JsValue::from_str(&e.to_string()))?;
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

    /// Convert the document to an HTML fragment.
    #[wasm_bindgen(js_name = "toHtml")]
    pub fn to_html(&self) -> String {
        self.inner.to_html()
    }

    /// Convert the document to a JSON IR representation.
    #[wasm_bindgen(js_name = "toIr")]
    pub fn to_ir(&self) -> Result<JsValue, JsValue> {
        let ir = self.inner.to_ir();
        serde_wasm_bindgen::to_value(&ir).map_err(|e| JsValue::from_str(&e.to_string()))
    }
}
