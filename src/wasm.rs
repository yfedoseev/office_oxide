// SPDX-License-Identifier: MIT OR Apache-2.0
use std::io::Cursor;

use wasm_bindgen::prelude::*;

use crate::Document;
use crate::format::DocumentFormat;

/// Options for one document: limits, which text is kept, and whether an
/// incomplete result is an error. `new WasmParseOptions(true)` starts from
/// tighter limits and makes an incomplete result throw; `set(name, value)`
/// sets any option by name (see `WasmParseOptions.names()`).
#[wasm_bindgen]
pub struct WasmParseOptions {
    inner: crate::options::ParseOptions,
}

#[wasm_bindgen]
impl WasmParseOptions {
    /// Options holding the defaults, or the strict ones when `strict` is true.
    #[wasm_bindgen(constructor)]
    pub fn new(strict: Option<bool>) -> WasmParseOptions {
        WasmParseOptions {
            inner: if strict.unwrap_or(false) {
                crate::options::ParseOptions::strict()
            } else {
                crate::options::ParseOptions::default()
            },
        }
    }

    /// Set an option by name. `value` is a number (`Infinity` for no bound), a
    /// boolean, or for `on_incomplete` the string `"report"` or `"error"`.
    pub fn set(&mut self, name: &str, value: JsValue) -> Result<(), JsValue> {
        let bad = |why: &str| JsValue::from_str(&format!("{name}: {why}"));
        let number = if let Some(flag) = value.as_bool() {
            u64::from(flag)
        } else if let Some(text) = value.as_string() {
            // Words mean one thing for one option each.
            match (name, text.as_str()) {
                ("on_incomplete", "report") => 0,
                ("on_incomplete", "error") => 1,
                ("on_incomplete", _) => return Err(bad("expected \"report\" or \"error\"")),
                (_, "max") => u64::MAX,
                _ => return Err(bad("a string is only accepted as \"max\", or for on_incomplete")),
            }
        } else if let Some(n) = value.as_f64() {
            if n == f64::INFINITY {
                u64::MAX
            } else if n.is_nan() || n < 0.0 || n.fract() != 0.0 || n >= 18_446_744_073_709_551_616.0
            {
                return Err(bad("expected a non-negative whole number, or Infinity for no bound"));
            } else {
                n as u64
            }
        } else {
            return Err(bad("expected a number, boolean or string"));
        };
        self.inner
            .set_option(name, number)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// The names of every option.
    pub fn names() -> Vec<String> {
        let mut names = vec![
            "include_hidden_text".to_string(),
            "include_deleted_text".to_string(),
            "on_incomplete".to_string(),
        ];
        names.extend(
            crate::limits::Limits::NAMES
                .iter()
                .map(|n| (*n).to_string()),
        );
        names
    }
}

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
        Self::open(data, format, None)
    }

    /// Create a document under `options` (a `WasmParseOptions`, which is only
    /// borrowed: the same object can open any number of documents).
    #[wasm_bindgen(js_name = "openWith")]
    pub fn open_with(
        data: &[u8],
        format: &str,
        options: &WasmParseOptions,
    ) -> Result<WasmDocument, JsValue> {
        Self::open(data, format, Some(options))
    }

    fn open(
        data: &[u8],
        format: &str,
        options: Option<&WasmParseOptions>,
    ) -> Result<WasmDocument, JsValue> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| JsValue::from_str(&format!("unsupported format: {format}")))?;
        let cursor = Cursor::new(data.to_vec());
        let inner = match options {
            Some(o) => Document::from_reader_with(cursor, fmt, &o.inner),
            None => Document::from_reader(cursor, fmt),
        }
        .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(WasmDocument { inner })
    }

    /// What is wrong with the document: an array of `{kind, part, detail}`
    /// (`kind` is `unreadable_part`, `truncated` or `recovered`); empty for a
    /// sound document. Converts the document once.
    pub fn damage(&self) -> Result<JsValue, JsValue> {
        use serde::Serialize;
        // `part` is null, not undefined, when unknown.
        let serializer = serde_wasm_bindgen::Serializer::new().serialize_missing_as_null(true);
        self.inner
            .damage()
            .serialize(&serializer)
            .map_err(|e| JsValue::from_str(&e.to_string()))
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

    /// Convert to markdown, embedding each image inline as
    /// `[image-base64:<data>]` at its position in the document flow.
    ///
    /// Images are otherwise dropped from markdown entirely, which loses
    /// both their content and their position.
    #[wasm_bindgen(js_name = toMarkdownWithImages)]
    pub fn to_markdown_with_images(&self) -> String {
        use crate::ir_render::{ImageEmbed, MarkdownOptions};
        self.inner.to_markdown_with(MarkdownOptions {
            image_embed: ImageEmbed::Base64,
        })
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

/// An Office document opened for editing from JavaScript/WASM.
///
/// WASM had no file-editing API at all, so an agent could read a document
/// but never change one. The document is held in memory with every OPC part
/// preserved, so parts this crate does not model (images, charts, custom
/// XML) survive the round-trip untouched.
///
/// Supports `replaceText` (DOCX/PPTX), `setCell` (XLSX) and `toBytes`. The
/// document builders, `createFromMarkdown` and `saveAs` are native-binding
/// only.
#[wasm_bindgen]
pub struct WasmEditableDocument {
    inner: crate::edit::EditableDocument,
}

#[wasm_bindgen]
impl WasmEditableDocument {
    /// Open a document for editing from raw bytes and a format string
    /// ("docx", "xlsx", "pptx").
    #[wasm_bindgen(constructor)]
    pub fn new(data: &[u8], format: &str) -> Result<WasmEditableDocument, JsValue> {
        let fmt = DocumentFormat::from_extension(format)
            .ok_or_else(|| JsValue::from_str(&format!("unsupported format: {format}")))?;
        let cursor = Cursor::new(data.to_vec());
        let inner = crate::edit::EditableDocument::from_reader(cursor, fmt)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(WasmEditableDocument { inner })
    }

    /// Replace every occurrence of `find` with `replace` in the document's
    /// text, returning how many replacements were made.
    #[wasm_bindgen(js_name = replaceText)]
    pub fn replace_text(&mut self, find: &str, replace: &str) -> Result<usize, JsValue> {
        self.inner
            .replace_text(find, replace)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Set a cell in an XLSX document. `value` is `null`/`undefined`
    /// (empty), a string, a number or a boolean; anything else throws.
    /// Throws for a non-XLSX document or a bad cell reference.
    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(
        &mut self,
        sheet_index: usize,
        cell_ref: &str,
        value: JsValue,
    ) -> Result<(), JsValue> {
        use crate::xlsx::edit::CellValue;
        let cv = if value.is_null() || value.is_undefined() {
            CellValue::Empty
        } else if let Some(b) = value.as_bool() {
            CellValue::Boolean(b)
        } else if let Some(s) = value.as_string() {
            CellValue::String(s)
        } else if let Some(n) = value.as_f64() {
            CellValue::Number(n)
        } else {
            return Err(JsValue::from_str(
                "setCell: value must be null, a string, a number or a boolean",
            ));
        };
        self.inner
            .set_cell(sheet_index, cell_ref, cv)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Serialise the edited document back to bytes.
    #[wasm_bindgen(js_name = toBytes)]
    pub fn to_bytes(&self) -> Result<Vec<u8>, JsValue> {
        let mut buf = Cursor::new(Vec::new());
        self.inner
            .write_to(&mut buf)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(buf.into_inner())
    }
}
