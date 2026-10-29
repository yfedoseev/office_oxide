use serde_json::{Value, json};

pub fn handle_initialize(id: &Value) -> Value {
    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": {}
            },
            "serverInfo": {
                "name": "office-oxide-mcp",
                "version": env!("CARGO_PKG_VERSION")
            }
        }
    })
}

pub fn handle_tools_list(id: &Value) -> Value {
    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "tools": [
                {
                    "name": "extract",
                    "description":
                        "Extract content from an Office document (DOCX, XLSX, PPTX, DOC, XLS, PPT)",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the document file"
                            },
                            "format": {
                                "type": "string",
                                "enum": [
                                    "text", "markdown", "markdown-with-images", "html", "ir"
                                ],
                                "description":
                                    "Output format (default: text). \
                                     `markdown-with-images` embeds each image inline as \
                                     [image-base64:...] at its position in the flow."
                            },
                            "strict": {
                                "type": "boolean",
                                "description":
                                    "Read untrusted input: tighter limits, and an error \
                                     instead of a partial result when part of the document \
                                     cannot be read or a limit cut it (default: false)"
                            }
                        },
                        "required": ["file_path"]
                    }
                },
                {
                    "name": "replace_text",
                    "description":
                        "Replace text in an Office document (DOCX or PPTX), preserving \
                         every other part of the file. Writes to output_path, or in place \
                         when output_path is omitted.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the document file"
                            },
                            "find": {
                                "type": "string",
                                "description": "Text to search for (must not be empty)"
                            },
                            "replace": {
                                "type": "string",
                                "description": "Replacement text"
                            },
                            "output_path": {
                                "type": "string",
                                "description":
                                    "Where to write the result (default: overwrite file_path)"
                            }
                        },
                        "required": ["file_path", "find", "replace"]
                    }
                },
                {
                    "name": "info",
                    "description":
                        "Get metadata about an Office document (DOCX, XLSX, PPTX, DOC, XLS, \
                         PPT): format, file size, title, document properties (author, \
                         subject, keywords, dates), warnings such as incomplete text \
                         extraction, and the section list",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Path to the document file"
                            },
                            "strict": {
                                "type": "boolean",
                                "description":
                                    "Read untrusted input: tighter limits, and an error \
                                     instead of a partial result when part of the document \
                                     cannot be read or a limit cut it (default: false)"
                            }
                        },
                        "required": ["file_path"]
                    }
                }
            ]
        }
    })
}

/// JSON-RPC 2.0 §5.1 "Invalid params". MCP also uses it for an unknown tool
/// name: `tools/call` itself exists, so -32601 "Method not found" is wrong.
const INVALID_PARAMS: i64 = -32602;

pub fn handle_tools_call(id: &Value, params: &Value) -> Value {
    let Some(tool_name) = params.get("name").and_then(Value::as_str) else {
        return error_response(id, INVALID_PARAMS, "tools/call requires a string \"name\"");
    };
    let arguments = &params["arguments"];

    match tool_name {
        "extract" => call_extract(id, arguments),
        "replace_text" => call_replace_text(id, arguments),
        "info" => call_info(id, arguments),
        _ => error_response(id, INVALID_PARAMS, &format!("unknown tool: {tool_name}")),
    }
}

/// Largest `extract` result the server will return, in UTF-8 bytes.
///
/// A tool result is one JSON string held in memory (and then escaped into a
/// second copy for the response line), and no MCP client can put tens of
/// megabytes in front of a model anyway. The CLI documents what an unbounded
/// buffer did on large workbooks — multi-gigabyte RSS and OOM kills — so an
/// oversized result is refused with a message saying so, rather than
/// silently truncated. The limit counts bytes, not characters: a character
/// is one to four bytes, so text in a script that needs more than one (CJK,
/// say) reaches the limit at fewer characters than ASCII does.
const MAX_EXTRACT_BYTES: usize = 32 * 1024 * 1024;

/// Largest total of image bytes one document conversion may carry, in bytes.
///
/// Images reach a result only as base64 (`markdown-with-images`, `ir`), which
/// is four thirds of the bytes, so this keeps them near [`MAX_EXTRACT_BYTES`].
const MAX_EXTRACT_IMAGE_BYTES: usize = 24 * 1024 * 1024;

/// The options a tool call reads its document under: the library's per-document
/// budgets brought down to what one tool result can hold.
///
/// The renderers build their whole output before it can be measured, and the
/// library's default text budget (256 Mi characters) lets a spreadsheet run
/// to gigabytes before the size check could refuse it. Under these options the
/// text budget is [`MAX_EXTRACT_BYTES`] characters, so a spreadsheet's output
/// stops there with the library's usual truncation notice; the image budget
/// is lowered the same way. A budget that is already lower is kept.
///
/// The options belong to the call: nothing process-wide is changed, so calls
/// no longer have to be serialised against each other or against a library
/// user's own settings. With `"strict": true` the document is read under the
/// library's strict limits and refused if it cannot be read completely.
///
/// Only the tools that convert a document to text or to the IR take them:
/// `extract`, and `info`, which converts the whole document to read its
/// metadata (it would otherwise build an IR of up to the default budget,
/// gigabytes for a hostile spreadsheet; a document with more than 32 Mi
/// characters of text is reported there as incompletely extracted).
/// `replace_text` edits the package's XML in place and reads none of these
/// budgets.
fn tool_options(args: &Value) -> office_oxide::options::ParseOptions {
    let mut options = if args["strict"].as_bool().unwrap_or(false) {
        office_oxide::options::ParseOptions::strict()
    } else {
        office_oxide::options::ParseOptions::default()
    };
    options.limits.text_chars = options.limits.text_chars.min(MAX_EXTRACT_BYTES);
    options.limits.image_bytes = options.limits.image_bytes.min(MAX_EXTRACT_IMAGE_BYTES);
    options
}

fn call_extract(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing file_path");
    };
    let format = args["format"].as_str().unwrap_or("text");

    // The options apply to the opening too: the legacy spreadsheet reader
    // copies shared strings into cells as it opens.
    let doc = match office_oxide::Document::open_with(file_path, &tool_options(args)) {
        Ok(d) => d,
        Err(e) => return tool_error(id, &e.to_string()),
    };

    match render(&doc, format, MAX_EXTRACT_BYTES) {
        Ok(content) => json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": {
                "content": [{ "type": "text", "text": content }]
            }
        }),
        Err(message) => tool_error(id, &message),
    }
}

/// Render `doc` in `format`, refusing a result larger than `limit` bytes.
///
/// Only a spreadsheet is produced as it is written, so only there does the
/// bounded buffer stop the work at the limit. Every other format has no
/// streaming writer in the library: its text, Markdown or HTML is built whole
/// (the IR is built whole, then serialised into the buffer) before the buffer
/// can refuse it, and memory is bounded there by the library's text and image
/// budgets (see [`tool_options`]), not by `limit`.
fn render(doc: &office_oxide::Document, format: &str, limit: usize) -> Result<String, String> {
    // Every format is written through a bounded buffer, so an oversized
    // document stops at the limit instead of building the whole string (and
    // a spreadsheet is produced a sheet at a time).
    let mut out = BoundedBuf {
        buf: Vec::new(),
        limit,
        exceeded: false,
    };
    let written = match format {
        "text" => doc.write_plain_text(&mut out),
        "markdown" => doc.write_markdown(&mut out),
        // Images are dropped from plain markdown entirely; this keeps both
        // their content and their position in one self-contained string.
        "markdown-with-images" => {
            use office_oxide::ir_render::{ImageEmbed, MarkdownOptions};
            doc.write_markdown_with(
                MarkdownOptions {
                    image_embed: ImageEmbed::Base64,
                },
                &mut out,
            )
        },
        "html" => doc.write_html(&mut out),
        "ir" => serde_json::to_writer_pretty(&mut out, &doc.to_ir()).map_err(std::io::Error::from),
        other => return Err(format!("unknown format: {other}")),
    };
    match written {
        Ok(()) => String::from_utf8(out.buf).map_err(|e| e.to_string()),
        Err(_) if out.exceeded => Err(too_large(format, limit)),
        Err(e) => Err(e.to_string()),
    }
}

fn too_large(format: &str, limit: usize) -> String {
    format!(
        "the {format} output of this document exceeds the {} MiB (UTF-8 bytes) limit for one tool result; \
         try a more compact format (\"text\" or \"markdown\"), or run the office-oxide CLI, \
         which streams its output",
        limit / (1024 * 1024)
    )
}

/// A `Vec<u8>` writer that fails once `limit` bytes would be exceeded.
struct BoundedBuf {
    buf: Vec<u8>,
    limit: usize,
    exceeded: bool,
}

impl std::io::Write for BoundedBuf {
    fn write(&mut self, data: &[u8]) -> std::io::Result<usize> {
        if self.buf.len().saturating_add(data.len()) > self.limit {
            self.exceeded = true;
            return Err(std::io::Error::other("output limit exceeded"));
        }
        self.buf.extend_from_slice(data);
        Ok(data.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn call_replace_text(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing file_path");
    };
    let Some(find) = args["find"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing find");
    };
    let Some(replace) = args["replace"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing replace");
    };
    let output_path = args["output_path"].as_str().unwrap_or(file_path);

    let mut doc = match office_oxide::edit::EditableDocument::open(file_path) {
        Ok(d) => d,
        Err(e) => return tool_error(id, &e.to_string()),
    };
    // Report an unsupported format as an error rather than "0 occurrences":
    // an agent cannot tell a no-match from an unimplemented operation, and
    // the file was rewritten either way.
    let count = match doc.replace_text(find, replace) {
        Ok(n) => n,
        Err(e) => return tool_error(id, &e.to_string()),
    };
    if let Err(e) = doc.save(output_path) {
        return tool_error(id, &e.to_string());
    }

    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{
                "type": "text",
                "text": format!("replaced {count} occurrence(s); wrote {output_path}")
            }]
        }
    })
}

fn call_info(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing file_path");
    };

    let doc = match office_oxide::Document::open_with(file_path, &tool_options(args)) {
        Ok(d) => d,
        Err(e) => return tool_error(id, &e.to_string()),
    };

    let ir = doc.to_ir();
    let size = std::fs::metadata(file_path).ok().map(|m| m.len());
    let info = info_json(&ir, size);

    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{ "type": "text", "text": info.to_string() }]
        }
    })
}

/// Shown when the parser detected that it could not recover all of the
/// document's text. Same wording as the CLI's `info`.
const TRUNCATION_WARNING: &str =
    "text extraction is incomplete — part of the document could not be recovered safely";

/// The `info` tool's result: format, file size, every document property
/// that is set (absent ones are omitted rather than reported as null), the
/// custom properties, the content flags, the section list, and `warnings`.
///
/// `warnings` carries the truncation signal the CLI already surfaced:
/// without it an agent had no way to learn that an extraction was
/// incomplete.
fn info_json(ir: &office_oxide::DocumentIR, file_size: Option<u64>) -> Value {
    let meta = &ir.metadata;
    let mut warnings: Vec<&str> = Vec::new();
    if meta.text_truncated {
        warnings.push(TRUNCATION_WARNING);
    }
    // What the reader worked around, one line each (a skipped part, a
    // flattened structure).
    warnings.extend(meta.warnings.iter().map(String::as_str));
    let mut info = json!({
        "format": format!("{:?}", meta.format),
        "file_size": file_size,
        "title": meta.title,
        "author": meta.author,
        "subject": meta.subject,
        "keywords": meta.keywords,
        "description": meta.description,
        "created": meta.created,
        "modified": meta.modified,
        "last_modified_by": meta.last_modified_by,
        "revision": meta.revision,
        "category": meta.category,
        "content_status": meta.content_status,
        "language": meta.language,
        "company": meta.company,
        "manager": meta.manager,
        "custom_properties": meta.custom_properties,
        "has_macros": meta.has_macros,
        "has_digital_signature": meta.has_digital_signature,
        "has_thumbnail": meta.thumbnail.is_some(),
        "text_truncated": meta.text_truncated,
        "warnings": warnings,
        "sections": ir.sections.len(),
        "section_names": ir.sections.iter().map(|s| s.title.clone()).collect::<Vec<_>>(),
    });
    if let Some(map) = info.as_object_mut() {
        map.retain(|_, v| !v.is_null());
    }
    info
}

fn error_response(id: &Value, code: i64, message: &str) -> Value {
    json!({
        "jsonrpc": "2.0",
        "id": id,
        "error": { "code": code, "message": message }
    })
}

fn tool_error(id: &Value, message: &str) -> Value {
    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{ "type": "text", "text": message }],
            "isError": true
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Serialises the tests that lower a process-wide limit against the ones
    /// that open and render documents under the defaults: the test that bounds
    /// the budgets holds the write lock, the others the read lock.
    static LIMITS: std::sync::RwLock<()> = std::sync::RwLock::new(());

    fn default_limits() -> std::sync::RwLockReadGuard<'static, ()> {
        LIMITS.read().unwrap_or_else(|e| e.into_inner())
    }

    /// The text and image budgets set for the length of a test and put back
    /// when it ends, however it ends. Holds the write lock: no test that
    /// opens documents under the defaults (they hold the read lock) runs
    /// meanwhile, and no tool call lowers the budgets meanwhile.
    struct TestLimits {
        before: (usize, usize),
        _lock: std::sync::RwLockWriteGuard<'static, ()>,
    }

    impl TestLimits {
        fn bound_to(text_chars: usize, image_bytes: usize) -> Self {
            let lock = LIMITS.write().unwrap_or_else(|e| e.into_inner());
            let before =
                (office_oxide::limits::max_text_chars(), office_oxide::limits::max_image_bytes());
            office_oxide::limits::set_max_text_chars(text_chars);
            office_oxide::limits::set_max_image_bytes(image_bytes);
            Self {
                before,
                _lock: lock,
            }
        }
    }

    impl Drop for TestLimits {
        fn drop(&mut self) {
            office_oxide::limits::set_max_text_chars(self.before.0);
            office_oxide::limits::set_max_image_bytes(self.before.1);
        }
    }

    /// A unique scratch directory for one test.
    fn scratch_dir(tag: &str) -> std::path::PathBuf {
        let dir =
            std::env::temp_dir().join(format!("office_oxide_mcp_{tag}_{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    fn write_docx(path: &std::path::Path, text: &str) {
        let mut w = office_oxide::docx::write::DocxWriter::new();
        w.add_paragraph(text);
        w.save(path).unwrap();
    }

    /// The result of `extract` was one unbounded in-memory string — the
    /// pattern the CLI documents as costing gigabytes on large workbooks —
    /// with no guard. Over the limit it is now an explicit tool error for
    /// every format, including the streamed IR path.
    #[test]
    fn test_extract_refuses_output_over_the_size_limit() {
        let _defaults = default_limits();
        let dir = scratch_dir("size_limit");
        let path = dir.join("doc.docx");
        write_docx(&path, &"lorem ipsum ".repeat(200));
        let doc = office_oxide::Document::open(&path).unwrap();
        for format in ["text", "markdown", "markdown-with-images", "html", "ir"] {
            let full = render(&doc, format, usize::MAX).expect("unbounded render");
            assert!(full.len() > 64, "{format} too small to test");
            let err = render(&doc, format, 64).expect_err("over the limit must fail");
            assert!(err.contains("exceeds"), "{format}: {err}");
            // At exactly its own size it still fits.
            assert_eq!(render(&doc, format, full.len()).unwrap(), full, "{format}");
        }
        std::fs::remove_dir_all(&dir).ok();
    }

    /// A spreadsheet whose text outgrows the text budget is cut at it with the
    /// library's notice, in the rendering and in `info`, instead of being
    /// built whole and then refused.
    #[test]
    fn test_budgets_bound_what_a_spreadsheet_renders_before_the_size_check() {
        use office_oxide::xlsx::write::{CellData, XlsxWriter};
        let dir = scratch_dir("budget");
        let path = dir.join("book.xlsx");
        let mut w = XlsxWriter::new();
        {
            let mut sheet = w.add_sheet("S");
            for _ in 0..400 {
                sheet.add_row(vec![CellData::String("x".repeat(50))]);
            }
        }
        w.save(&path).unwrap();
        let scope = TestLimits::bound_to(2_000, 1_000);
        let doc = office_oxide::Document::open_with(&path, &tool_options(&json!({}))).unwrap();
        let cut = render(&doc, "markdown", usize::MAX);
        let info = call_info(&json!(1), &json!({ "file_path": path.to_str().unwrap() }));
        drop(scope);
        let _defaults = default_limits();
        // The document keeps the options it was opened with, so the whole text
        // comes from one opened without them.
        let whole = office_oxide::Document::open(&path).unwrap();
        let full = render(&whole, "text", usize::MAX).unwrap();
        assert!(full.len() >= 20_000);

        let cut = cut.unwrap();
        assert!(cut.len() < full.len() / 2, "{} bytes of {}", cut.len(), full.len());
        let info = info["result"]["content"][0]["text"]
            .as_str()
            .unwrap()
            .to_string();
        assert!(info.contains("\"text_truncated\":true"), "{info}");
        std::fs::remove_dir_all(&dir).ok();
    }

    /// A test that fails inside the scope still puts the budgets back, and
    /// the lock it held is usable afterwards, so one failing test cannot
    /// leave its limits behind for the others.
    #[test]
    fn test_test_limits_restore_the_limits_when_the_test_inside_it_panics() {
        let _defaults = default_limits();
        let before =
            (office_oxide::limits::max_text_chars(), office_oxide::limits::max_image_bytes());
        drop(_defaults);
        let panicked = std::panic::catch_unwind(|| {
            let _scope = TestLimits::bound_to(123, 45);
            assert_eq!(office_oxide::limits::max_text_chars(), 123);
            assert_eq!(office_oxide::limits::max_image_bytes(), 45);
            panic!("an assertion inside the scope fails");
        });
        assert!(panicked.is_err());
        let _defaults = default_limits();
        assert_eq!(
            (office_oxide::limits::max_text_chars(), office_oxide::limits::max_image_bytes()),
            before
        );
    }

    /// A tool call's options cap the budgets at what one result can hold, keep
    /// a lower process-wide budget, and change nothing process-wide.
    #[test]
    fn test_tool_options_cap_the_budgets_without_touching_the_process() {
        use office_oxide::limits::{max_image_bytes, max_text_chars};
        {
            let _limits = TestLimits::bound_to(usize::MAX, usize::MAX);
            let options = tool_options(&json!({}));
            assert_eq!(options.limits.text_chars, MAX_EXTRACT_BYTES);
            assert_eq!(options.limits.image_bytes, MAX_EXTRACT_IMAGE_BYTES);
            assert_eq!((max_text_chars(), max_image_bytes()), (usize::MAX, usize::MAX));
        }
        let _limits = TestLimits::bound_to(1_000, 500);
        let options = tool_options(&json!({}));
        assert_eq!(
            (options.limits.text_chars, options.limits.image_bytes),
            (1_000, 500),
            "a lower budget is kept"
        );
    }

    /// `"strict": true` reads under the strict limits and refuses a document
    /// that cannot be read completely; without it the readable part is returned.
    #[test]
    fn test_strict_refuses_an_incomplete_document() {
        use office_oxide::core::opc::{OpcWriter, PartName};
        use office_oxide::core::relationships::rel_types;
        let _defaults = default_limits();
        let dir = scratch_dir("strict");
        let path = dir.join("cut.docx");
        let body = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>kept</w:t></w:r></w:p><w:p><w:r><w:t>cut"#;
        let mut writer = OpcWriter::new(std::io::Cursor::new(Vec::new())).unwrap();
        writer
            .add_part(
                &PartName::new("/word/document.xml").unwrap(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                body.as_bytes(),
            )
            .unwrap();
        writer.add_package_rel(rel_types::OFFICE_DOCUMENT, "word/document.xml");
        std::fs::write(&path, writer.finish().unwrap().into_inner()).unwrap();
        let file = path.to_str().unwrap();

        for tool in ["extract", "info"] {
            let lenient = handle_tools_call(
                &json!(1),
                &json!({ "name": tool, "arguments": { "file_path": file } }),
            );
            assert!(lenient["result"]["isError"].is_null(), "{tool}: {lenient}");
            let strict = handle_tools_call(
                &json!(1),
                &json!({ "name": tool, "arguments": { "file_path": file, "strict": true } }),
            );
            assert_eq!(strict["result"]["isError"], json!(true), "{tool}: {strict}");
            let message = strict["result"]["content"][0]["text"].as_str().unwrap();
            assert!(
                message.contains("incomplete") && message.contains("word/document.xml"),
                "{message}"
            );
        }
        std::fs::remove_dir_all(&dir).ok();
    }

    /// `extract` and `info` leave the process's budgets as they found them,
    /// and `replace_text`, which edits the package's XML and reads none of
    /// them, is not cut short by tiny ones.
    #[test]
    fn test_tool_calls_restore_the_budgets_and_replace_text_ignores_them() {
        use office_oxide::limits::{max_image_bytes, max_text_chars};
        let dir = scratch_dir("restore");
        let path = dir.join("doc.docx");
        let text = "lorem ipsum ".repeat(200);
        write_docx(&path, &text);
        let arguments = json!({ "file_path": path.to_str().unwrap() });
        let _limits = TestLimits::bound_to(10_000_000_000, 10_000_000_000);
        for tool in ["extract", "info"] {
            let out =
                handle_tools_call(&json!(1), &json!({ "name": tool, "arguments": arguments }));
            assert!(out["result"]["isError"].is_null(), "{tool}: {out}");
            assert_eq!(
                (max_text_chars(), max_image_bytes()),
                (10_000_000_000, 10_000_000_000),
                "{tool} left the budgets lowered"
            );
        }
        office_oxide::limits::set_max_text_chars(3);
        let out = handle_tools_call(
            &json!(1),
            &json!({
                "name": "replace_text",
                "arguments": {
                    "file_path": path.to_str().unwrap(), "find": "ipsum", "replace": "dolor"
                }
            }),
        );
        let message = out["result"]["content"][0]["text"].as_str().unwrap();
        assert!(message.starts_with("replaced 200 occurrence(s)"), "{out}");
        office_oxide::limits::set_max_text_chars(10_000_000_000);
        std::fs::remove_dir_all(&dir).ok();
    }

    /// The size limit counts UTF-8 bytes, not characters: 100 CJK characters
    /// are 300 bytes and do not fit a 200 byte limit though 100 < 200.
    #[test]
    fn test_the_result_limit_counts_bytes_not_characters() {
        let _defaults = default_limits();
        let dir = scratch_dir("bytes");
        let path = dir.join("doc.docx");
        let text = "\u{4e2d}".repeat(100);
        write_docx(&path, &text);
        let doc = office_oxide::Document::open(&path).unwrap();
        assert_eq!(render(&doc, "text", 300).unwrap().len(), 300);
        let err = render(&doc, "text", 299).expect_err("300 bytes do not fit 299");
        assert!(err.contains("exceeds") && err.contains("bytes"), "{err}");
        let err = render(&doc, "text", 200).expect_err("100 characters, 300 bytes");
        assert!(err.contains("exceeds"), "{err}");
        std::fs::remove_dir_all(&dir).ok();
    }

    /// The README documented `extract` as DOCX/XLSX/PPTX-only with three
    /// formats, omitted `replace_text`, and promised a file size `info` did
    /// not return; the tool descriptions said DOCX/XLSX/PPTX too. The CLI
    /// has a clap-vs-README drift test; this is the MCP equivalent, driven
    /// off the live `tools/list` response.
    #[test]
    fn test_readme_documents_every_tool_parameter_and_format() {
        let readme =
            std::fs::read_to_string(concat!(env!("CARGO_MANIFEST_DIR"), "/README.md")).unwrap();
        let list = handle_tools_list(&json!(1));
        let tools = list["result"]["tools"].as_array().unwrap();
        assert!(!tools.is_empty());
        for tool in tools {
            let name = tool["name"].as_str().unwrap();
            let start = readme
                .find(&format!("### `{name}`"))
                .unwrap_or_else(|| panic!("README has no section for tool `{name}`"));
            let end = readme[start + 1..]
                .find("\n##")
                .map_or(readme.len(), |i| start + 1 + i);
            let section = &readme[start..end];
            let props = tool["inputSchema"]["properties"].as_object().unwrap();
            for (prop, schema) in props {
                assert!(
                    section.contains(&format!("`{prop}`")),
                    "README section for `{name}` does not document `{prop}`"
                );
                for value in schema["enum"].as_array().into_iter().flatten() {
                    let value = value.as_str().unwrap();
                    assert!(
                        section.contains(&format!("`{value}`")),
                        "README section for `{name}` does not list `{prop}` value `{value}`"
                    );
                }
            }
            // Tools that read documents read all six formats; say so.
            let description = tool["description"].as_str().unwrap();
            if name != "replace_text" {
                for fmt in ["DOCX", "XLSX", "PPTX", "DOC,", "XLS,", "PPT"] {
                    assert!(
                        description.contains(fmt),
                        "`{name}` description omits {fmt}: {description}"
                    );
                }
            }
        }
    }

    /// `info` returned only format/title/sections: the truncation warning
    /// the CLI prints never reached an agent, and author, subject,
    /// keywords and the dates were parsed but not reported.
    #[test]
    fn test_info_reports_metadata_and_the_truncation_warning() {
        let _defaults = default_limits();
        use office_oxide::format::DocumentFormat;
        use office_oxide::ir::Metadata;
        let mut metadata = Metadata {
            format: DocumentFormat::Doc,
            ..Default::default()
        };
        metadata.author = Some("Ada".into());
        metadata.subject = Some("Numbers".into());
        metadata.keywords = vec!["q3".into()];
        metadata.created = Some("2024-01-02T03:04:05Z".into());
        metadata.modified = Some("2024-02-03T04:05:06Z".into());
        metadata.text_truncated = true;
        metadata.warnings = vec!["skipped unreadable part /word/header1.xml".into()];
        let ir = office_oxide::DocumentIR {
            metadata,
            sections: Vec::new(),
        };
        let info = info_json(&ir, Some(42));
        assert_eq!(info["file_size"], json!(42));
        assert_eq!(info["author"], json!("Ada"));
        assert_eq!(info["subject"], json!("Numbers"));
        assert_eq!(info["keywords"], json!(["q3"]));
        assert_eq!(info["created"], json!("2024-01-02T03:04:05Z"));
        assert_eq!(info["modified"], json!("2024-02-03T04:05:06Z"));
        assert_eq!(info["text_truncated"], json!(true));
        let warnings = info["warnings"].as_array().unwrap();
        assert_eq!(warnings.len(), 2);
        assert!(warnings[0].as_str().unwrap().contains("incomplete"));
        assert_eq!(warnings[1], json!("skipped unreadable part /word/header1.xml"));

        let clean = office_oxide::DocumentIR {
            metadata: Metadata::default(),
            sections: Vec::new(),
        };
        assert_eq!(info_json(&clean, None)["warnings"], json!([]));
    }

    /// An empty `find` interleaved the replacement between every character
    /// and — because `output_path` defaults to `file_path` — overwrote the
    /// original with the result while reporting success.
    #[test]
    fn test_replace_text_with_empty_find_is_an_error_and_leaves_the_file_alone() {
        let _defaults = default_limits();
        let dir = scratch_dir("empty_find");
        let path = dir.join("doc.docx");
        write_docx(&path, "Hello world");
        let before = std::fs::read(&path).unwrap();

        let out = handle_tools_call(
            &json!(1),
            &json!({
                "name": "replace_text",
                "arguments": {"file_path": path.to_str().unwrap(), "find": "", "replace": "X"}
            }),
        );
        assert_eq!(out["result"]["isError"], json!(true), "{out}");
        let text = out["result"]["content"][0]["text"].as_str().unwrap();
        assert!(text.contains("empty"), "the error must say why: {text}");
        assert!(std::fs::read(&path).unwrap() == before, "the original was rewritten");
        std::fs::remove_dir_all(&dir).ok();
    }
}

#[cfg(test)]
mod info_tests {
    use super::*;

    #[test]
    fn test_info_reports_every_set_document_property() {
        let ir = office_oxide::DocumentIR {
            metadata: office_oxide::ir::Metadata {
                format: office_oxide::DocumentFormat::Docx,
                title: Some("T".into()),
                author: Some("A".into()),
                subject: Some("S".into()),
                keywords: vec!["k1".into(), "k 2".into()],
                created: Some("2024-01-01T00:00:00Z".into()),
                modified: Some("2024-02-01T00:00:00Z".into()),
                last_modified_by: Some("L".into()),
                company: Some("C".into()),
                ..Default::default()
            },
            sections: Vec::new(),
        };
        let v = info_json(&ir, None);
        assert_eq!(v["author"], "A");
        assert_eq!(v["subject"], "S");
        assert_eq!(v["keywords"], json!(["k1", "k 2"]));
        assert_eq!(v["created"], "2024-01-01T00:00:00Z");
        assert_eq!(v["modified"], "2024-02-01T00:00:00Z");
        assert_eq!(v["last_modified_by"], "L");
        assert_eq!(v["company"], "C");
        assert!(v.get("manager").is_none(), "absent properties are omitted");
    }
}
