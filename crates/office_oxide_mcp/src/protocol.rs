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

/// Largest `extract` result the server will return, in bytes.
///
/// A tool result is one JSON string held in memory (and then escaped into a
/// second copy for the response line), and no MCP client can put tens of
/// megabytes in front of a model anyway. The CLI documents what an unbounded
/// buffer did on large workbooks — multi-gigabyte RSS and OOM kills — so an
/// oversized result is refused with a message saying so, rather than
/// silently truncated.
const MAX_EXTRACT_BYTES: usize = 32 * 1024 * 1024;

fn call_extract(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, INVALID_PARAMS, "missing file_path");
    };
    let format = args["format"].as_str().unwrap_or("text");

    let doc = match office_oxide::Document::open(file_path) {
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
fn render(doc: &office_oxide::Document, format: &str, limit: usize) -> Result<String, String> {
    let content = match format {
        "text" => doc.plain_text(),
        "markdown" => doc.to_markdown(),
        // Images are dropped from plain markdown entirely; this keeps both
        // their content and their position in one self-contained string.
        "markdown-with-images" => {
            use office_oxide::ir_render::{ImageEmbed, MarkdownOptions};
            doc.to_markdown_with(MarkdownOptions {
                image_embed: ImageEmbed::Base64,
            })
        },
        "html" => doc.to_html(),
        // The IR is serialised through a bounded writer, so an oversized
        // document stops at the limit instead of building the whole string.
        "ir" => {
            let mut out = BoundedBuf {
                buf: Vec::new(),
                limit,
                exceeded: false,
            };
            match serde_json::to_writer_pretty(&mut out, &doc.to_ir()) {
                Ok(()) => String::from_utf8(out.buf).map_err(|e| e.to_string())?,
                Err(_) if out.exceeded => return Err(too_large(format, limit)),
                Err(e) => return Err(e.to_string()),
            }
        },
        other => return Err(format!("unknown format: {other}")),
    };
    if content.len() > limit {
        return Err(too_large(format, limit));
    }
    Ok(content)
}

fn too_large(format: &str, limit: usize) -> String {
    format!(
        "the {format} output of this document exceeds the {} MiB limit for one tool result; \
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

    let doc = match office_oxide::Document::open(file_path) {
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
const TRUNCATION_WARNING: &str = "text extraction is incomplete — the source file's own \
     structure disagrees with itself about how much text there is, and the gap could not be \
     safely recovered";

/// The `info` tool's result.
///
/// `metadata` is the IR's own serde form, so every document property the
/// library parses (author, subject, keywords, dates, …) reaches the agent
/// without this tool having to be kept in step by hand. `warnings` carries
/// the truncation signal the CLI already surfaced: without it an agent had
/// no way to learn that an extraction was incomplete.
fn info_json(ir: &office_oxide::DocumentIR, file_size: Option<u64>) -> Value {
    let mut warnings = Vec::new();
    if ir.metadata.text_truncated {
        warnings.push(TRUNCATION_WARNING);
    }
    json!({
        "format": format!("{:?}", ir.metadata.format),
        "title": ir.metadata.title,
        "file_size": file_size,
        "metadata": ir.metadata,
        "warnings": warnings,
        "sections": ir.sections.len(),
        "section_names": ir.sections.iter().map(|s| s.title.clone()).collect::<Vec<_>>(),
    })
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
        let ir = office_oxide::DocumentIR {
            metadata,
            sections: Vec::new(),
            defined_names: Vec::new(),
        };
        let info = info_json(&ir, Some(42));
        assert_eq!(info["file_size"], json!(42));
        assert_eq!(info["metadata"]["author"], json!("Ada"));
        assert_eq!(info["metadata"]["subject"], json!("Numbers"));
        assert_eq!(info["metadata"]["keywords"], json!(["q3"]));
        assert_eq!(info["metadata"]["created"], json!("2024-01-02T03:04:05Z"));
        assert_eq!(info["metadata"]["modified"], json!("2024-02-03T04:05:06Z"));
        assert_eq!(info["metadata"]["text_truncated"], json!(true));
        let warnings = info["warnings"].as_array().unwrap();
        assert_eq!(warnings.len(), 1);
        assert!(warnings[0].as_str().unwrap().contains("incomplete"));

        let clean = office_oxide::DocumentIR {
            metadata: Metadata::default(),
            sections: Vec::new(),
            defined_names: Vec::new(),
        };
        assert_eq!(info_json(&clean, None)["warnings"], json!([]));
    }

    /// An empty `find` interleaved the replacement between every character
    /// and — because `output_path` defaults to `file_path` — overwrote the
    /// original with the result while reporting success.
    #[test]
    fn test_replace_text_with_empty_find_is_an_error_and_leaves_the_file_alone() {
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
