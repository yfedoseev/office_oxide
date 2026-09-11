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
                    "description": "Extract content from an Office document (DOCX, XLSX, PPTX)",
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
                                "description": "Text to search for"
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
                    "description": "Get metadata about an Office document",
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

pub fn handle_tools_call(id: &Value, params: &Value) -> Value {
    let tool_name = params["name"].as_str().unwrap_or("");
    let arguments = &params["arguments"];

    match tool_name {
        "extract" => call_extract(id, arguments),
        "replace_text" => call_replace_text(id, arguments),
        "info" => call_info(id, arguments),
        _ => error_response(id, -32601, &format!("unknown tool: {tool_name}")),
    }
}

fn call_extract(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, -32602, "missing file_path");
    };
    let format = args["format"].as_str().unwrap_or("text");

    let doc = match office_oxide::Document::open(file_path) {
        Ok(d) => d,
        Err(e) => return tool_error(id, &e.to_string()),
    };

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
        "ir" => match serde_json::to_string_pretty(&doc.to_ir()) {
            Ok(s) => s,
            Err(e) => return tool_error(id, &e.to_string()),
        },
        other => return tool_error(id, &format!("unknown format: {other}")),
    };

    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{ "type": "text", "text": content }]
        }
    })
}

fn call_replace_text(id: &Value, args: &Value) -> Value {
    let Some(file_path) = args["file_path"].as_str() else {
        return error_response(id, -32602, "missing file_path");
    };
    let Some(find) = args["find"].as_str() else {
        return error_response(id, -32602, "missing find");
    };
    let Some(replace) = args["replace"].as_str() else {
        return error_response(id, -32602, "missing replace");
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
        return error_response(id, -32602, "missing file_path");
    };

    let doc = match office_oxide::Document::open(file_path) {
        Ok(d) => d,
        Err(e) => return tool_error(id, &e.to_string()),
    };

    let ir = doc.to_ir();
    let info = json!({
        "format": format!("{:?}", ir.metadata.format),
        "title": ir.metadata.title,
        "sections": ir.sections.len(),
        "section_names": ir.sections.iter().map(|s| s.title.clone()).collect::<Vec<_>>(),
    });

    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{ "type": "text", "text": info.to_string() }]
        }
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
