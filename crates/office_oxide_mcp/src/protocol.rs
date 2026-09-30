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
    let info = info_json(&ir);

    json!({
        "jsonrpc": "2.0",
        "id": id,
        "result": {
            "content": [{ "type": "text", "text": info.to_string() }]
        }
    })
}

/// The `info` tool's payload: format, every document property that is
/// set, the custom properties, the content flags and the section list.
fn info_json(ir: &office_oxide::DocumentIR) -> Value {
    let meta = &ir.metadata;
    let mut info = json!({
        "format": format!("{:?}", meta.format),
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
        "sections": ir.sections.len(),
        "section_names": ir.sections.iter().map(|s| s.title.clone()).collect::<Vec<_>>(),
    });
    // Absent properties are omitted rather than reported as null.
    if let Some(map) = info.as_object_mut() {
        map.retain(|_, v| !v.is_null());
    }
    info
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
            defined_names: Vec::new(),
        };
        let v = info_json(&ir);
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
