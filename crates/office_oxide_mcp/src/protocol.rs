use serde_json::{json, Value};

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
                                "enum": ["text", "markdown", "ir"],
                                "description": "Output format (default: text)"
                            }
                        },
                        "required": ["file_path"]
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
        "ir" => match serde_json::to_string_pretty(&ir_to_json(&doc.to_ir())) {
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

fn ir_to_json(ir: &office_oxide::DocumentIR) -> Value {
    json!({
        "metadata": {
            "format": format!("{:?}", ir.metadata.format),
            "title": ir.metadata.title,
        },
        "sections": ir.sections.iter().map(|s| {
            json!({
                "title": s.title,
                "element_count": s.elements.len(),
            })
        }).collect::<Vec<_>>(),
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
