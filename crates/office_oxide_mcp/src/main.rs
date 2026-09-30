//! `office-oxide-mcp` — Model Context Protocol server for office_oxide.
//!
//! Speaks JSON-RPC 2.0 over stdin/stdout. Exposes three tools:
//! `extract` (text / markdown / html / ir from a DOCX/XLSX/PPTX/DOC/
//! XLS/PPT file), `replace_text` (DOCX/PPTX text replacement) and `info`
//! (format detection + metadata).

#![warn(missing_docs)]

mod protocol;

use std::io::{self, BufRead, Write};

/// Handles one JSON-RPC request object. Returns `None` for notifications
/// (a well-formed request with no `id` member, or one of the `initialized`
/// notification methods), which never get a response — including inside a
/// batch.
fn handle_request(request: &serde_json::Value) -> Option<serde_json::Value> {
    let id = &request["id"];
    // JSON-RPC 2.0 §4.1: a Request object without an `id` member is a
    // Notification, and "the Server MUST NOT reply to a Notification" —
    // whatever its method. Only a well-formed request qualifies: a malformed
    // one is answered below with -32600 and a null id, as §5 requires.
    let is_notification = request.as_object().is_some_and(|o| !o.contains_key("id"));

    if request.get("jsonrpc") != Some(&serde_json::Value::String("2.0".to_string())) {
        return Some(serde_json::json!({
            "jsonrpc": "2.0",
            "id": id,
            "error": {
                "code": -32600,
                "message": "Invalid Request: missing or invalid \"jsonrpc\" field, expected \"2.0\""
            }
        }));
    }

    let method_value = &request["method"];
    let method = match method_value {
        serde_json::Value::String(m) => m.as_str(),
        _ => {
            return Some(serde_json::json!({
                "jsonrpc": "2.0",
                "id": id,
                "error": { "code": -32600, "message": "Invalid Request: \"method\" must be a string" }
            }));
        },
    };

    if is_notification {
        // Nothing this server offers is worth doing without a reply: the
        // only notifications a client sends are lifecycle signals
        // (`initialized`, `cancelled`, `progress`), and every tool call is a
        // request whose result the client needs. So acknowledge none.
        return None;
    }

    let params = &request["params"];

    match method {
        "initialize" => Some(protocol::handle_initialize(id)),
        "tools/list" => Some(protocol::handle_tools_list(id)),
        "tools/call" => Some(protocol::handle_tools_call(id, params)),
        "notifications/initialized" | "initialized" => None,
        _ => Some(serde_json::json!({
            "jsonrpc": "2.0",
            "id": id,
            "error": { "code": -32601, "message": format!("unknown method: {method}") }
        })),
    }
}

/// Handles one already-parsed line of input (a single request object, or a
/// JSON-RPC batch array of them). Returns `None` when nothing should be
/// written to stdout (a lone notification, or a batch made up entirely of
/// notifications).
fn handle_line(request: serde_json::Value) -> Option<serde_json::Value> {
    match request {
        serde_json::Value::Array(items) => {
            if items.is_empty() {
                // JSON-RPC 2.0 §6: an empty array is an invalid Request, and
                // the reply is a single Response object — not an array.
                return Some(serde_json::json!({
                    "jsonrpc": "2.0",
                    "id": null,
                    "error": { "code": -32600, "message": "Invalid Request: empty batch" }
                }));
            }
            let responses: Vec<serde_json::Value> =
                items.iter().filter_map(handle_request).collect();

            if responses.is_empty() {
                None
            } else {
                Some(serde_json::Value::Array(responses))
            }
        },
        other => handle_request(&other),
    }
}

fn main() {
    let stdin = io::stdin();
    let mut stdout = io::stdout();

    for line in stdin.lock().lines() {
        let line = match line {
            Ok(l) => l,
            Err(_) => break,
        };

        if line.trim().is_empty() {
            continue;
        }

        let request: serde_json::Value = match serde_json::from_str(&line) {
            Ok(v) => v,
            Err(e) => {
                let err = serde_json::json!({
                    "jsonrpc": "2.0",
                    "id": null,
                    "error": { "code": -32700, "message": format!("parse error: {e}") }
                });
                let _ = writeln!(stdout, "{err}");
                let _ = stdout.flush();
                continue;
            },
        };

        if let Some(response) = handle_line(request) {
            let _ = writeln!(stdout, "{response}");
            let _ = stdout.flush();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn test_batch_of_two_valid_requests_returns_an_array_of_two_responses() {
        let batch = json!([
            {"jsonrpc": "2.0", "id": 200, "method": "tools/list"},
            {"jsonrpc": "2.0", "id": 201, "method": "initialize"}
        ]);
        let out = handle_line(batch).expect("batch should produce a response");
        let arr = out.as_array().expect("response must be a JSON array");
        assert_eq!(arr.len(), 2, "both batch entries should get a response: {arr:?}");
        assert_eq!(arr[0]["id"], json!(200));
        assert!(arr[0]["result"]["tools"].is_array());
        assert_eq!(arr[1]["id"], json!(201));
        assert!(arr[1]["result"]["protocolVersion"].is_string());
    }

    /// JSON-RPC 2.0 §6: an empty batch array is answered with a *single*
    /// Response object, not an array holding one.
    #[test]
    fn test_empty_batch_returns_a_single_invalid_request_error() {
        let out = handle_line(json!([])).expect("empty batch should still respond");
        assert!(out.is_object(), "an empty batch must get a single object, not an array: {out}");
        assert_eq!(out["error"]["code"], json!(-32600));
        assert_eq!(out["id"], json!(null));
    }

    /// JSON-RPC 2.0 §4.1: "The Server MUST NOT reply to a Notification" — any
    /// request without an `id` member, not only the `initialized` one. The
    /// standard `notifications/cancelled` used to draw a -32601 reply.
    #[test]
    fn test_no_notification_gets_a_reply() {
        for method in [
            "notifications/cancelled",
            "notifications/progress",
            "bogus",
            "tools/list",
        ] {
            let note = json!({"jsonrpc": "2.0", "method": method, "params": {}});
            assert!(handle_line(note.clone()).is_none(), "{method} notification was answered");
            assert!(
                handle_line(json!([note])).is_none(),
                "{method} notification answered in a batch"
            );
        }
    }

    /// An explicit `"id": null` is a request (discouraged, but not a
    /// notification), so it is still answered.
    #[test]
    fn test_a_null_id_is_a_request_not_a_notification() {
        let out = handle_line(json!({"jsonrpc": "2.0", "id": null, "method": "tools/list"}))
            .expect("id: null is still a request");
        assert!(out["result"]["tools"].is_array());
    }

    /// A malformed object without an `id` is an invalid Request, not a
    /// notification: it is answered with -32600 and `id: null`.
    #[test]
    fn test_invalid_request_without_id_is_still_answered() {
        let out =
            handle_line(json!({"method": "tools/list"})).expect("invalid request is answered");
        assert_eq!(out["error"]["code"], json!(-32600));
        assert_eq!(out["id"], json!(null));
    }

    /// `tools/call` with missing or malformed params is -32602 Invalid params
    /// (JSON-RPC 2.0 §5.1; MCP uses the same code for an unknown tool), not
    /// -32601 Method not found — the method exists.
    #[test]
    fn test_tools_call_with_bad_params_is_invalid_params() {
        for params in [
            json!(null),
            json!({}),
            json!({"name": 7}),
            json!({"name": "no_such_tool", "arguments": {}}),
            json!({"name": "extract"}),
            json!({"name": "extract", "arguments": {"format": "text"}}),
            json!({"name": "extract", "arguments": {"file_path": 3}}),
            json!({"name": "replace_text", "arguments": {"file_path": "a.docx", "find": "x"}}),
            json!({"name": "info", "arguments": "a.docx"}),
        ] {
            let mut req = json!({"jsonrpc": "2.0", "id": 3, "method": "tools/call"});
            if !params.is_null() {
                req["params"] = params.clone();
            }
            let out = handle_line(req).unwrap();
            assert_eq!(out["error"]["code"], json!(-32602), "params {params} gave {out}");
            assert_eq!(out["id"], json!(3));
        }
    }

    #[test]
    fn test_batch_notification_gets_no_entry_in_the_response_array() {
        let batch = json!([
            {"jsonrpc": "2.0", "id": 1, "method": "initialize"},
            {"jsonrpc": "2.0", "method": "notifications/initialized"}
        ]);
        let out = handle_line(batch).expect("batch has one real request");
        let arr = out.as_array().unwrap();
        assert_eq!(arr.len(), 1, "the notification must not appear in the response: {arr:?}");
        assert_eq!(arr[0]["id"], json!(1));
    }

    #[test]
    fn test_batch_of_only_notifications_produces_no_output() {
        let batch = json!([{"jsonrpc": "2.0", "method": "notifications/initialized"}]);
        assert!(handle_line(batch).is_none());
    }

    #[test]
    fn test_missing_jsonrpc_field_is_invalid_request() {
        let out = handle_line(json!({"id": 1, "method": "tools/list"})).unwrap();
        assert_eq!(out["error"]["code"], json!(-32600));
        assert_eq!(out["id"], json!(1));
    }

    #[test]
    fn test_wrong_jsonrpc_version_is_invalid_request() {
        let out =
            handle_line(json!({"jsonrpc": "1.0", "id": 400, "method": "tools/list"})).unwrap();
        assert_eq!(out["error"]["code"], json!(-32600));
        assert_eq!(out["id"], json!(400));
    }

    #[test]
    fn test_non_string_method_is_invalid_request_not_unknown_method() {
        let out = handle_line(json!({"jsonrpc": "2.0", "id": 5, "method": 12345})).unwrap();
        assert_eq!(out["error"]["code"], json!(-32600));
    }

    #[test]
    fn test_single_well_formed_request_still_works() {
        let out = handle_line(json!({"jsonrpc": "2.0", "id": 1, "method": "initialize"})).unwrap();
        assert_eq!(out["id"], json!(1));
        assert!(out["result"]["protocolVersion"].is_string());
    }

    #[test]
    fn test_lone_notification_produces_no_output() {
        assert!(
            handle_line(json!({"jsonrpc": "2.0", "method": "notifications/initialized"})).is_none()
        );
    }

    #[test]
    fn test_unknown_method_still_uses_method_not_found() {
        let out = handle_line(json!({"jsonrpc": "2.0", "id": 9, "method": "bogus"})).unwrap();
        assert_eq!(out["error"]["code"], json!(-32601));
    }
}
