# office-oxide MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude, Cursor, and other AI assistants the ability to read — and edit the text of — Office documents locally.

## Supported Formats

DOCX, XLSX, PPTX, DOC, XLS, PPT

## Installation

```bash
# From crates.io
cargo install office_oxide_mcp

# Pre-built binaries (via cargo-binstall)
cargo binstall office_oxide_mcp

# From source
cargo install --path crates/office_oxide_mcp
```

## Configuration

### Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office-oxide": {
      "command": "office-oxide-mcp"
    }
  }
}
```

### Claude Code

Add to your `.claude/settings.json`:

```json
{
  "mcpServers": {
    "office-oxide": {
      "command": "office-oxide-mcp"
    }
  }
}
```

## Tools

### `extract`

Extract content from an Office document (DOCX, XLSX, PPTX, DOC, XLS, PPT).

| Parameter | Type | Description |
|-----------|------|-------------|
| `file_path` | string | Path to the document |
| `format` | string | Output format: `text` (default), `markdown`, `markdown-with-images`, `html`, or `ir` |
| `strict` | boolean | Read untrusted input: tighter limits, and an error instead of a partial result when part of the document cannot be read or a limit cut it (default `false`) |

`markdown-with-images` embeds each image inline as `[image-base64:...]` at its
position in the flow; `ir` is the document IR as JSON. A result larger than
32 MiB (UTF-8 bytes, not characters) is refused with an error rather than
truncated — use a more compact format, or the `office-oxide` CLI, which streams
its output. Only a spreadsheet is written as it is produced, so only there does
the limit stop the work; every other format is built whole before the limit
can refuse it, and its memory is bounded by the library's text and image
budgets. `extract` and `info` read the document under per-call options that cap those
budgets at 32 Mi characters and 24 MiB (nothing process-wide is changed;
`replace_text` is not affected): a document with more text is cut with the library's truncation
notice, which `info` reports as a warning. A request line longer than 16 MiB
is rejected with a JSON-RPC `-32600` error and skipped.

### `replace_text`

Replace text in a DOCX or PPTX document, preserving every other part of the
file.

| Parameter | Type | Description |
|-----------|------|-------------|
| `file_path` | string | Path to the document |
| `find` | string | Text to search for (must not be empty) |
| `replace` | string | Replacement text |
| `output_path` | string | Where to write the result (default: overwrite `file_path`) |

The result is written to a temporary file and renamed into place, so a failed
save never leaves a truncated document behind.

### `info`

Get document metadata: format, file size, title, the full document properties
(author, subject, keywords, dates, …), a `warnings` list (e.g. when text
extraction is known to be incomplete), and the section list.

| Parameter | Type | Description |
|-----------|------|-------------|
| `file_path` | string | Path to the document |
| `strict` | boolean | As for `extract` |

## Protocol

JSON-RPC 2.0 over stdin/stdout, compatible with MCP protocol version `2024-11-05`.

## License

MIT OR Apache-2.0
