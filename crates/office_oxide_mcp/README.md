# office-oxide MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude, Cursor, and other AI assistants the ability to read Office documents locally.

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

Extract content from an Office document.

| Parameter | Type | Description |
|-----------|------|-------------|
| `file_path` | string | Path to the document |
| `format` | string | Output format: `text` (default), `markdown`, or `ir` |

### `info`

Get document metadata (format, file size).

| Parameter | Type | Description |
|-----------|------|-------------|
| `file_path` | string | Path to the document |

## Protocol

JSON-RPC 2.0 over stdin/stdout, compatible with MCP protocol version `2024-11-05`.

## License

MIT OR Apache-2.0
