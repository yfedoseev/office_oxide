# Getting Started with office_oxide (C#/.NET)

`OfficeOxide` is the .NET binding for the Rust `office_oxide` library. It gives .NET applications fast, allocation-conscious parsing, conversion, and editing of DOCX / XLSX / PPTX / DOC / XLS / PPT — with AOT-compatible `LibraryImport` P/Invoke under the hood.

## Installation

```bash
dotnet add package OfficeOxide --version 0.1.0
```

Requires .NET 8 or .NET 10. The NuGet package ships prebuilt native libraries for `win-x64`, `linux-x64`, `linux-arm64`, `osx-x64`, `osx-arm64` under `runtimes/<rid>/native/`. `dotnet publish` places the right one next to your binary automatically.

## Quickstart

Extract plain text from a DOCX file:

```csharp
using OfficeOxide;

using var doc = Document.Open("report.docx");
Console.WriteLine(doc.PlainText());
```

## Core API

### `Document`

`Document` is the read-only handle; dispose it (preferably with `using`) to free native memory.

```csharp
using OfficeOxide;

using var doc = Document.Open("file.xlsx");

Console.WriteLine(doc.Format);        // "xlsx"
Console.WriteLine(doc.PlainText());
Console.WriteLine(doc.ToMarkdown());
Console.WriteLine(doc.ToHtml());
Console.WriteLine(doc.ToIrJson());

// Save / convert — the target format is inferred from the extension.
doc.SaveAs("file.docx");
```

Async wrapper for blocking IO:

```csharp
using var doc = await Document.OpenAsync("huge.pptx", ct);
```

Open from bytes (no temp file needed):

```csharp
byte[] data = File.ReadAllBytes("report.docx");
using var doc = Document.FromBytes(data, "docx");
```

`format` must be `"docx"`, `"xlsx"`, `"pptx"`, `"doc"`, `"xls"`, or `"ppt"`.

Static helpers for one-shot calls:

```csharp
string text = OfficeOxide.ExtractText("file.docx");
string md   = OfficeOxide.ToMarkdown("file.pptx");
string html = OfficeOxide.ToHtml("file.xlsx");

string? fmt = Document.DetectFormat("mystery.bin"); // null if unsupported
Console.WriteLine(Document.Version);                // "0.1.0"
```

### `EditableDocument`

Editing preserves every unmodified OPC part (images, charts, relationships) on save. DOCX, XLSX, and PPTX only.

```csharp
using OfficeOxide;

using var ed = EditableDocument.Open("template.docx");
long n = ed.ReplaceText("{{name}}", "Alice");
Console.WriteLine($"{n} replacements");
ed.Save("out.docx");
```

`ReplaceText` returns the replacement count (0 on XLSX — use `SetCell` instead).

## Editing Examples

### Replace text in DOCX / PPTX

```csharp
using var ed = EditableDocument.Open("slides.pptx");
ed.ReplaceText("Q3", "Q4");
ed.ReplaceText("2024", "2025");

byte[] bytes = ed.SaveToBytes();            // for streaming / upload
File.WriteAllBytes("slides_q4.pptx", bytes);
```

### Set XLSX cells (four overloads)

```csharp
using var wb = EditableDocument.Open("budget.xlsx");

wb.SetCell(0u, "A1", "Total");        // string overload
wb.SetCell(0u, "B1", 42.5);           // double overload
wb.SetCell(0u, "C1", true);           // bool overload
wb.SetCellEmpty(0u, "D1");            // clear a cell

wb.Save("budget.xlsx");
```

`sheetIndex` is zero-based; `cellRef` is standard spreadsheet notation (`A1`, `AA12`).

## Advanced

### Format-agnostic IR

`ToIrJson()` returns a JSON string matching the Rust `DocumentIR` shape. Deserialize into your own record types:

```csharp
using System.Text.Json;

using var doc = Document.Open("report.docx");
string json = doc.ToIrJson();

using var ir = JsonDocument.Parse(json);
foreach (var section in ir.RootElement.GetProperty("sections").EnumerateArray())
{
    if (section.TryGetProperty("title", out var t) && t.ValueKind != JsonValueKind.Null)
        Console.WriteLine(t.GetString());
}
```

### Bytes-based workflow

```csharp
using var http = new HttpClient();
byte[] data = await http.GetByteArrayAsync("https://example.com/file.docx");
using var doc = Document.FromBytes(data, "docx");
Console.WriteLine(doc.ToMarkdown());
```

### Legacy formats (DOC / XLS / PPT)

The legacy parsers are first-class. Extension detection routes automatically; `SaveAs` transparently converts through the IR:

```csharp
using var legacy = Document.Open("old.xls");
legacy.SaveAs("modern.xlsx");
```

### AOT / trimming

The project sets `IsAotCompatible=true` and `IsTrimmable=true`. All P/Invoke uses `LibraryImport` source generators, so `dotnet publish -c Release -p:PublishAot=true` yields a single self-contained executable.

## Error Handling

Failures throw `OfficeOxideException` with a typed `Code` property:

```csharp
try
{
    using var doc = Document.Open("missing.docx");
}
catch (OfficeOxideException ex)
{
    Console.WriteLine($"code={ex.Code} op={ex.Operation}");
}
```

Calling methods on a disposed handle throws `ObjectDisposedException`.

### Error codes

| Code | Name | Meaning |
|---:|---|---|
| 0 | `Ok` | success |
| 1 | `InvalidArg` | null / empty / wrong format string |
| 2 | `Io` | filesystem error |
| 3 | `Parse` | malformed document |
| 4 | `Extraction` | parsed but rendering failed |
| 5 | `Internal` | bug — please file an issue |
| 6 | `Unsupported` | extension / feature not supported |

## Troubleshooting

| Symptom | Fix |
|---|---|
| `DllNotFoundException: office_oxide` | The native library wasn't copied next to your binary. Run `dotnet publish` (it picks the `runtimes/<rid>/native/` file automatically) rather than raw `dotnet build`. |
| `BadImageFormatException` on Windows | CPU architecture mismatch — deploy the matching `win-x64` or `win-arm64` build. |
| `OfficeOxideException` with code `Unsupported` on `.doc` | Ensure the extension is lowercase or pass the format explicitly via `FromBytes`. |
| Trimmed output missing symbols | Add `OfficeOxide` to `<TrimmerRootAssembly>`. The package itself is trim-safe, but reflection-heavy callers may strip it. |
| macOS "cannot be opened because the developer cannot be verified" | Run `xattr -d com.apple.quarantine /path/to/liboffice_oxide.dylib` or sign the bundle. |

## Links

- Binding source: `csharp/OfficeOxide/Document.cs`, `csharp/OfficeOxide/EditableDocument.cs`
- Native methods: `csharp/OfficeOxide/NativeMethods.cs`
- C header: `include/office_oxide_c/office_oxide.h`
- Package on NuGet: https://www.nuget.org/packages/OfficeOxide
- GitHub: https://github.com/yfedoseev/office_oxide
