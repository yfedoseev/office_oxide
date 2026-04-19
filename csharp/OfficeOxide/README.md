# OfficeOxide (.NET)

Idiomatic C# bindings for [office_oxide](https://github.com/yfedoseev/office_oxide) — a fast Rust library for parsing, converting, and editing Microsoft Office documents (DOCX, XLSX, PPTX, DOC, XLS, PPT).

- `IDisposable` / `using` pattern for native handles.
- `LibraryImport` source generator — NativeAOT-compatible, trim-safe.
- `net8.0` and `net10.0` target frameworks.
- Async helpers (`Document.OpenAsync`) that offload to the thread pool.

## Install

```bash
dotnet add package OfficeOxide
```

The NuGet package bundles pre-built native libraries for Windows x64/arm64,
Linux x64/arm64, and macOS x64/arm64.

## Usage

```csharp
using OfficeOxide;

using var doc = Document.Open("report.docx");
Console.WriteLine(doc.Format);        // "docx"
Console.WriteLine(doc.PlainText());
Console.WriteLine(doc.ToMarkdown());
```

### Editing

```csharp
using var ed = EditableDocument.Open("template.docx");
ed.ReplaceText("{{NAME}}", "Alice");
ed.Save("out.docx");
```

### Spreadsheet cells

```csharp
using var ed = EditableDocument.Open("report.xlsx");
ed.SetCell(0, "A1", "Revenue");
ed.SetCell(0, "B1", 12345.67);
ed.SetCell(0, "C1", true);
ed.Save("report.edited.xlsx");
```

### One-shot helpers

```csharp
string text = OfficeOxide.ExtractText("report.docx");
string md   = OfficeOxide.ToMarkdown("report.docx");
string html = OfficeOxide.ToHtml("report.docx");
```

## License

MIT OR Apache-2.0
