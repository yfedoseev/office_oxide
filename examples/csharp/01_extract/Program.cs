// 01_extract — Self-contained extract demo using the C# bindings.
//
// Creates a DOCX from Markdown (written to a temp file), opens it with
// Document.Open, and prints the format, plain text, and Markdown output.
// Exit 0 on success.
//
// Run: dotnet run --project examples/csharp/01_extract/01_extract.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using OfficeOxide;

const string markdown = """
    # Office Oxide Extract Demo

    This document was created from Markdown and parsed back via C# bindings.

    ## Features

    - Plain text extraction
    - Markdown conversion
    - IR (intermediate representation) access
    """;

var tmpPath = Path.Combine(Path.GetTempPath(), $"oo_01_extract_{Environment.TickCount64}.docx");

try
{
    Document.CreateFromMarkdown(markdown, "docx", tmpPath);

    using var doc = Document.Open(tmpPath);

    var format = doc.Format;
    var text = doc.PlainText();
    var md = doc.ToMarkdown();

    Console.WriteLine($"format: {format}");
    Console.WriteLine("--- plain text (first 200 chars) ---");
    Console.WriteLine(text.Length > 200 ? text[..200] : text);
    Console.WriteLine($"--- markdown length: {md.Length} chars ---");

    if (format != "docx") throw new Exception($"expected format=docx, got {format}");
    if (string.IsNullOrEmpty(text)) throw new Exception("plain text is empty");
    if (!text.Contains("Office Oxide Extract Demo")) throw new Exception("heading missing");

    Console.WriteLine("\nAll checks passed.");
}
finally
{
    if (File.Exists(tmpPath)) File.Delete(tmpPath);
}
