// 01_extract: read a document as plain text, Markdown and IR JSON.
//
// Run: dotnet run --project examples/csharp/01_extract/01_extract.csproj
using System.Text.Json;
using OfficeOxide;

const string Markdown = """
    # Quarterly Report

    Revenue grew in every region.

    | Region | Q1  | Q2  |
    |--------|-----|-----|
    | North  | 120 | 135 |
    | South  | 95  | 110 |

    """;

static void Check(bool condition, string message)
{
    if (condition) return;
    Console.Error.WriteLine($"FAIL: {message}");
    Environment.Exit(1);
}

static void Has(string text, string expected) =>
    Check(text.Contains(expected), $"\"{expected}\" is missing from the text");

// CreateFromMarkdown writes to a path, so the file lives in a private temp
// file just long enough to read its bytes back.
static byte[] MarkdownToBytes(string markdown, string format)
{
    var path = Path.Combine(Path.GetTempPath(), $"oo_example_{Guid.NewGuid():N}.{format}");
    try
    {
        Document.CreateFromMarkdown(markdown, format, path);
        return File.ReadAllBytes(path);
    }
    finally
    {
        File.Delete(path);
    }
}

static byte[] BuildDocxBytes(string title = "Quarterly Report") =>
    MarkdownToBytes(Markdown.Replace("Quarterly Report", title), "docx");

Console.WriteLine("step 1: build a DOCX in memory");
var data = BuildDocxBytes();

Console.WriteLine("step 2: open it from bytes");
using var doc = Document.FromBytes(data, "docx");
Console.WriteLine($"  format: {doc.Format}");
Check(doc.Format == "docx", "the opened format differs from the requested one");

Console.WriteLine("step 3: extract plain text, Markdown and IR JSON");
var text = doc.PlainText();
var markdown = doc.ToMarkdown();
var irJson = doc.ToIrJson();
Console.WriteLine($"  chars: {text.Length}");

Console.WriteLine("step 4: check the output");
Has(text, "Revenue grew in every region.");
Has(markdown, "Quarterly Report");
Check(JsonDocument.Parse(irJson).RootElement.ValueKind == JsonValueKind.Object, "the IR JSON is not an object");

Console.WriteLine("OK 01_extract");
