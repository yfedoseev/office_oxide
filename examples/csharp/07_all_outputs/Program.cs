// 07_all_outputs: render one document as plain text, Markdown, HTML and IR JSON.
//
// Run: dotnet run --project examples/csharp/07_all_outputs/07_all_outputs.csproj
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

Console.WriteLine("step 3: render every output");
var outputs = new (string Name, string Value)[]
{
    ("plain", doc.PlainText()),     // words only: no markup, no structure
    ("markdown", doc.ToMarkdown()), // headings and pipe tables survive as syntax
    ("html", doc.ToHtml()),         // a fragment (no <html> wrapper), tables as <table>
    ("ir_json", doc.ToIrJson()),    // the format-neutral tree as JSON, for programs
};
foreach (var (name, value) in outputs)
    Console.WriteLine($"  chars: {name} {value.Length}");

Console.WriteLine("step 4: check each output");
foreach (var (name, value) in outputs)
    Check(value.Trim() != "", $"the {name} output is empty");
Has(outputs[0].Value, "Quarterly Report");
Check(outputs[1].Value.Contains("# Quarterly Report"), "the Markdown is missing the heading");
Check(outputs[2].Value.Contains("<table"), "the HTML has no table");
Check(JsonDocument.Parse(outputs[3].Value).RootElement.ValueKind == JsonValueKind.Object, "the IR JSON does not parse to an object");

Console.WriteLine("OK 07_all_outputs");
