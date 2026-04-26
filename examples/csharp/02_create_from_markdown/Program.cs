// 02_create_from_markdown — Convert Markdown to DOCX, XLSX, and PPTX.
//
// For each format: calls Document.CreateFromMarkdown, opens with Document.Open,
// prints a summary. Exit 0 on success.
//
// Run: dotnet run --project examples/csharp/02_create_from_markdown/02_create_from_markdown.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using OfficeOxide;

const string markdown = """
    # Quarterly Report

    Generated automatically from Markdown using Office Oxide.

    ## Highlights

    - Revenue grew by **32%** year-over-year
    - Customer satisfaction: 4.8 / 5.0
    - New products launched: Widget Pro, Widget Lite

    ## Financial Summary

    | Category   | Q3 2025 | Q4 2025 |
    |------------|---------|---------|
    | Revenue    | $1.2M   | $1.6M   |
    | Expenses   | $0.8M   | $0.9M   |
    | Net Profit | $0.4M   | $0.7M   |
    """;

string[] formats = ["docx", "xlsx", "pptx"];
var tmps = new Dictionary<string, string>();

try
{
    foreach (var fmt in formats)
    {
        tmps[fmt] = Path.Combine(Path.GetTempPath(), $"oo_02_create_{Environment.TickCount64}_{fmt}.{fmt}");
    }

    foreach (var fmt in formats)
    {
        var path = tmps[fmt];
        Document.CreateFromMarkdown(markdown, fmt, path);

        using var doc = Document.Open(path);
        var text = doc.PlainText();
        var md = doc.ToMarkdown();

        if (string.IsNullOrEmpty(text)) throw new Exception($"{fmt}: plain text is empty");

        Console.WriteLine($"\n=== {fmt.ToUpper()} ===");
        Console.WriteLine($"plain text length: {text.Length} chars");
        Console.WriteLine($"markdown length:   {md.Length} chars");
        var preview = text.Length > 100 ? text[..100] : text;
        Console.WriteLine($"first 100 chars: {preview!.ReplaceLineEndings(" ")}");
    }

    Console.WriteLine("\nAll formats created and verified.");
}
finally
{
    foreach (var p in tmps.Values)
    {
        if (File.Exists(p)) File.Delete(p);
    }
}
