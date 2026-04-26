// 03_edit — Edit a document by replacing placeholder text.
//
// Creates a DOCX from Markdown with {{NAME}} and {{DATE}} placeholders,
// opens with EditableDocument.Open, replaces both, saves to a second temp
// file, then reads back and verifies. Exit 0 on success.
//
// Run: dotnet run --project examples/csharp/03_edit/03_edit.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using OfficeOxide;

const string markdown = """
    # Invoice

    Dear {{NAME}},

    Please find attached your invoice for services rendered on {{DATE}}.

    ## Summary

    - Service: Office Oxide Pro License
    - Amount: $499.00
    - Due date: 30 days from {{DATE}}

    Thank you for your business!
    """;

var ts = Environment.TickCount64;
var templatePath = Path.Combine(Path.GetTempPath(), $"oo_03_template_{ts}.docx");
var outputPath = Path.Combine(Path.GetTempPath(), $"oo_03_output_{ts}.docx");

try
{
    // Create template
    Document.CreateFromMarkdown(markdown, "docx", templatePath);

    // Edit
    int n1, n2;
    using (var ed = EditableDocument.Open(templatePath))
    {
        n1 = ed.ReplaceText("{{NAME}}", "Alice Smith");
        n2 = ed.ReplaceText("{{DATE}}", "2026-04-26");
        ed.Save(outputPath);
    }

    Console.WriteLine($"Replacements: {{NAME}} x{n1}, {{DATE}} x{n2}");

    // Verify
    using var doc = Document.Open(outputPath);
    var text = doc.PlainText();

    if (!text.Contains("Alice Smith")) throw new Exception("name replacement failed: 'Alice Smith' not found");
    if (!text.Contains("2026-04-26")) throw new Exception("date replacement failed: '2026-04-26' not found");
    if (text.Contains("{{NAME}}")) throw new Exception("placeholder {{NAME}} still present");
    if (text.Contains("{{DATE}}")) throw new Exception("placeholder {{DATE}} still present");

    Console.WriteLine("Edit verified successfully.");
    Console.WriteLine("--- final text ---");
    Console.WriteLine(text);
}
finally
{
    if (File.Exists(templatePath)) File.Delete(templatePath);
    if (File.Exists(outputPath)) File.Delete(outputPath);
}
