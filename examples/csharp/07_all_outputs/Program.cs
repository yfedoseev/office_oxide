// 07_all_outputs — one document, every output: plain text, Markdown, HTML and IR JSON.
//
// Run: dotnet run --project examples/csharp/07_all_outputs/07_all_outputs.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.IO.Compression;
using System.Text.Json;
using OfficeOxide;

// There is no DOCX writer in the binding, so the DOCX is a minimal OPC package
// zipped in memory: nothing touches the disk.
static byte[] BuildDocxBytes(string title = "Quarterly Report")
{
    static string Para(string t) => $"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>";
    static string Row(params string[] cells) =>
        "<w:tr>" + string.Concat(cells.Select(c => $"<w:tc>{Para(c)}</w:tc>")) + "</w:tr>";
    var body = Para(title) + Para("Revenue grew in every region.")
        + "<w:tbl>" + Row("Region", "Q1", "Q2") + Row("North", "120", "135") + Row("South", "95", "110") + "</w:tbl>";
    var parts = new Dictionary<string, string>
    {
        ["[Content_Types].xml"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
        ["_rels/.rels"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
        ["word/document.xml"] =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + $"<w:body>{body}</w:body></w:document>",
    };
    using var ms = new MemoryStream();
    using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        foreach (var (name, xml) in parts)
        {
            using var w = new StreamWriter(zip.CreateEntry(name).Open());
            w.Write(xml);
        }
    return ms.ToArray();
}

static void Check(bool condition, string message)
{
    if (!condition) throw new Exception($"check failed: {message}");
}

Console.WriteLine("step 1: build a DOCX in memory");
var bytes = BuildDocxBytes();

Console.WriteLine("step 2: open it from bytes");
using var doc = Document.FromBytes(bytes, "docx");
Console.WriteLine($"  format: {doc.Format}");

// Plain text: content only, no markup. Markdown: structure (headings, tables)
// as light markup for LLMs and diffs. HTML: the same structure as tags for a
// browser. IR JSON: the lossless intermediate representation (sections ->
// elements) that the other three are rendered from.
Console.WriteLine("step 3: render every output");
var outputs = new Dictionary<string, string>
{
    ["plain text"] = doc.PlainText(),
    ["markdown"] = doc.ToMarkdown(),
    ["html"] = doc.ToHtml(),
    ["ir json"] = doc.ToIrJson(),
};
Console.WriteLine($"  chars: {string.Join(" ", outputs.Select(o => $"{o.Key.Replace(' ', '_')}={o.Value.Length}"))}");

Console.WriteLine("step 4: check each output");
foreach (var (name, value) in outputs)
{
    Check(value.Length > 0, $"{name} is not empty");
    Check(value.Contains("Revenue grew in every region."), $"{name} contains the paragraph");
}
using (JsonDocument.Parse(outputs["ir json"])) { }   // throws if not valid JSON
Check(outputs["html"].Contains('<'), "html contains tags");
Check(!outputs["plain text"].Contains('<'), "plain text contains no markup");

Console.WriteLine("OK 07_all_outputs");
