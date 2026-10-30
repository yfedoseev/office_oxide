// 08_limits_strict_damage — strict mode rejects a truncated archive; a lenient open succeeds with a damage report or fails with an error.
//
// Run: dotnet run --project examples/csharp/08_limits_strict_damage/08_limits_strict_damage.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.IO.Compression;
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
var whole = BuildDocxBytes();

Console.WriteLine("step 2: keep only the first half of its bytes");
var damaged = whole[..(whole.Length / 2)];
Console.WriteLine($"  chars: {whole.Length} -> {damaged.Length}");

// Strict turns every recoverable problem into an error.
Console.WriteLine("step 3: open it in strict mode");
using (var strict = new ParseOptions(strict: true))
{
    string? error = null;
    try { using var _ = Document.FromBytes(damaged, "docx", strict); }
    catch (OfficeOxideException e) { error = e.Message; }
    Console.WriteLine($"  error: {error ?? "none"}");
    Check(error is not null, "strict mode rejects a truncated archive");
}

// Lenient may salvage what it can and report the loss, or give up with an
// error. Both are fine; hanging or crashing is not.
Console.WriteLine("step 4: open it leniently");
using (var lenient = new ParseOptions())
{
    try
    {
        using var doc = Document.FromBytes(damaged, "docx", lenient);
        var damage = doc.Damage();
        Console.WriteLine($"  damage: {damage.Length}");
        foreach (var d in damage) Console.WriteLine($"    {d.Kind} {d.Part}: {d.Detail}");
        Check(damage.Any(d => d.IsLoss), "a lenient open of damaged bytes reports the loss");
    }
    catch (OfficeOxideException e)
    {
        Console.WriteLine($"  error: {e.Message}");
    }
}

Console.WriteLine("OK 08_limits_strict_damage");
