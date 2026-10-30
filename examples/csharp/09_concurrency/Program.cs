// 09_concurrency — open and read 32 documents from 8 workers at once; every result must match.
//
// Run: dotnet run --project examples/csharp/09_concurrency/09_concurrency.csproj
//      (set LD_LIBRARY_PATH to the directory containing liboffice_oxide.so)
using System.Collections.Concurrent;
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

const int Workers = 8, Documents = 32;

// A plain TextWriter log keeps the example free of logging packages; it is
// synchronized because eight workers write to it.
var log = TextWriter.Synchronized(Console.Error);

Console.WriteLine("step 1: build the documents in memory");
var inputs = Enumerable.Range(0, Documents).Select(i => BuildDocxBytes($"Quarterly Report {i}")).ToArray();

Console.WriteLine("step 2: read them from 8 workers");
var texts = new ConcurrentDictionary<int, string>();
await Parallel.ForEachAsync(Enumerable.Range(0, Documents), new ParallelOptions { MaxDegreeOfParallelism = Workers },
    (i, _) =>
    {
        log.WriteLine($"[debug] worker thread {Environment.CurrentManagedThreadId} opens document {i}");
        using var doc = Document.FromBytes(inputs[i], "docx");
        texts[i] = doc.PlainText();
        return ValueTask.CompletedTask;
    }).WaitAsync(TimeSpan.FromSeconds(60));   // a deadlock fails the example instead of hanging it
Console.WriteLine($"  workers: {Workers}");
Console.WriteLine($"  documents: {texts.Count}");

Console.WriteLine("step 3: check every result");
Check(texts.Count == Documents, "every document was read");
for (var i = 0; i < Documents; i++)
    Check(texts[i].Contains($"Quarterly Report {i}"), $"document {i} reads back its own text");

Console.WriteLine("OK 09_concurrency");
