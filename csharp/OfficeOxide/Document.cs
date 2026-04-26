using System.Runtime.InteropServices;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// A parsed Office document (DOCX / XLSX / PPTX / DOC / XLS / PPT), read-only.
/// Always dispose the instance (preferably with <c>using</c>) to release native memory.
/// </summary>
public sealed class Document : IDisposable
{
    private IntPtr _handle;
    private readonly string? _source;

    private Document(IntPtr handle, string? source)
    {
        _handle = handle;
        _source = source;
    }

    /// <summary>Library version string.</summary>
    public static string Version
    {
        get
        {
            var ptr = NativeMethods.OfficeOxideVersion();
            return NativeMethods.PtrToStaticString(ptr) ?? string.Empty;
        }
    }

    /// <summary>
    /// Detect the document format from a file path. Returns one of
    /// <c>"docx" | "xlsx" | "pptx" | "doc" | "xls" | "ppt"</c>, or <c>null</c> if unsupported.
    /// </summary>
    public static string? DetectFormat(string path)
    {
        var ptr = NativeMethods.OfficeOxideDetectFormat(path);
        return NativeMethods.PtrToStaticString(ptr);
    }

    /// <summary>Open a document from a file path.</summary>
    public static Document Open(string path)
    {
        var h = NativeMethods.OfficeDocumentOpen(path, out var err);
        if (h == IntPtr.Zero) throw new OfficeOxideException(err, "Open");
        return new Document(h, path);
    }

    /// <summary>Open a document from an in-memory buffer. <paramref name="format"/> is one of "docx"/"xlsx"/"pptx"/"doc"/"xls"/"ppt".</summary>
    public static Document FromBytes(byte[] data, string format)
    {
        ArgumentNullException.ThrowIfNull(data);
        var h = NativeMethods.OfficeDocumentOpenFromBytes(data, (nuint)data.Length, format, out var err);
        if (h == IntPtr.Zero) throw new OfficeOxideException(err, "FromBytes");
        return new Document(h, null);
    }

    /// <summary>Asynchronous wrapper around <see cref="Open"/> that offloads the blocking work to the thread pool.</summary>
    public static Task<Document> OpenAsync(string path, CancellationToken ct = default) =>
        Task.Run(() => Open(path), ct);

    /// <summary>Format name ("docx", "xlsx", …).</summary>
    public string Format
    {
        get
        {
            EnsureOpen();
            var ptr = NativeMethods.OfficeDocumentFormat(_handle);
            return NativeMethods.PtrToStaticString(ptr)
                   ?? throw new OfficeOxideException(OfficeOxideErrorCode.Internal, "Format");
        }
    }

    /// <summary>Extract plain text.</summary>
    public string PlainText() => CallString(NativeMethods.OfficeDocumentPlainText, nameof(PlainText));

    /// <summary>Convert to Markdown.</summary>
    public string ToMarkdown() => CallString(NativeMethods.OfficeDocumentToMarkdown, nameof(ToMarkdown));

    /// <summary>Convert to an HTML fragment.</summary>
    public string ToHtml() => CallString(NativeMethods.OfficeDocumentToHtml, nameof(ToHtml));

    /// <summary>Return the format-agnostic IR serialised as JSON.</summary>
    public string ToIrJson() => CallString(NativeMethods.OfficeDocumentToIrJson, nameof(ToIrJson));

    /// <summary>Save/convert the document to a file. Target format is inferred from the extension.</summary>
    public void SaveAs(string path)
    {
        EnsureOpen();
        var rc = NativeMethods.OfficeDocumentSaveAs(_handle, path, out var err);
        if (rc != 0) throw new OfficeOxideException(err, nameof(SaveAs));
    }

    private delegate IntPtr StringCall(IntPtr h, out int err);

    /// <summary>
    /// Convert a Markdown string to an Office document file.
    /// </summary>
    /// <param name="markdown">The Markdown content.</param>
    /// <param name="format">Output format: "docx", "xlsx", or "pptx".</param>
    /// <param name="path">Destination file path.</param>
    public static void CreateFromMarkdown(string markdown, string format, string path)
    {
        int rc = NativeMethods.OfficeCreateFromMarkdown(markdown, format, path, out int errCode);
        if (rc != 0) throw new OfficeOxideException(errCode, nameof(CreateFromMarkdown));
    }

    private string CallString(StringCall call, string op)
    {
        EnsureOpen();
        var ptr = call(_handle, out var err);
        if (ptr == IntPtr.Zero) throw new OfficeOxideException(err, op);
        return NativeMethods.PtrToStringAndFree(ptr) ?? string.Empty;
    }

    private void EnsureOpen()
    {
        if (_handle == IntPtr.Zero)
            throw new ObjectDisposedException(nameof(Document));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_handle != IntPtr.Zero)
        {
            NativeMethods.OfficeDocumentFree(_handle);
            _handle = IntPtr.Zero;
        }
        GC.SuppressFinalize(this);
    }

    ~Document()
    {
        if (_handle != IntPtr.Zero)
            NativeMethods.OfficeDocumentFree(_handle);
    }

    /// <inheritdoc />
    public override string ToString() =>
        _handle == IntPtr.Zero
            ? "Document(disposed)"
            : _source is null ? $"Document({Format}, from bytes)" : $"Document({Format}, {_source})";
}
