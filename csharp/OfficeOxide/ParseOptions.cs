using System.Text.Json;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// Options for one document: limits, which text is kept, and whether an
/// incomplete result is an error. Nothing global changes; a document copies
/// the options it is opened with. Dispose when done.
/// </summary>
public sealed class ParseOptions : IDisposable
{
    /// <summary>The value that switches a limit off.</summary>
    public const ulong NoLimit = ulong.MaxValue;

    private IntPtr _handle;

    /// <summary>Options holding the process-wide settings in force now, or the strict ones.</summary>
    /// <param name="strict">Tighter limits, and a document that cannot be read completely is an error.</param>
    public ParseOptions(bool strict = false)
    {
        _handle = strict ? NativeMethods.OfficeOptionsStrict() : NativeMethods.OfficeOptionsNew();
        if (_handle == IntPtr.Zero) throw new OfficeOxideException(OfficeOxideErrorCode.Internal, nameof(ParseOptions));
    }

    /// <summary>
    /// Set an option by name (see <see cref="Names"/>): a limit such as
    /// <c>text_chars</c> or <c>part_bytes</c> with a count (<see cref="NoLimit"/> for none),
    /// <c>include_hidden_text</c> / <c>include_deleted_text</c> with 0 or 1,
    /// <c>on_incomplete</c> with 0 (report) or 1 (error).
    /// </summary>
    public ParseOptions Set(string name, ulong value)
    {
        var rc = NativeMethods.OfficeOptionsSet(Handle, name, value);
        if (rc != 0)
            throw new ArgumentException($"unknown option '{name}'; the options are: {string.Join(", ", Names())}", nameof(name));
        return this;
    }

    /// <summary>Set an on/off option such as <c>include_hidden_text</c>.</summary>
    public ParseOptions Set(string name, bool on) => Set(name, on ? 1UL : 0UL);

    /// <summary>Whether an incomplete document is an error (<see cref="OfficeOxideException.IsIncomplete"/>) instead of a report.</summary>
    public bool OnIncomplete
    {
        set => Set("on_incomplete", value);
    }

    /// <summary>The names <see cref="Set(string, ulong)"/> accepts.</summary>
    public static string[] Names()
    {
        var ptr = NativeMethods.OfficeOptionsNames();
        return (NativeMethods.PtrToStringAndFree(ptr) ?? string.Empty).Split(',');
    }

    internal IntPtr Handle
    {
        get
        {
            if (_handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(ParseOptions));
            return _handle;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_handle != IntPtr.Zero)
        {
            NativeMethods.OfficeOptionsFree(_handle);
            _handle = IntPtr.Zero;
        }
        GC.SuppressFinalize(this);
    }

    ~ParseOptions()
    {
        if (_handle != IntPtr.Zero) NativeMethods.OfficeOptionsFree(_handle);
    }
}

/// <summary>One thing wrong with a document.</summary>
public sealed class DocumentDamage
{
    /// <summary><c>unreadable_part</c>, <c>truncated</c> or <c>recovered</c>.</summary>
    public string Kind { get; init; } = "";

    /// <summary>The part concerned, or null when the reader does not know it.</summary>
    public string? Part { get; init; }

    /// <summary>What happened, as the reader reports it.</summary>
    public string Detail { get; init; } = "";

    /// <summary>Whether content is missing (every kind but <c>recovered</c>).</summary>
    public bool IsLoss => Kind != "recovered";

    internal static DocumentDamage[] Parse(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var found = new List<DocumentDamage>();
        foreach (var item in doc.RootElement.EnumerateArray())
        {
            found.Add(new DocumentDamage
            {
                Kind = item.GetProperty("kind").GetString() ?? "",
                Part = item.TryGetProperty("part", out var part) && part.ValueKind == JsonValueKind.String
                    ? part.GetString()
                    : null,
                Detail = item.GetProperty("detail").GetString() ?? "",
            });
        }
        return found.ToArray();
    }
}
