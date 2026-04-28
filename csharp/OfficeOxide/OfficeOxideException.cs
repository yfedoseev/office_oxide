namespace OfficeOxide;

/// <summary>Exception thrown for any office_oxide FFI failure.</summary>
public sealed class OfficeOxideException : Exception
{
    /// <summary>The underlying FFI error code (see <see cref="OfficeOxideErrorCode"/>).</summary>
    public int Code { get; }

    /// <summary>The operation that produced the error (e.g. "Open", "ToMarkdown").</summary>
    public string Operation { get; }

    internal OfficeOxideException(int code, string operation)
        : base($"office_oxide: {operation}: {Describe(code)}")
    {
        Code = code;
        Operation = operation;
    }

    private static string Describe(int code) => code switch
    {
        OfficeOxideErrorCode.Ok => "ok",
        OfficeOxideErrorCode.InvalidArg => "invalid argument",
        OfficeOxideErrorCode.Io => "io error",
        OfficeOxideErrorCode.Parse => "parse error",
        OfficeOxideErrorCode.Extraction => "extraction failed",
        OfficeOxideErrorCode.Internal => "internal error",
        OfficeOxideErrorCode.Unsupported => "unsupported format",
        _ => $"code={code}",
    };
}

/// <summary>Numeric error codes returned by the FFI layer.</summary>
public static class OfficeOxideErrorCode
{
    public const int Ok = 0;
    public const int InvalidArg = 1;
    public const int Io = 2;
    public const int Parse = 3;
    public const int Extraction = 4;
    public const int Internal = 5;
    public const int Unsupported = 6;
}
