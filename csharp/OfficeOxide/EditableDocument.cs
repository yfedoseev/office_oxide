using System.Runtime.InteropServices;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// A DOCX, XLSX, or PPTX document opened for editing.
/// Always dispose to release native memory.
/// </summary>
public sealed class EditableDocument : IDisposable
{
    private IntPtr _handle;

    private EditableDocument(IntPtr handle) { _handle = handle; }

    /// <summary>Open a document for editing.</summary>
    public static EditableDocument Open(string path)
    {
        var h = NativeMethods.OfficeEditableOpen(path, out var err);
        if (h == IntPtr.Zero) throw new OfficeOxideException(err, "Open");
        return new EditableDocument(h);
    }

    /// <summary>Replace every occurrence of <paramref name="find"/> with <paramref name="replace"/>. Returns the replacement count.</summary>
    public long ReplaceText(string find, string replace)
    {
        EnsureOpen();
        var n = NativeMethods.OfficeEditableReplaceText(_handle, find, replace, out var err);
        if (n < 0) throw new OfficeOxideException(err, nameof(ReplaceText));
        return n;
    }

    /// <summary>Set a cell to an empty value.</summary>
    public void SetCellEmpty(uint sheetIndex, string cellRef)
    {
        SetCellInternal(sheetIndex, cellRef, NativeMethods.OfficeCellEmpty, null, 0.0);
    }

    /// <summary>Set a cell to a string value.</summary>
    public void SetCell(uint sheetIndex, string cellRef, string value)
    {
        SetCellInternal(sheetIndex, cellRef, NativeMethods.OfficeCellString, value, 0.0);
    }

    /// <summary>Set a cell to a numeric value.</summary>
    public void SetCell(uint sheetIndex, string cellRef, double value)
    {
        SetCellInternal(sheetIndex, cellRef, NativeMethods.OfficeCellNumber, null, value);
    }

    /// <summary>Set a cell to a boolean value.</summary>
    public void SetCell(uint sheetIndex, string cellRef, bool value)
    {
        SetCellInternal(sheetIndex, cellRef, NativeMethods.OfficeCellBoolean, null, value ? 1.0 : 0.0);
    }

    private void SetCellInternal(uint sheetIndex, string cellRef, int valueType, string? valueStr, double valueNum)
    {
        EnsureOpen();
        var rc = NativeMethods.OfficeEditableSetCell(
            _handle, sheetIndex, cellRef, valueType, valueStr ?? string.Empty, valueNum, out var err);
        if (rc != 0) throw new OfficeOxideException(err, nameof(SetCell));
    }

    /// <summary>Persist the edited document to disk.</summary>
    public void Save(string path)
    {
        EnsureOpen();
        var rc = NativeMethods.OfficeEditableSave(_handle, path, out var err);
        if (rc != 0) throw new OfficeOxideException(err, nameof(Save));
    }

    /// <summary>Persist the edited document to a managed byte array.</summary>
    public byte[] SaveToBytes()
    {
        EnsureOpen();
        var ptr = NativeMethods.OfficeEditableSaveToBytes(_handle, out var len, out var err);
        if (ptr == IntPtr.Zero) throw new OfficeOxideException(err, nameof(SaveToBytes));
        try
        {
            var bytes = new byte[(int)len];
            Marshal.Copy(ptr, bytes, 0, bytes.Length);
            return bytes;
        }
        finally { NativeMethods.OfficeOxideFreeBytes(ptr, len); }
    }

    private void EnsureOpen()
    {
        if (_handle == IntPtr.Zero)
            throw new ObjectDisposedException(nameof(EditableDocument));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_handle != IntPtr.Zero)
        {
            NativeMethods.OfficeEditableFree(_handle);
            _handle = IntPtr.Zero;
        }
        GC.SuppressFinalize(this);
    }

    ~EditableDocument()
    {
        if (_handle != IntPtr.Zero)
            NativeMethods.OfficeEditableFree(_handle);
    }
}

/// <summary>One-shot convenience helpers that open + extract in a single call.</summary>
public static class OfficeOxide
{
    /// <summary>Open a file and return its plain text.</summary>
    public static string ExtractText(string path)
    {
        var ptr = NativeMethods.OfficeExtractText(path, out var err);
        if (ptr == IntPtr.Zero) throw new OfficeOxideException(err, nameof(ExtractText));
        return NativeMethods.PtrToStringAndFree(ptr) ?? string.Empty;
    }

    /// <summary>Open a file and return it rendered as Markdown.</summary>
    public static string ToMarkdown(string path)
    {
        var ptr = NativeMethods.OfficeToMarkdown(path, out var err);
        if (ptr == IntPtr.Zero) throw new OfficeOxideException(err, nameof(ToMarkdown));
        return NativeMethods.PtrToStringAndFree(ptr) ?? string.Empty;
    }

    /// <summary>Open a file and return it rendered as HTML.</summary>
    public static string ToHtml(string path)
    {
        var ptr = NativeMethods.OfficeToHtml(path, out var err);
        if (ptr == IntPtr.Zero) throw new OfficeOxideException(err, nameof(ToHtml));
        return NativeMethods.PtrToStringAndFree(ptr) ?? string.Empty;
    }
}
