using System;
using System.Runtime.InteropServices;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// Builder for creating XLSX workbooks from scratch.
/// </summary>
public sealed class XlsxWriter : IDisposable
{
    private IntPtr _handle;

    /// <summary>Create a new empty XLSX workbook builder.</summary>
    public XlsxWriter()
    {
        _handle = NativeMethods.OfficeXlsxWriterNew();
        if (_handle == IntPtr.Zero)
            throw new OfficeOxideException(5, "XlsxWriter.new");
    }

    private void EnsureHandle()
    {
        if (_handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(XlsxWriter));
    }

    /// <summary>Add a worksheet; returns its 0-based index.</summary>
    public uint AddSheet(string name)
    {
        EnsureHandle();
        return NativeMethods.OfficeXlsxWriterAddSheet(_handle, name);
    }

    /// <summary>Set a cell value. value may be null, string, double, or bool.</summary>
    public void SetCell(uint sheet, uint row, uint col, object? value)
    {
        EnsureHandle();
        int t; string? s = null; double n = 0;
        switch (value)
        {
            case null: t = 0; break;
            case string sv: t = 1; s = sv; break;
            case double dv: t = 2; n = dv; break;
            case float fv: t = 2; n = fv; break;
            case int iv: t = 2; n = iv; break;
            case long lv: t = 2; n = lv; break;
            case bool bv: t = 2; n = bv ? 1 : 0; break;
            default: t = 1; s = value.ToString(); break;
        }
        NativeMethods.OfficeXlsxSheetSetCell(_handle, sheet, row, col, t, s, n);
    }

    /// <summary>
    /// Set a cell with styling. bgColor is a 6-char hex string ("D3D3D3") or null.
    /// </summary>
    public void SetCellStyled(uint sheet, uint row, uint col, object? value, bool bold, string? bgColor = null)
    {
        EnsureHandle();
        int t; string? s = null; double n = 0;
        switch (value)
        {
            case null: t = 0; break;
            case string sv: t = 1; s = sv; break;
            case double dv: t = 2; n = dv; break;
            case float fv: t = 2; n = fv; break;
            case int iv: t = 2; n = iv; break;
            case long lv: t = 2; n = lv; break;
            default: t = 1; s = value?.ToString(); break;
        }
        NativeMethods.OfficeXlsxSheetSetCellStyled(_handle, sheet, row, col, t, s, n, bold, bgColor);
    }

    /// <summary>Merge a rectangular range. rowSpan and colSpan must be >= 1.</summary>
    public void MergeCells(uint sheet, uint row, uint col, uint rowSpan, uint colSpan)
    {
        EnsureHandle();
        NativeMethods.OfficeXlsxSheetMergeCells(_handle, sheet, row, col, rowSpan, colSpan);
    }

    /// <summary>Set column width in Excel character units (e.g. 20.0).</summary>
    public void SetColumnWidth(uint sheet, uint col, double width)
    {
        EnsureHandle();
        NativeMethods.OfficeXlsxSheetSetColumnWidth(_handle, sheet, col, width);
    }

    /// <summary>Save the workbook to a file.</summary>
    public void Save(string path)
    {
        EnsureHandle();
        int rc = NativeMethods.OfficeXlsxWriterSave(_handle, path, out int errorCode);
        if (rc != NativeMethods.OfficeOk)
            throw new OfficeOxideException(errorCode, "XlsxWriter.Save");
    }

    /// <summary>Serialize the workbook to a byte array.</summary>
    public byte[] ToBytes()
    {
        EnsureHandle();
        IntPtr ptr = NativeMethods.OfficeXlsxWriterToBytes(_handle, out nuint len, out int errorCode);
        if (ptr == IntPtr.Zero)
            throw new OfficeOxideException(errorCode, "XlsxWriter.ToBytes");
        try
        {
            var result = new byte[(int)len];
            Marshal.Copy(ptr, result, 0, (int)len);
            return result;
        }
        finally
        {
            NativeMethods.OfficeOxideFreeBytes(ptr, len);
        }
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (_handle != IntPtr.Zero)
        {
            NativeMethods.OfficeXlsxWriterFree(_handle);
            _handle = IntPtr.Zero;
        }
    }
}
