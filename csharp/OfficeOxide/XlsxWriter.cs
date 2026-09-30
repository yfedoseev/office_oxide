using System;
using System.Runtime.InteropServices;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// A formula cell value for <see cref="XlsxWriter.SetCell"/>, e.g.
/// <c>new Formula("SUM(A1:A3)")</c>. A leading '=' is accepted.
/// </summary>
public sealed record Formula(string Text);

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

    /// <summary>
    /// Set a cell value. value may be null, a string, a bool, any numeric type,
    /// or a <see cref="Formula"/>. A string is always text, even one starting
    /// with '='.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A <c>long</c>/<c>ulong</c> beyond ±2^53, which Excel (storing doubles) would change.
    /// </exception>
    /// <exception cref="InvalidOperationException">The sheet or cell is out of range; nothing was written.</exception>
    public void SetCell(uint sheet, uint row, uint col, object? value)
    {
        EnsureHandle();
        var (t, s, n) = Encode(value);
        // A non-zero status means the value was NOT written — an out-of-grid
        // row/column or a bad sheet index. Ignoring it silently discarded the
        // caller's data while reporting success.
        int rc = NativeMethods.OfficeXlsxSheetSetCell(_handle, sheet, row, col, t, s, n);
        if (rc != 0)
        {
            throw new InvalidOperationException(
                $"SetCell({sheet},{row},{col}) wrote nothing (status {rc})");
        }
    }

    /// <summary>
    /// Set a cell with styling. bgColor is a 6-char hex string ("D3D3D3") or null.
    /// </summary>
    public void SetCellStyled(uint sheet, uint row, uint col, object? value, bool bold, string? bgColor = null)
    {
        EnsureHandle();
        var (t, s, n) = Encode(value);
        int rc = NativeMethods.OfficeXlsxSheetSetCellStyled(
            _handle, sheet, row, col, t, s, n, bold, bgColor);
        if (rc != 0)
        {
            throw new InvalidOperationException(
                $"SetCellStyled({sheet},{row},{col}) wrote nothing (status {rc})");
        }
    }

    /// <summary>Set a formula cell, e.g. <c>SetFormula(0, 3, 1, "SUM(B1:B3)")</c>. A leading '=' is accepted.</summary>
    /// <exception cref="InvalidOperationException">The sheet or cell is out of range; nothing was written.</exception>
    public void SetFormula(uint sheet, uint row, uint col, string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        SetCell(sheet, row, col, new Formula(formula));
    }

    /// <summary>The largest magnitude up to which every integer is exactly a double.</summary>
    private const long MaxExactInt = 1L << 53;

    private static (int Type, string? Str, double Num) Encode(object? value)
    {
        switch (value)
        {
            case null: return (NativeMethods.OfficeCellEmpty, null, 0);
            case Formula f: return (NativeMethods.OfficeCellFormula, f.Text, 0);
            case string sv: return (NativeMethods.OfficeCellString, sv, 0);
            case bool bv: return (NativeMethods.OfficeCellBoolean, null, bv ? 1 : 0);
            case double dv: return (NativeMethods.OfficeCellNumber, null, dv);
            case float fv: return (NativeMethods.OfficeCellNumber, null, fv);
            case int iv: return (NativeMethods.OfficeCellNumber, null, iv);
            case short hv: return (NativeMethods.OfficeCellNumber, null, hv);
            case ushort uhv: return (NativeMethods.OfficeCellNumber, null, uhv);
            case uint uiv: return (NativeMethods.OfficeCellNumber, null, uiv);
            case byte byv: return (NativeMethods.OfficeCellNumber, null, byv);
            case sbyte sbv: return (NativeMethods.OfficeCellNumber, null, sbv);
            case decimal mv: return (NativeMethods.OfficeCellNumber, null, (double)mv);
            // A long/ulong beyond 2^53 cannot be held exactly by the double
            // Excel stores; converting silently changed IDs and account numbers.
            case long lv:
                if (lv > MaxExactInt || lv < -MaxExactInt)
                    throw new ArgumentOutOfRangeException(nameof(value), lv,
                        "cannot be stored exactly as an Excel number (a double is exact only up to 2^53); write it as a string");
                return (NativeMethods.OfficeCellNumber, null, lv);
            case ulong ulv:
                if (ulv > (ulong)MaxExactInt)
                    throw new ArgumentOutOfRangeException(nameof(value), ulv,
                        "cannot be stored exactly as an Excel number (a double is exact only up to 2^53); write it as a string");
                return (NativeMethods.OfficeCellNumber, null, ulv);
            // Anything else is rendered with the invariant culture: the
            // default ToString() made output depend on the host locale, so a
            // decimal became "1,5" under de-DE.
            default:
                return (NativeMethods.OfficeCellString,
                    System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture), 0);
        }
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
