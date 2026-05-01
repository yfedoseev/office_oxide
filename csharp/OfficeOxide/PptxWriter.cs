using System;
using System.Runtime.InteropServices;
using OfficeOxide.Internal;

namespace OfficeOxide;

/// <summary>
/// Builder for creating PPTX presentations from scratch.
/// </summary>
public sealed class PptxWriter : IDisposable
{
    private IntPtr _handle;

    /// <summary>Create a new empty PPTX presentation builder.</summary>
    public PptxWriter()
    {
        _handle = NativeMethods.OfficePptxWriterNew();
        if (_handle == IntPtr.Zero)
            throw new OfficeOxideException(5, "PptxWriter.new");
    }

    private void EnsureHandle()
    {
        if (_handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(PptxWriter));
    }

    /// <summary>
    /// Override the canvas size. 914400 EMU = 1 inch.
    /// Default: 12192000 x 6858000 (standard 16:9).
    /// </summary>
    public void SetPresentationSize(ulong cx, ulong cy)
    {
        EnsureHandle();
        NativeMethods.OfficePptxWriterSetPresentationSize(_handle, cx, cy);
    }

    /// <summary>Add a slide; returns its 0-based index.</summary>
    public uint AddSlide()
    {
        EnsureHandle();
        return NativeMethods.OfficePptxWriterAddSlide(_handle);
    }

    /// <summary>Set the title of a slide.</summary>
    public void SetSlideTitle(uint slide, string title)
    {
        EnsureHandle();
        NativeMethods.OfficePptxSlideSetTitle(_handle, slide, title);
    }

    /// <summary>Add a plain text paragraph to a slide's body area.</summary>
    public void AddSlideText(uint slide, string text)
    {
        EnsureHandle();
        NativeMethods.OfficePptxSlideAddText(_handle, slide, text);
    }

    /// <summary>
    /// Embed an image on a slide.
    /// format is "png", "jpeg"/"jpg", or "gif".
    /// x, y, cx, cy are in EMU (914400 = 1 inch).
    /// </summary>
    public void AddSlideImage(uint slide, byte[] data, string format, long x, long y, ulong cx, ulong cy)
    {
        EnsureHandle();
        NativeMethods.OfficePptxSlideAddImage(_handle, slide, data, (nuint)data.Length, format, x, y, cx, cy);
    }

    /// <summary>Save the presentation to a file.</summary>
    public void Save(string path)
    {
        EnsureHandle();
        int rc = NativeMethods.OfficePptxWriterSave(_handle, path, out int errorCode);
        if (rc != NativeMethods.OfficeOk)
            throw new OfficeOxideException(errorCode, "PptxWriter.Save");
    }

    /// <summary>Serialize the presentation to a byte array.</summary>
    public byte[] ToBytes()
    {
        EnsureHandle();
        IntPtr ptr = NativeMethods.OfficePptxWriterToBytes(_handle, out nuint len, out int errorCode);
        if (ptr == IntPtr.Zero)
            throw new OfficeOxideException(errorCode, "PptxWriter.ToBytes");
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
            NativeMethods.OfficePptxWriterFree(_handle);
            _handle = IntPtr.Zero;
        }
    }
}
