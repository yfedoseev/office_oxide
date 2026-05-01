using System.Runtime.InteropServices;

namespace OfficeOxide.Internal;

/// <summary>
/// P/Invoke declarations for the office_oxide C FFI.
/// Uses <c>LibraryImport</c> source generator so every call is AOT-safe.
/// </summary>
internal static partial class NativeMethods
{
    // The library name is "office_oxide" — .NET will look for
    // liboffice_oxide.so / .dylib / office_oxide.dll on the native search path.
    private const string Lib = "office_oxide";

    internal const int OfficeOk = 0;
    internal const int OfficeErrInvalidArg = 1;
    internal const int OfficeErrIo = 2;
    internal const int OfficeErrParse = 3;
    internal const int OfficeErrExtraction = 4;
    internal const int OfficeErrInternal = 5;
    internal const int OfficeErrUnsupported = 6;

    internal const int OfficeCellEmpty = 0;
    internal const int OfficeCellString = 1;
    internal const int OfficeCellNumber = 2;
    internal const int OfficeCellBoolean = 3;

    // ── Library info / memory ──────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_oxide_version")]
    internal static partial IntPtr OfficeOxideVersion();

    [LibraryImport(Lib, EntryPoint = "office_oxide_free_string")]
    internal static partial void OfficeOxideFreeString(IntPtr ptr);

    [LibraryImport(Lib, EntryPoint = "office_oxide_free_bytes")]
    internal static partial void OfficeOxideFreeBytes(IntPtr ptr, nuint len);

    [LibraryImport(Lib, EntryPoint = "office_oxide_detect_format", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeOxideDetectFormat(string path);

    // ── Document (read-only) ───────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_document_open", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeDocumentOpen(string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_open_from_bytes", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeDocumentOpenFromBytes(
        [In] byte[] data, nuint len, string format, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_free")]
    internal static partial void OfficeDocumentFree(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_document_format")]
    internal static partial IntPtr OfficeDocumentFormat(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_document_plain_text")]
    internal static partial IntPtr OfficeDocumentPlainText(IntPtr handle, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_to_markdown")]
    internal static partial IntPtr OfficeDocumentToMarkdown(IntPtr handle, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_to_html")]
    internal static partial IntPtr OfficeDocumentToHtml(IntPtr handle, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_to_ir_json")]
    internal static partial IntPtr OfficeDocumentToIrJson(IntPtr handle, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_document_save_as", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficeDocumentSaveAs(IntPtr handle, string path, out int errorCode);

    // ── Editable ───────────────────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_editable_open", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeEditableOpen(string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_editable_free")]
    internal static partial void OfficeEditableFree(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_editable_replace_text", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial long OfficeEditableReplaceText(
        IntPtr handle, string find, string replace, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_editable_set_cell", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficeEditableSetCell(
        IntPtr handle, uint sheetIndex, string cellRef, int valueType,
        string? valueStr, double valueNum, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_editable_save", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficeEditableSave(IntPtr handle, string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_editable_save_to_bytes")]
    internal static partial IntPtr OfficeEditableSaveToBytes(IntPtr handle, out nuint outLen, out int errorCode);

    // ── One-shot ───────────────────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_extract_text", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeExtractText(string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_to_markdown", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeToMarkdown(string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_to_html", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial IntPtr OfficeToHtml(string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_create_from_markdown", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficeCreateFromMarkdown(string markdown, string format, string path, out int errorCode);

    // ── XlsxWriter ────────────────────────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_xlsx_writer_new")]
    internal static partial IntPtr OfficeXlsxWriterNew();

    [LibraryImport(Lib, EntryPoint = "office_xlsx_writer_free")]
    internal static partial void OfficeXlsxWriterFree(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_writer_add_sheet", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial uint OfficeXlsxWriterAddSheet(IntPtr handle, string name);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_sheet_set_cell", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial void OfficeXlsxSheetSetCell(
        IntPtr handle, uint sheet, uint row, uint col,
        int valueType, string? valueStr, double valueNum);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_sheet_set_cell_styled", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial void OfficeXlsxSheetSetCellStyled(
        IntPtr handle, uint sheet, uint row, uint col,
        int valueType, string? valueStr, double valueNum,
        [MarshalAs(UnmanagedType.U1)] bool bold, string? bgColor);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_sheet_merge_cells")]
    internal static partial void OfficeXlsxSheetMergeCells(
        IntPtr handle, uint sheet, uint row, uint col, uint rowSpan, uint colSpan);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_sheet_set_column_width")]
    internal static partial void OfficeXlsxSheetSetColumnWidth(
        IntPtr handle, uint sheet, uint col, double width);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_writer_save", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficeXlsxWriterSave(IntPtr handle, string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_xlsx_writer_to_bytes")]
    internal static partial IntPtr OfficeXlsxWriterToBytes(IntPtr handle, out nuint outLen, out int errorCode);

    // ── PptxWriter ────────────────────────────────────────────────────────────

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_new")]
    internal static partial IntPtr OfficePptxWriterNew();

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_free")]
    internal static partial void OfficePptxWriterFree(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_set_presentation_size")]
    internal static partial void OfficePptxWriterSetPresentationSize(IntPtr handle, ulong cx, ulong cy);

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_add_slide")]
    internal static partial uint OfficePptxWriterAddSlide(IntPtr handle);

    [LibraryImport(Lib, EntryPoint = "office_pptx_slide_set_title", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial void OfficePptxSlideSetTitle(IntPtr handle, uint slide, string title);

    [LibraryImport(Lib, EntryPoint = "office_pptx_slide_add_text", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial void OfficePptxSlideAddText(IntPtr handle, uint slide, string text);

    [LibraryImport(Lib, EntryPoint = "office_pptx_slide_add_image", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial void OfficePptxSlideAddImage(
        IntPtr handle, uint slide,
        [In] byte[] data, nuint len,
        string format,
        long x, long y, ulong cx, ulong cy);

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_save", StringMarshalling = StringMarshalling.Utf8)]
    internal static partial int OfficePptxWriterSave(IntPtr handle, string path, out int errorCode);

    [LibraryImport(Lib, EntryPoint = "office_pptx_writer_to_bytes")]
    internal static partial IntPtr OfficePptxWriterToBytes(IntPtr handle, out nuint outLen, out int errorCode);

    /// <summary>
    /// Take an FFI-allocated UTF-8 C string, copy it to a managed string,
    /// and free the original allocation.
    /// </summary>
    internal static string? PtrToStringAndFree(IntPtr ptr)
    {
        if (ptr == IntPtr.Zero) return null;
        try { return Marshal.PtrToStringUTF8(ptr); }
        finally { OfficeOxideFreeString(ptr); }
    }

    internal static string? PtrToStaticString(IntPtr ptr)
    {
        return ptr == IntPtr.Zero ? null : Marshal.PtrToStringUTF8(ptr);
    }
}
