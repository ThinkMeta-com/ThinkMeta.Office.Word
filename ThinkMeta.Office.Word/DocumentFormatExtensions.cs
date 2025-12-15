using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

internal static class DocumentFormatExtensions
{
    public static WdSaveFormat GetSaveDocumentFormat(this DocumentFormat documentFormat)
    {
        return documentFormat switch {
            DocumentFormat.Pdf => WdSaveFormat.wdFormatPDF,
            DocumentFormat.Xps => WdSaveFormat.wdFormatXPS,
            DocumentFormat.Docx => WdSaveFormat.wdFormatDocument,
            _ => throw new NotSupportedException($"Output format '{documentFormat}' is not supported")
        };
    }
}
