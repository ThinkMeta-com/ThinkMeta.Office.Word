using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

internal static class DocumentFormatExtensions
{
    public static WdOpenFormat GetOpenDocumentFormat(this DocumentFormat documentFormat)
    {
        return documentFormat switch {
            DocumentFormat.Rtf => WdOpenFormat.wdOpenFormatRTF,
            _ => throw new NotSupportedException($"Input format '{documentFormat}' is not supported")
        };
    }

    public static WdSaveFormat GetSaveDocumentFormat(this DocumentFormat documentFormat)
    {
        return documentFormat switch {
            DocumentFormat.Pdf => WdSaveFormat.wdFormatPDF,
            _ => throw new NotSupportedException($"Output format '{documentFormat}' is not supported")
        };
    }
}
