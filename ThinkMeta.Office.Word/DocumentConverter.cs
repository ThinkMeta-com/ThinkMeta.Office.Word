using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

/// <summary>
/// Provides document conversion functions.
/// </summary>
public static class DocumentConverter
{
    /// <summary>
    /// Converts an RTF file to PDF.
    /// </summary>
    /// <param name="rtfFilePath">The input path of the RTF file.</param>
    /// <param name="pdfFilePath">The output path of the PDF file.</param>
    public static void ConvertRtfToPdf(string rtfFilePath, string pdfFilePath) => ConvertFile(rtfFilePath, WdOpenFormat.wdOpenFormatRTF, pdfFilePath, WdSaveFormat.wdFormatPDF);

    private static void ConvertFile(string inputFilePath, WdOpenFormat inputFormat, string outputFilePath, WdSaveFormat outputFormat)
    {
        var document = OpenWordApplication().Documents.Open(FileName: inputFilePath, Format: inputFormat);
        document.SaveAs2(outputFilePath, outputFormat);
        document.Close();
    }

    private static Application OpenWordApplication() => new() { Visible = false };
}
