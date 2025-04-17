namespace ThinkMeta.Office.Word;

/// <summary>
/// Provides document conversion functions.
/// </summary>
public static class DocumentConverter
{
    /// <summary>
    /// Creates a converter object for faster processing if more than one file needs to be converted.
    /// In this case, the Word application is opened only once and reused for all conversions.
    /// </summary>
    /// <returns>The converter object.</returns>
    public static IBatchDocumentConverter CreateBatchConverter() => new BatchDocumentConverter();

    /// <summary>
    /// Converts a file.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="inputFormat">The document format of the input file.</param>
    /// <param name="outputFilePath">The ouput file path.</param>
    /// <param name="outputFormat">The document format of the output file.</param>
    public static void ConvertFile(string inputFilePath, DocumentFormat inputFormat, string outputFilePath, DocumentFormat outputFormat)
    {
        using var batchConverter = CreateBatchConverter();
        batchConverter.ConvertFile(inputFilePath, inputFormat, outputFilePath, outputFormat);
    }
}
