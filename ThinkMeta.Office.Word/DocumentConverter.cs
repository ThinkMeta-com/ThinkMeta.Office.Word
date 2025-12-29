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
    /// <param name="outputFilePath">The ouput file path.</param>
    /// <param name="outputFormat">The document format of the output file.</param>
    public static void ConvertFile(string inputFilePath, string outputFilePath, DocumentFormat outputFormat)
    {
        using var batchConverter = CreateBatchConverter();
        batchConverter.ConvertFile(inputFilePath, outputFilePath, outputFormat);
    }

    /// <summary>
    /// Replaces all occurrences of the specified strings in the file.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <param name="replacements">A dictionary of string replacements.</param>
    public static void ReplaceStringsInFile(string filePath, Dictionary<string, string> replacements)
    {
        using var batchConverter = CreateBatchConverter();
        batchConverter.ReplaceStringsInFile(filePath, replacements);
    }

    /// <summary>
    /// Truncates the specified file at the first occurrence of the given search string, removing all content after that
    /// point.
    /// </summary>
    /// <param name="filePath">The full path to the file to be truncated. Cannot be null or empty.</param>
    /// <param name="searchString">The string to search for within the file. All content after the first occurrence of this string will be removed.
    /// Cannot be null or empty.</param>
    public static void TruncateFile(string filePath, string searchString)
    {
        using var batchConverter = CreateBatchConverter();
        batchConverter.TruncateFile(filePath, searchString);
    }
}
