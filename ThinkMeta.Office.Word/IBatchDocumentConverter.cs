namespace ThinkMeta.Office.Word;

/// <summary>
/// The batch converter interface.
/// </summary>
public interface IBatchDocumentConverter : IDisposable
{
    /// <summary>
    /// Converts a file.
    /// </summary>
    /// <param name="inputFilePath">The input file path.</param>
    /// <param name="outputFilePath">The ouput file path.</param>
    /// <param name="outputFormat">The document format of the output file.</param>
    void ConvertFile(string inputFilePath, string outputFilePath, DocumentFormat outputFormat);

    /// <summary>
    /// Replaces all occurrences of the specified strings in the file.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <param name="replacements">A dictionary of string replacements.</param>
    void ReplaceStringsInFile(string filePath, Dictionary<string, string> replacements);
}
