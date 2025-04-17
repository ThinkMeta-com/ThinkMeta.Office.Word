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
    /// <param name="inputFormat">The document format of the input file.</param>
    /// <param name="outputFilePath">The ouput file path.</param>
    /// <param name="outputFormat">The document format of the output file.</param>
    void ConvertFile(string inputFilePath, DocumentFormat inputFormat, string outputFilePath, DocumentFormat outputFormat);
}
