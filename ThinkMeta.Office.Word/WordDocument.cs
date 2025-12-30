using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

/// <summary>
/// Represents a Microsoft Word document, providing access to its content and structure.
/// </summary>
public class WordDocument
{
    internal Document Document { get; }

    internal WordDocument(Document document) => Document = document;

    /// <summary>
    /// Saves the document to the specified file path in the given format and returns the document for fluent chaining.
    /// </summary>
    /// <param name="outputFilePath">The file path to save the document to.</param>
    /// <param name="outputFormat">The format in which to save the document.</param>
    /// <returns>The same <see cref="WordDocument"/> instance for fluent chaining.</returns>
    public WordDocument SaveAs(string outputFilePath, DocumentFormat outputFormat)
    {
        var wdFormat = outputFormat.GetSaveDocumentFormat();
        Document.SaveAs2(outputFilePath, wdFormat);
        return this;
    }

    /// <summary>
    /// Saves the current state of the specified Word document to its underlying storage.
    /// </summary>
    /// <returns>The same <see cref="WordDocument"/> instance that was saved.</returns>
    public WordDocument Save()
    {
        Document.Save();
        return this;
    }

    /// <summary>
    /// Closes the document and returns the document for fluent chaining.
    /// </summary>
    /// <returns>The same <see cref="WordDocument"/> instance that was closed.</returns>
    public WordDocument Close()
    {
        Document.Close();
        return this;
    }

    /// <summary>
    /// Exports the document as a fixed format (such as PDF or XPS) to the specified file path and returns the document for fluent chaining.
    /// </summary>
    /// <param name="outputFilePath">The file path to export the document to.</param>
    /// <param name="exportFormat">The fixed format to export as.</param>
    /// <returns>The same <see cref="WordDocument"/> instance for fluent chaining.</returns>
    public WordDocument ExportAsFixedFormat(string outputFilePath, DocumentFormat exportFormat)
    {
        var wdExportFormat = exportFormat.GetExportDocumentFormat();
        Document.ExportAsFixedFormat(outputFilePath, wdExportFormat, OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint);
        return this;
    }

    /// <summary>
    /// Replaces all occurrences of specified strings in the document with their corresponding replacement values and returns the document for fluent chaining.
    /// </summary>
    /// <param name="replacements">A dictionary of string replacements where the key is the string to find and the value is the replacement.</param>
    /// <returns>The same <see cref="WordDocument"/> instance for fluent chaining.</returns>
    public WordDocument ReplaceStrings(Dictionary<string, string> replacements)
    {
        var range = Document.Content;
        var find = range.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.MatchWildcards = false;

        foreach (var kvp in replacements) {
            find.Text = kvp.Key;
            find.Replacement.Text = kvp.Value;
            _ = find.Execute(Replace: WdReplace.wdReplaceAll, Forward: true);
        }

        return this;
    }

    /// <summary>
    /// Truncates the document content from the first occurrence of the specified search string to the end of the document and returns the document for fluent chaining.
    /// </summary>
    /// <param name="searchString">The string to search for as the truncation point.</param>
    /// <returns>The same <see cref="WordDocument"/> instance for fluent chaining.</returns>
    public WordDocument Truncate(string searchString)
    {
        var searchRange = Document.Content;
        if (searchRange.Find.Execute(FindText: searchString, MatchCase: true, MatchWholeWord: false, Forward: true))
            _ = Document.Range(searchRange.Start, Document.Content.End).Delete();

        return this;
    }
}
