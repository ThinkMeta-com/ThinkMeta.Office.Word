using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

internal class BatchDocumentConverter : IBatchDocumentConverter
{
    private readonly Application _wordApplication = new() { Visible = false };
    private bool _disposedValue;

    public void ConvertFile(string inputFilePath, string outputFilePath, DocumentFormat outputFormat)
    {
        if (outputFormat == DocumentFormat.Xps)
            ConvertFileToXps(inputFilePath, outputFilePath);
        else
            ConvertFile(inputFilePath, outputFilePath, outputFormat.GetSaveDocumentFormat());
    }

    private void ConvertFile(string inputFilePath, string outputFilePath, WdSaveFormat outputFormat)
    {
        var document = _wordApplication.Documents.Open(FileName: inputFilePath, Format: WdOpenFormat.wdOpenFormatAuto);
        document.SaveAs2(outputFilePath, outputFormat);
        document.Close();
    }

    private void ConvertFileToXps(string inputFilePath, string outputFilePath)
    {
        var document = _wordApplication.Documents.Open(FileName: inputFilePath, Format: WdOpenFormat.wdOpenFormatAuto);
        document.ExportAsFixedFormat(outputFilePath, WdExportFormat.wdExportFormatXPS, OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint);
        document.Close();
    }

    public void ReplaceStringsInFile(string filePath, Dictionary<string, string> replacements)
    {
        var document = _wordApplication.Documents.Open(FileName: filePath, Format: WdOpenFormat.wdOpenFormatAuto);
        try {
            foreach (var kvp in replacements) {
                var find = document.Content.Find;
                find.ClearFormatting();
                find.Text = kvp.Key;
                find.Replacement.ClearFormatting();
                find.Replacement.Text = kvp.Value;
                _ = find.Execute(Replace: WdReplace.wdReplaceAll, Forward: true);
            }
            document.Save();
        }
        finally {
            document.Close();
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue) {
            if (disposing) {
                _wordApplication?.Quit();
            }

            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}