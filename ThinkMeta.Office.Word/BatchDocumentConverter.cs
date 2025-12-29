using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

internal class BatchDocumentConverter : IBatchDocumentConverter
{
    private readonly Application _wordApplication = new() { Visible = false };
    private bool _disposedValue;

    public WordDocument Open(string filePath) => new(_wordApplication.Documents.Open(FileName: filePath, Format: WdOpenFormat.wdOpenFormatAuto));

    public void ConvertFile(string inputFilePath, string outputFilePath, DocumentFormat outputFormat)
    {
        if (outputFormat == DocumentFormat.Xps) {
            _ = Open(inputFilePath)
                .ExportAsFixedFormat(outputFilePath, DocumentFormat.Xps)
                .Close();
        }
        else {
            _ = Open(inputFilePath)
                .SaveAs(outputFilePath, outputFormat)
                .Close();
        }
    }

    public void ReplaceStringsInFile(string filePath, Dictionary<string, string> replacements)
    {
        var document = Open(filePath);
        try {
            _ = document
                .ReplaceStrings(replacements)
                .Save();
        }
        finally {
            _ = document.Close();
        }
    }

    public void TruncateFile(string filePath, string searchString)
    {
        var document = Open(filePath);
        try {
            _ = document
                .Truncate(searchString)
                .Save();
        }
        finally {
            _ = document.Close();
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