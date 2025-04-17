using Microsoft.Office.Interop.Word;

namespace ThinkMeta.Office.Word;

internal class BatchDocumentConverter : IBatchDocumentConverter
{
    private readonly Application _wordApplication = new() { Visible = false };
    private bool _disposedValue;

    public void ConvertFile(string inputFilePath, DocumentFormat inputFormat, string outputFilePath, DocumentFormat outputFormat) => ConvertFile(inputFilePath, inputFormat.GetOpenDocumentFormat(), outputFilePath, outputFormat.GetSaveDocumentFormat());

    private void ConvertFile(string inputFilePath, WdOpenFormat inputFormat, string outputFilePath, WdSaveFormat outputFormat)
    {
        var document = _wordApplication.Documents.Open(FileName: inputFilePath, Format: inputFormat);
        document.SaveAs2(outputFilePath, outputFormat);
        document.Close();
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