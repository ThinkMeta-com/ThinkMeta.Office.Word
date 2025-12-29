# ThinkMeta.Office.Word

[![NuGet Package](https://img.shields.io/nuget/v/ThinkMeta.Office.Word)](https://www.nuget.org/packages/ThinkMeta.Office.Word)

Provides functions that use Microsoft Office Word Interop

# Requirements

Microsoft Word must be installed on the system.

# Supported Formats

- PDF
- RTF
- XPS
- DOCX

# Usage

```cs
// Microsoft Word is opened for each call
DocumentConverter.ConvertFile("input.rtf", "output.pdf", DocumentFormat.Pdf);
DocumentConverter.ConvertFile("input.docx", "output.xps", DocumentFormat.Xps);

// Replace strings in a document (in-place)
DocumentConverter.ReplaceStringsInFile("input.docx", new Dictionary<string, string> {
    {"oldText", "newText"},
    {"foo", "bar"}
});
```

For batch processing, use the `IBatchDocumentConverter` interface:

```cs
// Microsoft Word is opened only once
using var batchConverter = DocumentConverter.CreateBatchConverter();

batchConverter.ConvertFile("input1.rtf", "output1.pdf", DocumentFormat.Pdf);
batchConverter.ConvertFile("input2.docx", "output2.xps", DocumentFormat.Xps);

// Replace strings in a document (in-place)
batchConverter.ReplaceStringsInFile("input.docx", new Dictionary<string, string> {
    {"oldText", "newText"},
    {"foo", "bar"}
});