# ThinkMeta.Office.Word

[![NuGet Package](https://img.shields.io/nuget/v/ThinkMeta.Office.Word)](https://www.nuget.org/packages/ThinkMeta.Office.Word)

Provides that use Microsoft Office Word Interop

# Requirements

Microsoft Word must be installed on the system.

# Usage

Currently, only RTF-to-PDF file conversion tool is available:

```cs
// Microsoft Word is opened for each call
DocumentConverter.ConvertFile("input.rtf", DocumentFormat.Rtf, "output.pdf", DocumentFormat.Pdf);
```

For batch processing, use the `IBatchDocumentConverter` interface:

```cs
// Microsoft Word is opened only once
using var batchConverter = DocumentConverter.CreateBatchConverter();

batchConverter.ConvertFile("input1.rtf", DocumentFormat.Rtf, "output1.pdf", DocumentFormat.Pdf);
batchConverter.ConvertFile("input2.rtf", DocumentFormat.Rtf, "output2.pdf", DocumentFormat.Pdf);
```
