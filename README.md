# ThinkMeta.Office.Word

Provides that use Microsoft Office Word Interop

# Requirements

Microsoft Word must be installed on the system.

# Usage

Currently, only RTF-to-PDF file conversion tool is available:

```cs
DocumentConverter.ConvertRtfToPdf("input.rtf", "output.pdf");
```