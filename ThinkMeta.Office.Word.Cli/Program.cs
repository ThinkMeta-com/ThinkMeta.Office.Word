namespace ThinkMeta.Office.Word.Cli;

internal static class Program
{
    private static void Main(string[] args)
    {
        if (args.Length == 0)
            return;

        switch (args[0]) {
            case "convert":
                ConvertFile(args[1], args[2]);
                break;

            default:
                Console.WriteLine($"Unknown command '{args[0]}.'");
                break;
        }
    }

    private static void ConvertFile(string input, string output)
    {
        var inputExtension = Path.GetExtension(input).ToLowerInvariant();
        var outputExtension = Path.GetExtension(output).ToLowerInvariant();

        var inputFormat = inputExtension switch {
            ".rtf" => DocumentFormat.Rtf,
            _ => throw new NotSupportedException($"Input file '{input}' not supported.")
        };

        var outputFormat = outputExtension switch {
            ".pdf" => DocumentFormat.Pdf,
            _ => throw new NotSupportedException($"Output file '{output}' not supported.")
        };

        DocumentConverter.ConvertFile(input, inputFormat, output, outputFormat);
    }
}
