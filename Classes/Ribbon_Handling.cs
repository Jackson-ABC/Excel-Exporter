using DocumentFormat.OpenXml.Packaging;

namespace ExcelExporter.Classes
{
    internal class Ribbon_Handling
    {
        private string fileType = "xml";

        public void ExportRibbonXML(string tempWorkbookPath, string outputDir)
        {
            Directory.CreateDirectory(outputDir);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(tempWorkbookPath, false))
            {
                var ribbonParts = document.GetAllParts()
                .Where(p => p.Uri.OriginalString.StartsWith("/customUI/", StringComparison.OrdinalIgnoreCase) &&
                            p.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .ToList();

                if (!ribbonParts.Any())
                {
                    Console.WriteLine("No ribbon XML parts found.");
                    return;
                }

                int count = 1;
                foreach (var part in ribbonParts)
                {
                    using (var stream = part.GetStream())
                    using (var reader = new StreamReader(stream))
                    {
                        string ribbonXml = reader.ReadToEnd();
                        //string fileName = $"RibbonExport_{count}_{Path.GetFileName(part.Uri.OriginalString)}"; Have kept here, in case in the future this is actually required.
                        string fileName = $"{Path.GetFileName(part.Uri.OriginalString)}";
                        string outputPath = Path.Combine(outputDir, fileName);

                        File.WriteAllText(outputPath, ribbonXml);
                        Console.WriteLine($"Exported: {outputPath}");
                        count++;
                    }
                }
            }
        }
    }
}
