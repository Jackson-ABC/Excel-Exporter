using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Excel Exporter - Command Line Interface");
        Console.WriteLine("====================================");
        Console.WriteLine("A tool for exporting data to Excel files");
        Console.WriteLine();

        if (!ArgumentsHandler.HandleArguments(args, out string? inputFilePath, out string? fileType, out string? outputDir, out string? outputText))
        {
            Console.WriteLine(outputText);
            return;
        }

        if (string.IsNullOrEmpty(inputFilePath) || string.IsNullOrEmpty(fileType) || string.IsNullOrEmpty(outputDir))
        {
            Console.WriteLine("No inputs specified. Use --help to see usage.");
            return;
        }

        string workbookDir = GenerateRequiredDirectories.Run(inputFilePath, fileType, outputDir);
        string worksheetDir = Path.Combine(workbookDir, "Sheets");
        string vbaDir = Path.Combine(workbookDir, "VBA");
        string ribbonXDir = Path.Combine(workbookDir, "RibbonX");
        
        #region Excel Export
        VBA_Handling vbaHandling = new VBA_Handling();

        Excel.Application excelApp = new Excel.Application();
        excelApp.Visible = false;
        Excel.Workbook workbook = excelApp.Workbooks.Open(inputFilePath);

        foreach (Excel.Worksheet worksheet in workbook.Sheets)
        {
            // Debugging for another project
            Excel.Range range = worksheet.Range["A1"];
            Console.WriteLine(range.Value2.ToString() ?? "");

            Console.WriteLine($"Exporting worksheet '{worksheet.Name}'...");
            string displaySheetName = worksheet.Name + "_display";
            string formulaSheetName = worksheet.Name + "_formula";

            // Convert Excel range values to string arrays
            Excel.Range usedRange = worksheet.UsedRange;
            object[,] valueArray = usedRange.Value2 as object[,];
            object[,] formulaArray = usedRange.Formula as object[,];

            CSVHandler.WriteToCSV(valueArray, worksheetDir, displaySheetName);
            CSVHandler.WriteToCSV(formulaArray, worksheetDir, formulaSheetName);

            // Export VBA code for the worksheet
            try
            {
                vbaHandling.ExportWorksheetVBA(worksheet, vbaDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not export VBA code for worksheet '{worksheet.Name}': {ex.Message}");
            }
        }

        vbaHandling.ExportModulesVBA(workbook, vbaDir);
        vbaHandling.ExportClassesVBA(workbook, vbaDir);
        vbaHandling.ExportFormsVBA(workbook, vbaDir);
        vbaHandling.ExportThisWorkbookVBA(workbook, vbaDir);
        #endregion

        #region Cleanup
        workbook.Close(false);
        excelApp.Quit();
        #endregion
    }
}