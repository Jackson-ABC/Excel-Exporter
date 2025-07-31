using ExcelExporter.Classes;
using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Excel Exporter - Command Line Interface");
        Console.WriteLine("====================================");
        Console.WriteLine("A tool for exporting data to Excel files");
        Console.WriteLine();

        if (!ArgumentsHandler.HandleArguments(args, out ParsedArguments parsedArguments, out string? messages))
        {
            Console.WriteLine(messages);
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
        Excel.Application xlApp = new Excel.Application();
        xlApp.Visible = false;
        Excel.Workbook xlWB = null;

        try
        {
            xlWB = xlApp.Workbooks.Open(inputFilePath);

            foreach (Excel.Worksheet xlWS in xlWB.Sheets)
            {
                Console.WriteLine($"Exporting worksheet '{xlWS.Name}'...");
                string displaySheetName = xlWS.Name + "_display";
                string formulaSheetName = xlWS.Name + "_formula";

                // Convert Excel range values to string arrays
                Excel.Range xlRng_used = xlWS.UsedRange;
                object[,] valueArray = xlRng_used.Value2 as object[,];
                object[,] formulaArray = xlRng_used.Formula as object[,];

                CSVHandler.WriteToCSV(valueArray, worksheetDir, displaySheetName);
                CSVHandler.WriteToCSV(formulaArray, worksheetDir, formulaSheetName);

                // Export VBA code for the worksheet
                try
                {
                    vbaHandling.ExportWorksheetVBA(xlWS, vbaDir);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not export VBA code for worksheet '{xlWS.Name}': {ex.Message}");
                }
            }

            vbaHandling.ExportModulesVBA(xlWB, vbaDir);
            vbaHandling.ExportClassesVBA(xlWB, vbaDir);
            vbaHandling.ExportFormsVBA(xlWB, vbaDir);
            vbaHandling.ExportThisWorkbookVBA(xlWB, vbaDir);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }

            #endregion

        #region Cleanup
        Console.WriteLine("Cleaning up...");
        
        if (xlWB != null)
            xlWB.Close(false);

        xlApp.Quit();
        #endregion
    }
}