public class GenerateRequiredDirectories
{
    public GenerateRequiredDirectories() { }

    /// <summary>
    /// Generates the required directories for the workbook
    /// </summary>
    /// <param name="inputFilePath">The file path of the workbook</param>
    /// <param name="fileType">The file type of the workbook (e.g. .xlsx, .xltx, .xlsm, .xltm, .xlam)</param>
    /// <param name="outputDir">The directory to create the workbook directory in</param>
    public static string Run(string inputFilePath, string fileType, string outputDir)
    {
        // Create "Workbook" directory
        string workbookName = Path.GetFileNameWithoutExtension(inputFilePath) + "_internals";
        string workbookDir = Path.Combine(outputDir, workbookName);
        Directory.CreateDirectory(workbookDir);

        // Create "Sheets" directory
        Directory.CreateDirectory(Path.Combine(workbookDir, "Sheets"));

        if(fileType == ".xl?m")
        {
            // Create "VBA" directories
            Directory.CreateDirectory(Path.Combine(workbookDir, "VBA"));
            Directory.CreateDirectory(Path.Combine(workbookDir, "VBA", "Microsoft Excel Objects"));
            Directory.CreateDirectory(Path.Combine(workbookDir, "VBA", "Modules"));
            Directory.CreateDirectory(Path.Combine(workbookDir, "VBA", "Classes"));
            Directory.CreateDirectory(Path.Combine(workbookDir, "VBA", "Forms"));

            // Create "RibbonX" directory (Custom Ribbons)
            Directory.CreateDirectory(Path.Combine(workbookDir, "RibbonX"));
            Directory.CreateDirectory(Path.Combine(workbookDir, "RibbonX", "Icons"));
        }

        return workbookDir;
    }
}