namespace ExcelExporter.Classes
{
    public class GenerateRequiredDirectories
    {
        public GenerateRequiredDirectories() { }

        /// <summary>
        /// Generates the required directories for the workbook
        /// </summary>
        /// <param name="inputFilePath">The file path of the workbook</param>
        /// <param name="fileType">The file type of the workbook (e.g. .xlsx, .xltx, .xlsm, .xltm, .xlam)</param>
        /// <param name="outputDir">The directory to create the workbook directory in</param>
        public static void Run(string inputFilePath, string fileType, string savePath)
        {
            // Create "Workbook" directory
            Directory.CreateDirectory(savePath);

            // Create "Sheets" directory
            Directory.CreateDirectory(Path.Combine(savePath, "Sheets"));

            if (fileType == ".xl?m")
            {
                // Create "VBA" directories
                Directory.CreateDirectory(Path.Combine(savePath, "VBA"));
                Directory.CreateDirectory(Path.Combine(savePath, "VBA", "Microsoft Excel Objects"));
                Directory.CreateDirectory(Path.Combine(savePath, "VBA", "Modules"));
                Directory.CreateDirectory(Path.Combine(savePath, "VBA", "Classes"));
                Directory.CreateDirectory(Path.Combine(savePath, "VBA", "Forms"));

                // Create "RibbonX" directory (Custom Ribbons)
                Directory.CreateDirectory(Path.Combine(savePath, "RibbonX"));
                Directory.CreateDirectory(Path.Combine(savePath, "RibbonX", "Icons"));
            }
        }
    }
}