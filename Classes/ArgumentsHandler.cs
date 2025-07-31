namespace ExcelExporter.Classes
{
    public class ArgumentsHandler
    {
        private static Dictionary<string, Command> commands = new Dictionary<string, Command>();
        static ArgumentsHandler()
        {
            commands.Add("help",
                new Command("help",
                    "--help; -h",
                    "Display this help message",
                    HelpHandler
                )
            );
            commands.Add("version",
                new Command("version",
                    "--version; -v",
                    "Display the version of the program",
                    VersionHandler
                )
            );
            commands.Add("outputPath",
                new Command("outputPath",
                    "--outputPath; -op",
                    "The path to extract the workbook to\nUsage: ExcelExporter.exe <input_file> --outputPath <output_path>\nExample: ExcelExporter.exe input.xlsx --outputPath ./output\n",
                    OutputPathHandler
                )
            );
            commands.Add("fileType",
                new Command("fileType",
                    "--fileType; -ft",
                    "The file type to extract the workbook to\nUsage: ExcelExporter.exe <input_file> --fileType <file_type>\nExample: ExcelExporter.exe input.xlsx --fileType xlsm\n",
                    FileTypeHandler
                )
            );

            // TODO: --export_vba, --export_powerquery, --export_tables, --export_ribbons, --export_graphs (maybe as images?)
        }

        public static bool HandleArguments(
            string[] args,
            out string? inputFilePath, out string? fileType, out string? outputDir, out string? outputText
        )
        {
            inputFilePath = args[0];
            fileType = null;
            outputDir = null;
            outputText = null;
            bool success = true;
            
            var invokedHandlers = new HashSet<string>();

            foreach (var cmdEntry in commands)
            {
                string[] aliases = cmdEntry.Value.Aliases.Split(';');
                foreach (var alias in aliases)
                {
                    string trimmedAlias = alias.Trim();
                    int index = Array.IndexOf(args, trimmedAlias);
                    if (index != -1)
                    {
                        if (invokedHandlers.Contains(cmdEntry.Key))
                            continue;

                        invokedHandlers.Add(cmdEntry.Key);

                        bool result = cmdEntry.Value.Handler(args, out string parsedInputFilePath, out string parsedFileType, out string parsedOutputDir, out string handlerOutput);
                        if (!result)
                        {
                            outputText += handlerOutput + "\n";
                            success = false;
                        }

                        if (!string.IsNullOrWhiteSpace(parsedInputFilePath))
                            inputFilePath = parsedInputFilePath;

                        if (!string.IsNullOrWhiteSpace(parsedFileType))
                            fileType = parsedFileType;

                        if (!string.IsNullOrWhiteSpace(parsedOutputDir))
                            outputDir = parsedOutputDir;
                    }
                }
            }

            // Auto-derive fileType from inputFilePath if missing
            if (string.IsNullOrWhiteSpace(fileType) && !string.IsNullOrWhiteSpace(inputFilePath))
            {
                fileType = Path.GetExtension(inputFilePath).TrimStart('.');
                try
                {
                    ValidateFileType(fileType);
                    outputText += $"Auto-detected file type from input file: {fileType}\n";
                }
                catch (Exception ex)
                {
                    outputText += $"Auto-detected file type from input file is invalid: {ex.Message}\n";
                    success = false;
                }
            }

            // Auto-derive outputDir from inputFilePath if missing
            if (string.IsNullOrWhiteSpace(outputDir) && !string.IsNullOrWhiteSpace(inputFilePath))
            {
                outputDir = Path.GetDirectoryName(inputFilePath);
                outputText += $"Auto-detected output directory from input file: {outputDir}\n";
            }

            if (!success || invokedHandlers.Count == 0)
            {
                if (invokedHandlers.Count == 0)
                    outputText += "Error: No valid arguments provided.\n";
            }

            return success;
        }

        #region Command Handlers
        private static bool HelpHandler(
            string[] args,
            out string? parsedInputFilePath, out string? parsedFileType, out string? parsedOutputDir, out string? outputText
        )
        {
            parsedInputFilePath = null;
            parsedFileType = null;
            parsedOutputDir = null;

            outputText = "Usage: ExcelExporter.exe <input_file> <arguments>\n";
            outputText += "Example: ExcelExporter.exe input.xlsx --outputPath ./output\n";
            outputText += "<input_file> should be an Excel file (xlsx, xltx, xlsm, xltm, xlam)\n";
            outputText += "Macro-enabled workbooks (xlsm, xltm, xlam) will have their VBA code extracted\n";
            outputText += "\n";
            outputText += "Available arguments:\n";

            foreach (var command in commands)
            {
                outputText += $"{command.Value.Key}: {command.Value.Description}\n";
                outputText += $"  Aliases: {command.Value.Aliases}\n";
                outputText += "\n";
            }

            return false;
        }

        private static bool VersionHandler(
            string[] args,
            out string? parsedInputFilePath, out string? parsedFileType, out string? parsedOutputDir, out string? outputText
        )
        {
            parsedInputFilePath = null;
            parsedFileType = null;
            parsedOutputDir = null;

            outputText = "Excel Exporter - Command Line Interface\n";
            outputText += "====================================\n";
            outputText += "Version " + File.ReadAllText("version.txt") + "\n";
            return false;
        }

        private static bool OutputPathHandler(
            string[] args,
            out string? parsedInputFilePath, out string? parsedFileType, out string? parsedOutputDir, out string? outputText
        )
        {
            parsedInputFilePath = null;
            parsedFileType = null;
            parsedOutputDir = null;
            outputText = null;

            // Find --outputPath or -op
            int pathIndex = Array.IndexOf(args, "--outputPath");
            if (pathIndex == -1)
            {
                pathIndex = Array.IndexOf(args, "-op");
            }

            // Check if output path is provided
            if (pathIndex == -1 || pathIndex + 1 >= args.Length)
            {
                return false;
            }

            parsedOutputDir = args[pathIndex + 1];
            return true;
        }

        private static bool FileTypeHandler(
            string[] args,
            out string? parsedInputFilePath, out string? parsedFileType, out string? parsedOutputDir, out string? outputText)
        {
            parsedInputFilePath = null;
            parsedFileType = null;
            parsedOutputDir = null;
            outputText = null;

            // Find --fileType or -ft
            int index = Array.IndexOf(args, "--fileType");
            if (index == -1)
                index = Array.IndexOf(args, "-ft");

            // Check if file type is provided
            if (index == -1 || index + 1 >= args.Length)
            {
                outputText = "Error: Missing file type after --fileType or -ft.";
                return false;
            }

            parsedFileType = args[index + 1];

            // Validate file type
            try
            {
                ValidateFileType(parsedFileType);
            }
            catch (Exception ex)
            {
                outputText = $"Invalid file type: {ex.Message}";
                return false;
            }

            return true;
        }
        #endregion

        #region Verifiers
        private static void ValidateFileType(string fileType)
        {
            if(fileType != "xlsx" && fileType != "xltx" && fileType != "xlsm" && fileType != "xltm" && fileType != "xlam")
            {
                throw new Exception("Invalid file type");
            }
        }
        #endregion
    }
}