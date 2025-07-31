using ExcelExporter.Resources;
using System.Reflection;

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
                    "The path to extract the workbook to\n" +
                    "Usage: ExcelExporter.exe <input_file> --outputPath <output_path>\n" +
                    "Example: ExcelExporter.exe input.xlsx --outputPath ./output\n",
                    OutputPathHandler
                )
            );
            commands.Add("fileType",
                new Command("fileType",
                    "--fileType; -ft",
                    "How to handle the workbook." +
                    "If an xlsm is handled as an xlsx, the VBA code will not be exported." +
                    "If an xlsx is handled as an xlsm, a vba folder will be generated.\n" +
                    "Usage: ExcelExporter.exe <input_file> --fileType <file_type>\n" +
                    "Standard Options: 'xlsm', 'xlsx'\n" +
                    "Example: ExcelExporter.exe input.xlsx --fileType xlsm\n",
                    FileTypeHandler
                )
            );
            commands.Add("saveType",
                new Command("saveType",
                    "--saveType; -st",
                    "The filetype to save the extraction as\n" +
                    "Usage: ExcelExporter.exe <input_file> --saveType <save_type>\n" +
                    "Options: 1. 'csv'; 2. 'json'\n" +
                    "Example: ExcelExporter.exe input.xlsm --saveType json\n",
                    SaveTypeHandler
                )
            );

            // TODO: --export_vba, --export_powerquery, --export_tables, --export_ribbons, --export_graphs (maybe as images?)
        }

        public static bool HandleArguments(
            string[] args,
            out ParsedArguments parsed_args,
            out string message
        )
        {
            bool success = true;
            string inputFilePath = args[0];
            message = "";
            parsed_args = new ParsedArguments();

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

                        bool result = cmdEntry.Value.Handler(args, parsed_args, out string handlerOutput);
                        if (!result)
                        {
                            message += handlerOutput + "\n";
                            success = false;
                        }
                    }
                }
            }

            if (!success || invokedHandlers.Count == 0)
            {
                if (invokedHandlers.Count == 0)
                    message += "Error: No valid arguments provided.\n";
            }

            return success;
        }

        #region Command Handlers
        /// <summary>
        /// Handles the <c>--help</c> or <c>-h</c> argument by generating usage instructions 
        /// and listing all available commands and their aliases.
        /// </summary>
        /// <param name="args">The full array of command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with parsed values (unused).</param>
        /// <param name="message">The message containing usage information and available commands.</param>
        /// <returns>
        /// Always returns <c>false</c> to indicate help was shown and normal processing should stop.
        /// </returns>
        private static bool HelpHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string? message
        )
        {
            message =
                "Usage: ExcelExporter.exe <input_file> <arguments>\n" +
                "Example: ExcelExporter.exe input.xlsx\n" +
                "<input_file> should be an Excel file (xlsx, xltx, xlsm, xltm, xlam)\n" +
                "Macro-enabled workbooks (xlsm, xltm, xlam) will have their VBA code extracted\n" +
                "\n" +
                "Available arguments:\n";

            if(args.Length > 1)
            {
                foreach (string arg in args)
                {
                    message += $"{arg}: {commands[arg].Description}\n" +
                        $"  Aliases: {commands[arg].Aliases}\n" +
                        "\n";
                }
            }
            else
            {
                foreach (var command in commands)
                {
                    message +=
                        $"{command.Value.Key}: {command.Value.Description}\n" +
                        $"  Aliases: {command.Value.Aliases}\n" +
                        "\n";
                }
            }

            return false;
        }

        /// <summary>
        /// Handles the <c>--version</c> argument by reading the version number from a file
        /// and returning it in the <paramref name="message"/>.
        /// </summary>
        /// <param name="args">The full array of command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with parsed values (unused).</param>
        /// <param name="message">The message containing version information.</param>
        /// <returns>
        /// Always returns <c>false</c> to indicate no further argument processing is required.
        /// </returns>
        private static bool VersionHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string? message
        )
        {
            message = $"Version {Assembly.GetExecutingAssembly().GetName().Version}\n";
            return false;
        }

        /// <summary>
        /// Parses the <c>--outputPath</c> or <c>-op</c> argument from the command-line
        /// and sets the output directory in <paramref name="parsedArguments"/>.
        /// </summary>
        /// <param name="args">The full array of command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with parsed values.</param>
        /// <param name="message">An optional error or info message to return.</param>
        /// <returns>
        /// <c>true</c> if the output path was successfully found and set; 
        /// <c>false</c> if the argument was missing or incomplete.
        /// </returns>
        private static bool OutputPathHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string? message
        )
        {
            message = "";

            // Find --outputPath or -op
            int pathIndex = Array.IndexOf(args, "--outputPath");
            if (pathIndex == -1)
                pathIndex = Array.IndexOf(args, "-op");
            if(pathIndex == -1)
            {
                parsedArguments.OutputPath = Path.GetDirectoryName(args[0]);
                return true;
            }

            // Check if output path is provided
            if (pathIndex + 1 >= args.Length)
            {
                message = "Error: Missing output path after --outputPath or -op.";
                return false;
            }

            parsedArguments.OutputPath = args[pathIndex + 1];
            return true;
        }

        private static bool FileTypeHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string? message
        )
        {
            message = "";
            string file_type;

            // Find --fileType or -ft
            int index = Array.IndexOf(args, "--fileType");
            if (index == -1)
                index = Array.IndexOf(args, "-ft");
            if(index == -1)
            {
                file_type = "";
                if (ValidateFileType(file_type))
                    message += $"Auto-detected file type from input file: {file_type}\n";
                else
                {
                    message += $"Auto-detected file type from input file is invalid: {file_type}\n";
                    return false;
                }
                goto end;
            }

            // Check if file type is provided
            if (index + 1 >= args.Length)
            {
                message = "Error: Missing file type after --fileType or -ft.";
                return false;
            }

            // Check if file type is valid
            file_type = args[index + 1];
            
            if (!ValidateFileType(file_type))
            {
                file_type = Path.GetExtension(args[0]).TrimStart('.');
                if (ValidateFileType(file_type))
                    message += $"Auto-detected file type from input file: {file_type}\n";
                else
                {
                    message += $"Auto-detected file type from input file is invalid: {file_type}\n";
                    return false;
                }
            }

            end:
            parsedArguments.FileType = file_type;
            return true;
        }

        private static bool SaveTypeHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string? message
        )
        {
            message = "";
            string save_type;

            // Find --saveType or -st
            int pathIndex = Array.IndexOf(args, "--saveType");
            if (pathIndex == -1)
                pathIndex = Array.IndexOf(args, "-st");

            if(pathIndex == -1)
            {
                save_type = "json";
                message += $"Standard save type used: {save_type}\n";
                return true;
            }

            // Check if save type is provided
            if (pathIndex + 1 >= args.Length)
            {
                message += "Error: Missing save type after --saveType or -st.";
                return false;
            }

            // Check if save type is valid
            if (!ValidateSaveType(save_type = args[pathIndex + 1]))
            {
                message += "Error: Invalid save type after --saveType or -st.";
                return false;
            }
            parsedArguments.SaveType = save_type;
            return true;
        }
        #endregion

        #region Verifiers
        private static bool ValidateFileType(string fileType)
        {
            if (GlobalDictionaries.AllowedFileTypes.Contains(fileType))
                return true;
            else
                return false;
        }

        private static bool ValidateSaveType(string saveType)
        {
            if (GlobalDictionaries.AllowedSaveTypes.Contains(saveType))
                return true;
            else
                return false;
        }
        #endregion
    }
}