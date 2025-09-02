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
            commands.Add("outputDir",
                new Command("outputDir",
                    "--outputDir; -od",
                    "The directory to extract the workbook to\n" +
                    "Usage: ExcelExporter.exe <input_file> --outputDir <output_path>\n" +
                    "Example: ExcelExporter.exe input.xlsm --outputDir ./output\n" +
                    "This will generate a folder structure like so: `./output/input_xlsm/(exported files)`",
                    OutputDirHandler
                )
            );
            commands.Add("outputName",
                new Command("outputName",
                    "--outputName; -on",
                    "The folder to extract the workbook to." +
                    "If outputPath is specified alongside this, this will be generated inside that folder\n" +
                    "Usage: ExcelExporter.exe <input_file> --outputName <output_folder>\n" +
                    "Example: ExcelExporter.exe input.xlsm --outputName ./output\n" +
                    "This will generate a folder structure like so: `./path/to/input/file/output_xlsm/(exported files)`",
                    OutputNameHandler
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
                    "Example: ExcelExporter.exe input.xlsx --fileType xlsm",
                    FileTypeHandler
                )
            );
            commands.Add("saveType",
                new Command("saveType",
                    "--saveType; -st",
                    "The filetype to save the extraction as\n" +
                    "Usage: ExcelExporter.exe <input_file> --saveType <save_type>\n" +
                    "Options: 1. 'csv'; 2. 'json'\n" +
                    "Example: ExcelExporter.exe input.xlsm --saveType json",
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

            foreach (KeyValuePair<string, Command> cmdEntry in commands)
            {
                bool result = cmdEntry.Value.Handler(args, parsed_args, out string handlerOutput);
                if (!result)
                {
                    message += handlerOutput + "\n";
                    success = false;
                }
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
            out string message
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
            out string message
        )
        {
            if(args.Contains("--version") || args.Contains("-v"))
            {
                message = $"Version {Assembly.GetExecutingAssembly().GetName().Version}\n";
                return false;
            }

            message = "";
            return true;
        }

        /// <summary>
        /// Parses the <c>--outputDir</c> or <c>-od</c> argument from the command-line
        /// and sets the output directory in <paramref name="parsedArguments"/>.
        /// </summary>
        /// <param name="args">The full array of command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with parsed values.</param>
        /// <param name="message">An optional error or info message to return.</param>
        /// <returns>
        /// <c>true</c> if the output path was successfully found and set; 
        /// <c>false</c> if the argument was missing or incomplete.
        /// </returns>
        private static bool OutputDirHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string message
        )
        {
            message = "";

            // Find --outputDir or -od
            int pathIndex = Array.IndexOf(args, "--outputDir");
            if (pathIndex == -1)
                pathIndex = Array.IndexOf(args, "-od");
            if(pathIndex == -1)
            {
                parsedArguments.OutputDir = Path.GetDirectoryName(args[0]);
                return true;
            }

            // Check if output path is provided
            if (pathIndex + 1 >= args.Length)
            {
                message = "Error: Missing output path after --outputDir or -od.";
                return false;
            }

            parsedArguments.OutputDir = args[pathIndex + 1];
            return true;
        }

        /// <summary>
        /// Parses the <c>--outputName</c> or <c>-on</c> argument from the command-line arguments.
        /// If not provided, defaults the output folder to the directory of the input file.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        /// <param name="parsedArguments">The object that will store the parsed output folder path.</param>
        /// <param name="message">Unused in this handler; set to an empty string.</param>
        /// <returns>
        /// <c>true</c> if the output folder was found or defaulted successfully; <c>false</c> if the argument
        /// was specified but the folder path was missing.
        /// </returns>
        private static bool OutputNameHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string message
        )
        {
            message = "";

            // Find --outputName or -on
            int pathIndex = Array.IndexOf(args, "--outputName");
            if (pathIndex == -1)
                pathIndex = Array.IndexOf(args, "-on");
            if(pathIndex == -1)
            {
                parsedArguments.OutputName = Path.GetFileName(args[0]);
                return true;
            }

            // Check if output folder is provided
            if (pathIndex + 1 >= args.Length)
                return false;

            parsedArguments.OutputName = args[pathIndex + 1];
            return true;
        }

        /// <summary>
        /// Parses the <c>--fileType</c> or <c>-ft</c> argument from the command-line input.
        /// If not provided, attempts to auto-detect the file type from the input file's extension.
        /// Validates the file type using <c>ValidateFileType</c>.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with the detected or provided file type.</param>
        /// <param name="message">
        /// An informational or error message indicating the result of the parsing and validation.
        /// Includes auto-detection feedback or missing/invalid argument errors.
        /// </param>
        /// <returns>
        /// <c>true</c> if a valid file type is detected or provided; <c>false</c> if validation fails
        /// or the required value is missing when explicitly specified.
        /// </returns>
        private static bool FileTypeHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string message
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
                file_type = HandleIncorrectFileType(args[0]);
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
                file_type = HandleIncorrectFileType(args[0]);
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

        private static string HandleIncorrectFileType(string inputFilePath)
        {
            string outputStr = inputFilePath;

            if (outputStr.Contains('.'))
                outputStr = outputStr.Split('.').Last();

            if (outputStr.Contains('\''))
                outputStr = outputStr.TrimEnd('\'');

            return outputStr;
        }

        /// <summary>
        /// Parses the <c>--saveType</c> or <c>-st</c> argument from the command-line input.
        /// If not provided, defaults to <c>"json"</c> as the save type.
        /// Validates the save type using <c>ValidateSaveType</c>.
        /// </summary>
        /// <param name="args">The command-line arguments.</param>
        /// <param name="parsedArguments">The object to populate with the detected or provided save type.</param>
        /// <param name="message">
        /// A message indicating whether a default was used, or describing any validation error encountered.
        /// </param>
        /// <returns>
        /// <c>true</c> if a valid save type is provided or defaulted; <c>false</c> if the argument is missing a value
        /// or specifies an invalid type.
        /// </returns>
        private static bool SaveTypeHandler(
            string[] args,
            ParsedArguments parsedArguments,
            out string message
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