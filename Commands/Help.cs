using ExcelExporter.Classes;

namespace ExcelExporter.Commands
{
    public class Help
    {
        string key = "help";
        string longAlias = "--help";
        string shortAlias = "-h";
        string description = "Display this help message";
        Function handler = Handler;
        
        /// <summary>
        /// Returns the command object for this command.
        /// </summary>
        /// <returns>The command object for this command.</returns>
        public Command Command()
        {
            return new Command(key,
                $"{longAlias}; {shortAlias}",
                description,
                handler
            );
        }

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
        public static bool Handler(
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

            /*if(args.Length > 1)
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
            }*/

            return false;
        }
    }
}