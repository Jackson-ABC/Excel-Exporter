namespace ExcelExporter.Classes
{
    public delegate bool Function(
        string[] args,
        ParsedArguments parsedArguments,
        out string handlerOutput
    );

    public class Command(
        string key,
        string aliases,
        string description,
        Function handler
    )
    {
        public string Key { get; set; } = key;
        public string Aliases { get; set; } = aliases;
        public string Description { get; set; } = description;
        public Function Handler { get; set; } = handler;
    }
}