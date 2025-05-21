public delegate bool Function(
    string[] args,
    out string? parsedInput,
    out string? parsedFileType,
    out string? parsedOutputDir,
    out string? handlerOutput
);

public class Command
{
    public string Key { get; set; }
    public string Aliases { get; set; }
    public string Description { get; set; }
    public Function Handler { get; set; }

    public Command(string key, string aliases, string description, Function handler)
    {
        Key = key;
        Aliases = aliases;
        Description = description;
        Handler = handler;
    }
}