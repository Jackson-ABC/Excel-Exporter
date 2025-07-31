namespace ExcelExporter.Resources
{
    static class GlobalDictionaries
    {
        public static HashSet<string> AllowedFileTypes = new HashSet<string>
        {
            "xlsx",
            "xltx",
            "xlsm",
            "xltm",
            "xlam"
        };

        public static HashSet<string> AllowedSaveTypes = new HashSet<string>
        {
            "csv",
            "json"
        };
    }
}