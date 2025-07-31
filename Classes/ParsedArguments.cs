namespace ExcelExporter.Classes
{
    public class ParsedArguments
    {
        private string? input_file_path;
        public string InputFilePath
        {
            set => input_file_path = value;
        }

        private string? file_type;
        public string? FileType
        {
            get => file_type;
            set => file_type = value;
        }

        private string? output_dir;
        public string? OutputDir
        {
            get => output_dir;
            set => output_dir = value;
        }

        private string? output_text;
        public string? OutputText
        {
            get => output_text;
            set => output_text = value;
        }

        private string? output_folder;
        public string? OutputFolder
        {
            get => output_folder;
            set => output_folder = value;
        }
    }
}