using System.IO;
using System.Text;

public class CSVHandler
{
    /// <summary>
    /// Writes a 2D array to a CSV file
    /// </summary>
    /// <param name="array">The 2D array to save</param>
    /// <param name="outputDir">The directory to save the file to</param>
    /// <param name="fileName">The name of the file to save</param>
    public static void WriteToCSV(object[,] array, string outputDir, string fileName)
    {
        string csvPath = Path.Combine(outputDir, fileName + ".csv");
        
        var sb = new StringBuilder();
        int rows = array.GetLength(0);
        int cols = array.GetLength(1);

        for (int i = 1; i <= rows; i++)  // Excel interop is 1-based
        {
            for (int j = 1; j <= cols; j++)
            {
                object cell = array[i, j];
                string value = cell?.ToString() ?? "";

                // Quote if it contains comma, quote, or newline
                if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
                {
                    value = "\"" + value.Replace("\"", "\"\"") + "\"";
                }

                sb.Append(value);

                if (j < cols)
                    sb.Append(",");
            }
            sb.AppendLine(); // new row
        }

        File.WriteAllText(csvPath, sb.ToString());
    }

}