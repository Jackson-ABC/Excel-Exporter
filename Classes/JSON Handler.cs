using System.Text;

namespace ExcelExporter.Classes
{
    internal class JSON_Handler
    {
        public static void Write(object[,] array, string outputDir, string filename)
        {
            string jsonPath = Path.Combine(outputDir, filename + ".json");

            var sb = new StringBuilder();
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);
        }
    }
}
