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

            /*
             * {
             *   "name":"document name",
             *   "sheets":{
             *     "sheetName":{
             *       "lastCol":10,
             *       "lastRow":10,
             *       "values":{
             *         "A":{
             *           "1":{
             *             "displayedVal":"",
             *             "formulaVal":"",
             *             "?other vals?":""
             *           },
             *           "2":{
             *             "displayedVal":"",
             *             "formulaVal":"=IF(TRUE, \"\", \"false displayed\")"
             *           }
             *         },
             *         "B":{
             *           "1":{
             *             "displayedVal":"",
             *             "formulaVal":"",
             *           }
             *         }
             *       }
             *     },
             *     "customRanges":{
             *       "rangeName":"Sheet1!A1:C3"
             *     }
             *   }
             * }
             */

        }
    }
}
